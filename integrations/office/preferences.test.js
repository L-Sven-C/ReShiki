import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, writeFile, rm, symlink } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import http from "node:http";
import {
  createPreferences,
  normalizePreferences,
  translate,
} from "./preferences.js";
import { preferenceDOM } from "./test-support/preferences-dom.js";
import { HOSTS, manifest } from "./setup.js";
import { ribbonLanguageManager } from "./companion/ribbon-language.js";
import { makeHandler } from "./companion/server.js";

function ui(storage, api = async () => ({ managed: true, language: "en-US" })) {
  const nodes = new Map([["file", {}]]);
  const doc = {
    ...preferenceDOM(nodes),
    getElementById: (id) => nodes.get(id),
    querySelectorAll: () => [],
  };
  return { ...createPreferences({ document: doc, storage, api }), nodes, doc };
}
test("preferences default to English, persist only validated UI values and apply without reloading", async () => {
  let saved;
  const storage = {
    getItem: () => saved,
    setItem: (_key, value) => {
      saved = value;
    },
  };
  const p = ui(storage);
  assert.equal(p.doc.documentElement.lang, "en-US");
  const status = { textContent: "" };
  p.message(
    status,
    "Drawing updated in Office. Save your Office file to keep it.",
  );
  p.nodes.get("ui-language").value = "zh-CN";
  p.nodes.get("compact-layout").checked = false;
  p.nodes.get("show-hints").checked = true;
  p.nodes.get("ui-language").listeners.change();
  assert.match(status.textContent, /已更新 Office/);
  assert.equal(p.nodes.get("usage-hints").hidden, false);
  assert.equal(p.doc.body.dataset.compact, "false");
  const reopened = ui(storage);
  assert.equal(reopened.doc.documentElement.lang, "zh-CN");
  assert.deepEqual(JSON.parse(saved), {
    language: "zh-CN",
    compact: false,
    hints: true,
  });
  assert.deepEqual(
    normalizePreferences({
      language: "bad",
      compact: "false",
      token: "do not store",
    }),
    { language: "en-US", compact: true, hints: false },
  );
  assert.equal(
    translate("Unknown diagnostic <x>", "zh-CN"),
    "Unknown diagnostic <x>",
  );
});
test("unavailable storage and malformed preferences never block editing UI", () => {
  for (const storage of [
    undefined,
    { getItem: () => "{broken" },
    {
      getItem() {
        throw new Error("denied");
      },
    },
  ]) {
    const p = ui(storage);
    p.nodes.get("ui-language").value = "zh-CN";
    p.nodes.get("ui-language").listeners.change();
    assert.equal(p.doc.documentElement.lang, "zh-CN");
    assert.match(p.nodes.get("preferences-status").textContent, /仅在本次会话/);
    p.showDrawings();
    assert.equal(p.nodes.get("workspace").hidden, false);
  }
});
test("preferences and ribbon application do not mutate the document or lose error feedback", async () => {
  const calls = [];
  const p = ui(undefined, async (route, body) => {
    calls.push({ route, body });
    if (body) throw new Error("read only");
    return { managed: true, language: "zh-CN" };
  });
  await p.showPreferences();
  assert.equal(p.nodes.get("workspace").hidden, true);
  assert.equal(p.nodes.get("ribbon-language").value, "zh-CN");
  await p.nodes.get("apply-ribbon").listeners.click();
  assert.equal(p.nodes.get("ribbon-status").textContent, "read only");
  assert.equal(p.nodes.get("apply-ribbon").disabled, false);
  assert.deepEqual(calls, [
    { route: "ribbon-language", body: undefined },
    { route: "ribbon-language", body: { language: "zh-CN" } },
  ]);
});
async function fixture(t) {
  const root = await mkdtemp(path.join(os.tmpdir(), "reshiki-language-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  const file = path.join(root, "word.xml");
  await writeFile(file, manifest(HOSTS[0]));
  return { root, file, manage: ribbonLanguageManager(file, 43127) };
}
test("ribbon language survives manager restart; writes are serialized, same ID/runtime and only fixed language", async (t) => {
  const { file, manage } = await fixture(t);
  assert.equal((await manage()).language, "en-US");
  await manage("zh-CN");
  assert.equal(
    await readFile(file, "utf8"),
    manifest(HOSTS[0], 43127, "zh-CN"),
  );
  assert.equal((await ribbonLanguageManager(file, 43127)()).language, "zh-CN");
  await Promise.all([manage("en-US"), manage("zh-CN"), manage("en-US")]);
  assert.equal(await readFile(file, "utf8"), manifest(HOSTS[0]));
  await assert.rejects(manage("../../evil"), /Unsupported/);
  assert.equal(await readFile(file, "utf8"), manifest(HOSTS[0]));
  assert.match(manifest(HOSTS[0]), /DefaultValue="Preferences"/);
  assert.match(manifest(HOSTS[0], 43127, "zh-CN"), /DefaultValue="偏好设置"/);
});
test("language manager refuses unconfigured paths, custom manifests and symlinks without clobbering", async (t) => {
  const { root, file, manage } = await fixture(t);
  assert.deepEqual(await ribbonLanguageManager()(), { managed: false });
  await assert.rejects(ribbonLanguageManager()("zh-CN"), /not configured/);
  assert.throws(() => ribbonLanguageManager("relative.xml", 43127), /absolute/);
  await writeFile(file, "custom installation");
  await assert.rejects(manage("zh-CN"), /custom changes/);
  const link = path.join(root, "link.xml");
  await symlink(file, link);
  await assert.rejects(
    ribbonLanguageManager(link, 43127)("zh-CN"),
    /regular file/,
  );
  assert.equal(await readFile(file, "utf8"), "custom installation");
});
test("ribbon language API requires authenticated Word client and rejects request-supplied paths", async (t) => {
  const { manage } = await fixture(t);
  let handler;
  const server = http.createServer((req, res) => handler(req, res));
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  t.after(() => {
    server.closeAllConnections();
    server.close();
  });
  const origin = `http://127.0.0.1:${server.address().port}`;
  handler = makeHandler({
    origin,
    token: "test",
    sessions: {
      touch(id) {
        if (!["word", "excel"].includes(id)) throw new Error("Invalid client");
        return { host: id === "word" ? "Word" : "Excel" };
      },
    },
    ribbonLanguage: manage,
  });
  const send = (body, headers = {}) =>
    fetch(origin + "/api/ribbon-language", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Origin: origin,
        "X-ReShiki-Token": "test",
        ...headers,
      },
      body: JSON.stringify(body),
    });
  assert.equal(
    (
      await send(
        { clientId: "word", language: "zh-CN" },
        { Origin: "https://evil.invalid" },
      )
    ).status,
    403,
  );
  assert.equal(
    (await send({ clientId: "word" }, { "X-ReShiki-Token": "bad" })).status,
    403,
  );
  assert.equal(
    (await send({ clientId: "excel", language: "zh-CN" })).status,
    403,
  );
  assert.equal(
    (
      await send({
        clientId: "word",
        language: "zh-CN",
        path: "/tmp/other.xml",
      })
    ).status,
    400,
  );
  assert.equal((await manage()).language, "en-US");
  assert.equal(
    (await send({ clientId: "word", language: "zh-CN" })).status,
    200,
  );
  assert.equal((await manage()).language, "zh-CN");
  for (const name of ["insert", "edit", "pane", "preferences"]) {
    for (const size of [16, 32, 80]) {
      const response = await fetch(`${origin}/assets/${name}-${size}.png`);
      assert.equal(response.status, 200);
      assert.equal(response.headers.get("content-type"), "image/png");
      const bytes = Buffer.from(await response.arrayBuffer());
      assert.equal(bytes.readUInt32BE(16), size);
      assert.equal(bytes.readUInt32BE(20), size);
    }
  }
  assert.equal((await fetch(origin + "/preferences.js")).status, 200);
  assert.equal((await fetch(origin + "/preferences.js?_host_Info=Word")).status, 403);
  assert.equal((await fetch(origin + "/assets/arbitrary-32.png")).status, 404);
});
