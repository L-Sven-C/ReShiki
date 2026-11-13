import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { DOMParser } from "@xmldom/xmldom";
import { registerRibbon } from "./ribbon.js";
import { HOSTS, manifest } from "./setup.js";

test("ribbon commands wait for initialization and showing the shared pane, then complete once", async () => {
  const handlers = new Map(),
    calls = [];
  let initialize, show;
  const ready = new Promise((resolve) => {
    initialize = resolve;
  });
  const visible = new Promise((resolve) => {
    show = resolve;
  });
  registerRibbon({
    Office: {
      actions: { associate: (name, fn) => handlers.set(name, fn) },
      addin: {
        showAsTaskpane: () => {
          calls.push("show");
          return visible;
        },
      },
    },
    ready,
    insert: () => calls.push("insert"),
    edit: () => calls.push("edit"),
    report: (error) => assert.fail(error),
  });
  const running = handlers.get("reshikiEdit")({
    completed: () => calls.push("completed"),
  });
  assert.deepEqual(calls, []);
  initialize();
  await new Promise((resolve) => setImmediate(resolve));
  assert.deepEqual(calls, ["show"]);
  show();
  await running;
  assert.deepEqual(calls, ["show", "edit", "completed"]);
  await handlers.get("reshikiInsert")({
    completed: () => calls.push("completed"),
  });
  assert.deepEqual(calls.slice(3), ["show", "insert", "completed"]);
});

test("ribbon failures report and complete without performing a later action", async () => {
  for (const stage of ["initialize", "show", "edit"]) {
    const handlers = new Map(),
      calls = [];
    registerRibbon({
      Office: {
        actions: { associate: (name, fn) => handlers.set(name, fn) },
        addin: {
          showAsTaskpane: async () => {
            calls.push("show");
            if (stage === "show") throw new Error(stage);
          },
        },
      },
      ready:
        stage === "initialize"
          ? Promise.reject(new Error(stage))
          : Promise.resolve(),
      insert: () => assert.fail("unexpected insert"),
      edit: () => {
        calls.push("edit");
        throw new Error(stage);
      },
      report: (error) => calls.push(error + " error"),
    });
    await handlers.get("reshikiEdit")({
      completed: () => calls.push("completed"),
    });
    assert.equal(calls.filter((call) => call === "completed").length, 1);
    assert.deepEqual(calls.slice(-2), [stage + " error", "completed"]);
    if (stage !== "edit") assert.ok(!calls.includes("edit"));
  }
});

test("legacy panes without command APIs retain their normal entry points", () => {
  assert.doesNotThrow(() => registerRibbon({ Office: {} }));
});

test("Word ribbon has four distinct controls sharing one loopback runtime at custom ports", () => {
  for (const port of [43127, 44444]) {
    const doc = new DOMParser().parseFromString(
      manifest(HOSTS[0], port),
      "text/xml",
    );
    const elements = (name) => Array.from(doc.getElementsByTagName(name));
    assert.equal(elements("CustomTab").length, 1);
    assert.deepEqual(
      elements("Control").map((x) => x.getAttribute("id")),
      ["ReShiki.Insert", "ReShiki.Edit", "ReShiki.Open", "ReShiki.Preferences"],
    );
    assert.deepEqual(
      elements("FunctionName").map((x) => x.textContent),
      ["reshikiInsert", "reshikiEdit", "reshikiOpen", "reshikiPreferences"],
    );
    assert.equal(elements("Runtime")[0].getAttribute("lifetime"), "long");
    const runtime = elements("Runtime")[0].getAttribute("resid");
    assert.equal(elements("FunctionFile")[0].getAttribute("resid"), runtime);
    assert.equal(elements("Group").length, 2);
    assert.ok(
      elements("Action").every(
        (x) => x.getAttribute("xsi:type") === "ExecuteFunction",
      ),
    );
    const icons = elements("Control").map((x) =>
      x.getElementsByTagName("bt:Image")[0].getAttribute("resid"),
    );
    assert.equal(new Set(icons).size, 4);
    assert.equal(elements("TaskpaneId").length, 0);
    const urls = elements("bt:Url");
    assert.equal(urls.length, 1);
    assert.equal(urls[0].getAttribute("id"), runtime);
    assert.equal(
      urls[0].getAttribute("DefaultValue"),
      `https://localhost:${port}/taskpane.html`,
    );
    assert.ok(
      elements("Set").some((x) => x.getAttribute("Name") === "SharedRuntime"),
    );
  }
  for (const host of HOSTS.slice(1)) {
    assert.ok(!manifest(host).includes("VersionOverrides"));
    assert.ok(!manifest(host).includes("SharedRuntime"));
  }
});

test("checked-in manifests match the setup generator", async () => {
  for (const host of HOSTS)
    assert.equal(
      await readFile(
        new URL(`./manifests/${host.name}.xml`, import.meta.url),
        "utf8",
      ),
      manifest(host),
    );
});
