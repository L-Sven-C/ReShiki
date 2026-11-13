#!/usr/bin/env node
// Generates configuration/manifests only. It never trusts a certificate, writes
// Office's sideload directories, changes a registry, or starts an Office app.
import {
  readFile,
  writeFile,
  mkdir,
  stat,
  lstat,
  link,
  unlink,
} from "node:fs/promises";
import {
  X509Certificate,
  createPrivateKey,
  createPublicKey,
  randomUUID,
} from "node:crypto";
import path from "node:path";
import os from "node:os";
import { fileURLToPath } from "node:url";

export const HOSTS = [
  {
    name: "word",
    host: "Document",
    api: "WordApi",
    version: "1.4",
    id: "ea615cbd-ce62-4492-a8ea-cd5b426bd02e",
  },
  {
    name: "excel",
    host: "Workbook",
    api: "ExcelApi",
    version: "1.19",
    id: "118ea7ec-c8ea-4a11-a1f3-3c6a8f08d9dd",
  },
  {
    name: "powerpoint",
    host: "Presentation",
    api: "PowerPointApi",
    version: "1.8",
    id: "59b424f2-7de1-4c81-aa8b-f9273f0d06b4",
  },
];
function wordCommands(origin, language) {
  const zh = language === "zh-CN";
  const buttons = [
    [
      "Insert",
      zh ? "插入结构" : "Insert Structure",
      zh
        ? "从文件或剪贴板插入可编辑结构。"
        : "Insert an editable drawing from a file or the clipboard.",
      "reshikiInsert",
      "insert",
    ],
    [
      "Edit",
      zh ? "编辑结构" : "Edit Structure",
      zh
        ? "在 ReShiki 中编辑选中结构。"
        : "Edit the selected drawing in ReShiki.",
      "reshikiEdit",
      "edit",
    ],
    [
      "Open",
      zh ? "打开窗格" : "Open Pane",
      zh ? "查看结构和编辑会话。" : "Show drawings and current edit sessions.",
      "reshikiOpen",
      "pane",
    ],
    [
      "Preferences",
      zh ? "偏好设置" : "Preferences",
      zh
        ? "调整语言、布局和操作提示。"
        : "Choose languages, layout and guidance.",
      "reshikiPreferences",
      "preferences",
    ],
  ];
  const icon = (name) =>
    `<Icon>${[16, 32, 80].map((size) => `<bt:Image size="${size}" resid="ReShiki.${name}${size}"/>`).join("")}</Icon>`;
  const control = ([id, , , fn, asset]) =>
    `<Control xsi:type="Button" id="ReShiki.${id}"><Label resid="ReShiki.${id}.Label"/><Supertip><Title resid="ReShiki.${id}.Label"/><Description resid="ReShiki.${id}.Help"/></Supertip>${icon(asset)}<Action xsi:type="ExecuteFunction"><FunctionName>${fn}</FunctionName></Action></Control>`;
  return `
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xsi:type="VersionOverridesV1_0">
    <Requirements><bt:Sets DefaultMinVersion="1.1"><bt:Set Name="WordApi" MinVersion="1.4"/><bt:Set Name="SharedRuntime" MinVersion="1.1"/></bt:Sets></Requirements>
    <Hosts><Host xsi:type="Document"><Runtimes><Runtime resid="ReShiki.Page" lifetime="long"/></Runtimes><DesktopFormFactor>
      <FunctionFile resid="ReShiki.Page"/>
      <ExtensionPoint xsi:type="PrimaryCommandSurface"><CustomTab id="ReShiki.Tab">
        <Group id="ReShiki.Drawings"><Label resid="ReShiki.Group"/>${icon("insert")}${buttons.slice(0, 2).map(control).join("")}</Group>
        <Group id="ReShiki.Tools"><Label resid="ReShiki.Tools.Label"/>${icon("preferences")}${buttons.slice(2).map(control).join("")}</Group>
        <Label resid="ReShiki.Tab.Label"/>
      </CustomTab></ExtensionPoint>
    </DesktopFormFactor></Host></Hosts>
    <Resources>
      <bt:Images>${buttons.flatMap(([, , , , name]) => [16, 32, 80].map((size) => `<bt:Image id="ReShiki.${name}${size}" DefaultValue="${origin}/assets/${name}-${size}.png"/>`)).join("")}</bt:Images>
      <bt:Urls><bt:Url id="ReShiki.Page" DefaultValue="${origin}/taskpane.html"/></bt:Urls>
      <bt:ShortStrings><bt:String id="ReShiki.Tab.Label" DefaultValue="ReShiki"/><bt:String id="ReShiki.Group" DefaultValue="${zh ? "结构" : "Drawings"}"/><bt:String id="ReShiki.Tools.Label" DefaultValue="${zh ? "工具" : "Tools"}"/>${buttons.map(([id, label]) => `<bt:String id="ReShiki.${id}.Label" DefaultValue="${label}"/>`).join("")}</bt:ShortStrings>
      <bt:LongStrings>${buttons.map(([id, , help]) => `<bt:String id="ReShiki.${id}.Help" DefaultValue="${help}"/>`).join("")}</bt:LongStrings>
    </Resources>
  </VersionOverrides>`;
}
export function manifest(host, port = 43127, language = "en-US") {
  if (!Number.isInteger(port) || port < 1024 || port > 65535)
    throw new Error("Invalid localhost port");
  if (!["en-US", "zh-CN"].includes(language))
    throw new Error("Unsupported ribbon language");
  const origin = `https://localhost:${port}`;
  return `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <Id>${host.id}</Id><Version>${host.name === "word" ? "1.1.0.0" : "1.0.0.0"}</Version><ProviderName>ReShiki</ProviderName><DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="ReShiki"/><Description DefaultValue="Insert and edit native ReShiki molecular drawings with the local companion."/>
  <IconUrl DefaultValue="${origin}/assets/icon-32.png"/><HighResolutionIconUrl DefaultValue="${origin}/assets/icon-80.png"/>
  <SupportUrl DefaultValue="https://github.com/Ameyanagi/ReShiki/issues"/>
  <AppDomains><AppDomain>${origin}</AppDomain></AppDomains>
  <Hosts><Host Name="${host.host}"/></Hosts>
  <Requirements><Sets DefaultMinVersion="1.1"><Set Name="${host.api}" MinVersion="${host.version}"/>${host.name === "word" ? '<Set Name="SharedRuntime" MinVersion="1.1"/>' : ""}</Sets></Requirements>
  <DefaultSettings><SourceLocation DefaultValue="${origin}/taskpane.html"/></DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>${host.name === "word" ? wordCommands(origin, language) : ""}
</OfficeApp>
`;
}
// Publish complete files without overwriting any destination. Identical files
// from a previous or interrupted run can be reused on the next invocation.
async function publish(file, contents) {
  const temporary = path.join(path.dirname(file), `.${path.basename(file)}.${randomUUID()}.tmp`);
  try {
    await writeFile(temporary, contents, { flag: "wx", mode: 0o600, flush: true });
    try {
      await link(temporary, file);
    } catch (error) {
      if (error.code !== "EEXIST") throw error;
      const existing = await lstat(file);
      if (
        !existing.isFile() ||
        existing.isSymbolicLink() ||
        existing.size !== Buffer.byteLength(contents) ||
        (await readFile(file, "utf8")) !== contents
      )
        throw error;
    }
  } finally {
    await unlink(temporary).catch((error) => {
      if (error.code !== "ENOENT") throw error;
    });
  }
}

export async function setup({ executable, cert, key, directory, port = 43127 }) {
  for (const [label, value] of Object.entries({ executable, cert, key, directory }))
    if (typeof value !== "string" || !path.isAbsolute(value))
      throw new Error(`${label} must be an absolute path`);
  if (!(await stat(executable)).isFile()) throw new Error("ReShiki executable must be a file");
  const certificate = new X509Certificate(await readFile(cert));
  if (!certificate.checkHost("localhost", { subject: "never" }))
    throw new Error("Certificate must include localhost in its subject alternative names");
  if (
    Date.parse(certificate.validTo) <= Date.now() ||
    Date.parse(certificate.validFrom) > Date.now()
  )
    throw new Error("Certificate is not currently valid");
  const privateKey = createPrivateKey(await readFile(key));
  if (
    !certificate.publicKey
      .export({ type: "spki", format: "der" })
      .equals(createPublicKey(privateKey).export({ type: "spki", format: "der" }))
  )
    throw new Error("Certificate and private key do not match");
  const rendered = HOSTS.map((host) => [host.name, manifest(host, port)]);
  await mkdir(path.join(directory, "manifests"), { recursive: true, mode: 0o700 });
  await mkdir(path.join(directory, "recovery"), { recursive: true, mode: 0o700 });
  // The config is the installation's completion marker. Publish it only once
  // all three complete manifests are available; matching partial output is safe
  // to reuse and conflicting existing files are never replaced.
  for (const [name, xml] of rendered)
    await publish(path.join(directory, "manifests", `${name}.xml`), xml);
  await publish(
    path.join(directory, "config.json"),
    JSON.stringify(
      { executable, cert, key, recoveryDir: path.join(directory, "recovery"), port },
      null,
      2,
    ) + "\n",
  );
  return directory;
}
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const args = new Map();
  for (let i = 2; i < process.argv.length; i += 2) args.set(process.argv[i], process.argv[i + 1]);
  try {
    const directory = await setup({
      executable: args.get("--executable"),
      cert: args.get("--cert"),
      key: args.get("--key"),
      directory: args.get("--directory") ?? path.join(os.homedir(), ".reshiki-office"),
      port: args.has("--port") ? Number(args.get("--port")) : 43127,
    });
    console.log(
      `Configuration prepared: ${path.join(directory, "config.json")}\nManifests: ${path.join(directory, "manifests")}\nNo certificate trust or Office sideload settings were changed.`,
    );
  } catch (error) {
    console.error(
      `${error.message}\nUsage: node setup.js --executable /absolute/path/reshiki --cert /absolute/path/localhost.crt --key /absolute/path/localhost.key [--directory /absolute/path/install] [--port 43127]`,
    );
    process.exitCode = 1;
  }
}
