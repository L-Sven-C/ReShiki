// UI-only preferences. No document data, credentials or session IDs are stored.
export const DEFAULTS = Object.freeze({
  language: "en-US",
  compact: true,
  hints: false,
});
const KEY = "reshiki.office.preferences.v1";
export function normalizePreferences(value) {
  return {
    language: value?.language === "zh-CN" ? "zh-CN" : "en-US",
    compact: typeof value?.compact === "boolean" ? value.compact : true,
    hints: typeof value?.hints === "boolean" ? value.hints : false,
  };
}
const labels = {
  subtitle: ["Chemical drawings for Office", "Office 化学结构"],
  drawing: ["Drawing", "结构"],
  edit: ["Edit Selected Structure", "编辑选中结构"],
  import: ["Insert File…", "插入文件…"],
  paste: ["Paste Structure", "粘贴结构"],
  copy: ["Copy Selected Structure", "复制选中结构"],
  saveHint: [
    "Save in ReShiki, wait for confirmation, then save your Office document.",
    "在 ReShiki 保存并等待更新确认，再保存 Office 文档。",
  ],
  insertHint: [
    "Place the cursor outside existing drawings before inserting. Plain images cannot be edited.",
    "插入前请将光标置于已有结构之外。普通图片无法直接编辑。",
  ],
  endHint: [
    "Keep this pane visible during saving. Close the native drawing after confirmation to finish editing.",
    "保存时保持窗格可见。收到确认后关闭原生结构窗口即可结束编辑。",
  ],
  about: ["About editable drawings", "关于可编辑结构"],
  aboutText: [
    "Native data stays with each preview in your document. Use Copy Structure and Paste Structure to keep it editable. Recovery drafts stay on this computer.",
    "原始数据与预览图一起保存在文档中。使用复制结构与粘贴结构以保留可编辑性。恢复草稿保存在本机。",
  ],
  preferences: ["Preferences", "偏好设置"],
  done: ["Done", "完成"],
  paneLanguage: ["Pane language", "窗格语言"],
  paneLanguageHint: [
    "Applies immediately. Saved on this computer.",
    "立即生效，保存在本机。",
  ],
  ribbonLanguage: ["Ribbon language", "功能区语言"],
  applyRibbon: ["Apply Ribbon Language", "应用功能区语言"],
  appearance: ["Appearance", "外观"],
  compact: ["Compact spacing", "紧凑间距"],
  hints: ["Show workflow tips", "显示操作提示"],
  scope: [
    "These preferences change the add-in interface only. Your structures and document formatting stay unchanged.",
    "这些设置仅改变加载项界面，不修改结构数据或文档格式。",
  ],
};
const messages = [
  ["Connecting to Office…", "正在连接 Office…"],
  ["Connected to ", "已连接到 "],
  [
    "An Office operation is in progress. Please try again when it finishes.",
    "Office 操作正在进行，请完成后重试。",
  ],
  [
    "Select an editable ReShiki drawing. A plain image does not contain editable drawing data.",
    "请选择可编辑的 ReShiki 结构。普通图片不包含可编辑结构数据。",
  ],
  [
    "Editable drawing inserted. Save your Office file to keep it.",
    "已插入可编辑结构，请保存 Office 文档。",
  ],
  [
    "Editable drawing copied. Use Paste editable drawing in the destination Office file.",
    "已复制可编辑结构，请在目标 Office 文档中使用粘贴结构。",
  ],
  [
    "Edit session finished. You can select and reopen the drawing.",
    "编辑会话已结束，可以再次选中并打开结构。",
  ],
  [
    "Editing in ReShiki. Save there to update this drawing.",
    "正在 ReShiki 中编辑，保存后更新此结构。",
  ],
  [
    "ReShiki opened. Keep this pane open until the save is confirmed.",
    "已打开 ReShiki。请保持窗格可见，直到保存获得确认。",
  ],
  [
    "Edit finished. Select the drawing to reopen it.",
    "编辑已结束，选中结构即可再次打开。",
  ],
  [
    "Drawing updated in Office. Save your Office file to keep it.",
    "已更新 Office 中的结构，请保存 Office 文档。",
  ],
  ["This drawing already has an open edit session", "此结构已有打开的编辑会话"],
  ["End edit session", "结束编辑会话"],
  ["Recovery: ", "恢复文件："],
  ["Update not confirmed: ", "更新未获确认："],
  ["Your native draft is retained: ", "原始草稿已保留："],
  [
    "Resolve the Office error, then save again in ReShiki to retry.",
    "解决 Office 错误后，在 ReShiki 中再次保存以重试。",
  ],
  [
    "Reopen this pane and import the recovery file to recover it as a new drawing.",
    "重新打开窗格并导入恢复文件，将其恢复为新结构。",
  ],
  [
    "Office read-back succeeded; retrying save acknowledgement: ",
    "Office 数据核对成功，正在重试保存确认：",
  ],
  ["Companion connection interrupted: ", "本地服务连接中断："],
  ["Companion heartbeat interrupted: ", "本地服务心跳中断："],
  [
    "Native drafts remain in the recovery folder.",
    "原始草稿仍保留在恢复目录。",
  ],
  ["Native drafts remain in recovery.", "原始草稿仍保留在恢复目录。"],
  [
    "Choose a .rsk file or Paste Structure. Place the cursor outside existing drawings first.",
    "选择 .rsk 文件或粘贴结构。请先将光标置于已有结构之外。",
  ],
  ["Preferences saved on this computer.", "偏好设置已保存到本机。"],
  [
    "Preferences apply for this session; local storage is unavailable.",
    "本地存储不可用，偏好设置仅在本次会话生效。",
  ],
  [
    "Finish editing and save your documents, then restart Word to load the selected ribbon language.",
    "结束编辑并保存文档后，重启 Word 以加载所选功能区语言。",
  ],
  [
    "Ribbon language is set in the installed manifest. Restart Word after applying a change.",
    "功能区语言由安装清单决定，应用修改后需重启 Word。",
  ],
  [
    "Ribbon language management is not configured. See the local installation guide.",
    "尚未配置功能区语言管理，请参阅本地安装说明。",
  ],
];
export function translate(message, language) {
  if (language !== "zh-CN") return message;
  return messages.reduce((text, [en, zh]) => text.split(en).join(zh), message);
}
export function createPreferences({ document, storage, api }) {
  let preferences;
  try {
    preferences = normalizePreferences(
      JSON.parse(storage?.getItem(KEY) || "null"),
    );
  } catch {
    preferences = { ...DEFAULTS };
  }
  const node = (id) => document.getElementById(id);
  const liveMessages = new Map();
  const message = (target, text) => {
    liveMessages.set(target, text);
    target.textContent = translate(text, preferences.language);
  };
  const apply = () => {
    document.documentElement.lang = preferences.language;
    document.body.dataset.compact = String(preferences.compact);
    document.querySelectorAll("[data-i18n]").forEach((target) => {
      const text = labels[target.dataset.i18n];
      if (text)
        target.textContent = text[preferences.language === "zh-CN" ? 1 : 0];
    });
    for (const [target, text] of liveMessages)
      target.textContent = translate(text, preferences.language);
    node("ui-language").value = preferences.language;
    node("compact-layout").checked = preferences.compact;
    node("show-hints").checked = preferences.hints;
    node("usage-hints").hidden = !preferences.hints;
    node("preferences-open").setAttribute(
      "aria-label",
      labels.preferences[preferences.language === "zh-CN" ? 1 : 0],
    );
    node("preferences-open").title =
      node("preferences-open").getAttribute("aria-label");
    node("file").setAttribute(
      "aria-label",
      labels.import[preferences.language === "zh-CN" ? 1 : 0],
    );
  };
  const save = () => {
    preferences = normalizePreferences({
      language: node("ui-language").value,
      compact: node("compact-layout").checked,
      hints: node("show-hints").checked,
    });
    apply();
    try {
      if (!storage) throw new Error("Unavailable");
      storage.setItem(KEY, JSON.stringify(preferences));
      message(
        node("preferences-status"),
        "Preferences saved on this computer.",
      );
    } catch {
      message(
        node("preferences-status"),
        "Preferences apply for this session; local storage is unavailable.",
      );
    }
  };
  const showDrawings = () => {
    node("workspace").hidden = false;
    node("preferences").hidden = true;
  };
  const showPreferences = async () => {
    node("workspace").hidden = true;
    node("preferences").hidden = false;
    node("ui-language").focus();
    try {
      const result = await api("ribbon-language");
      node("ribbon-language").value = result.language || "en-US";
      node("ribbon-language").disabled = !result.managed;
      node("apply-ribbon").disabled = !result.managed;
      message(
        node("ribbon-status"),
        result.managed
          ? "Ribbon language is set in the installed manifest. Restart Word after applying a change."
          : "Ribbon language management is not configured. See the local installation guide.",
      );
    } catch (error) {
      message(node("ribbon-status"), error.message);
    }
  };
  node("preferences-open").addEventListener("click", showPreferences);
  node("preferences-close").addEventListener("click", showDrawings);
  for (const id of ["ui-language", "compact-layout", "show-hints"])
    node(id).addEventListener("change", save);
  node("apply-ribbon").addEventListener("click", async () => {
    node("apply-ribbon").disabled = true;
    try {
      await api("ribbon-language", { language: node("ribbon-language").value });
      message(
        node("ribbon-status"),
        "Finish editing and save your documents, then restart Word to load the selected ribbon language.",
      );
    } catch (error) {
      message(node("ribbon-status"), error.message);
    } finally {
      node("apply-ribbon").disabled = false;
    }
  });
  apply();
  return { message, showDrawings, showPreferences };
}
