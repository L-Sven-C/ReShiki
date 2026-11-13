export function preferenceDOM(nodes) {
  for (const id of [
    "workspace",
    "preferences",
    "preferences-open",
    "preferences-close",
    "ui-language",
    "ribbon-language",
    "apply-ribbon",
    "compact-layout",
    "show-hints",
    "usage-hints",
    "preferences-status",
    "ribbon-status",
  ])
    nodes.set(id, {
      id,
      value: "",
      checked: false,
      disabled: false,
      hidden: false,
      textContent: "",
      listeners: {},
      attributes: {},
      addEventListener(name, fn) {
        this.listeners[name] = fn;
      },
      focus() {
        this.focused = true;
      },
      setAttribute(key, value) {
        this.attributes[key] = value;
      },
      getAttribute(key) {
        return this.attributes[key];
      },
    });
  const file = nodes.get("file");
  if (file && !file.setAttribute)
    file.setAttribute = function (key, value) {
      this[key] = value;
    };
  return { documentElement: {}, body: { dataset: {} } };
}
