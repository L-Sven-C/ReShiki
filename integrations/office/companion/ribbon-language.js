// Only the installation owner supplies this path in the companion config.
// Requests may choose a supported language, never a path or arbitrary XML.
import { lstat, readFile, writeFile, rename, unlink } from "node:fs/promises";
import path from "node:path";
import { randomUUID } from "node:crypto";
import { HOSTS, manifest } from "../setup.js";

export function ribbonLanguageManager(file, port) {
  if (file && !path.isAbsolute(file))
    throw new Error("wordManifest must be an absolute path");
  let queue = Promise.resolve();
  const apply = async (language) => {
    if (language !== undefined && !["en-US", "zh-CN"].includes(language))
      throw new Error("Unsupported ribbon language");
    if (!file) {
      if (language !== undefined)
        throw new Error("Ribbon language management is not configured");
      return { managed: false };
    }
    const info = await lstat(file);
    if (!info.isFile() || info.isSymbolicLink() || info.size > 65536)
      throw new Error("Managed Word manifest must be a regular file");
    const original = await readFile(file, "utf8");
    const current = ["en-US", "zh-CN"].find(
      (value) => original === manifest(HOSTS[0], port, value),
    );
    if (!current)
      throw new Error(
        "Word manifest has custom changes. Reinstall the current generated manifest before changing its language.",
      );
    if (language !== undefined && language !== current) {
      const temporary = path.join(
        path.dirname(file),
        `.reshiki-${randomUUID()}.tmp`,
      );
      try {
        await writeFile(temporary, manifest(HOSTS[0], port, language), {
          flag: "wx",
          mode: 0o600,
          flush: true,
        });
        // Do not clobber an installation changed by another process while writing.
        const latest = await lstat(file);
        if (
          latest.isSymbolicLink() ||
          latest.ino !== info.ino ||
          (await readFile(file, "utf8")) !== original
        )
          throw new Error(
            "Word manifest changed; retry after checking the installation",
          );
        await rename(temporary, file);
      } finally {
        await unlink(temporary).catch((error) => {
          if (error.code !== "ENOENT") throw error;
        });
      }
    }
    return {
      managed: true,
      language: language ?? current,
      restartRequired: language !== undefined,
    };
  };
  return (language) => {
    const operation = queue.then(() => apply(language));
    queue = operation.catch(() => {});
    return operation;
  };
}
