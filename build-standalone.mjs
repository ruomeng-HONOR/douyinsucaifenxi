import { readFile, writeFile } from "node:fs/promises";

const [html, css, analysis, storage, app] = await Promise.all([
  readFile(new URL("./index.html", import.meta.url), "utf8"),
  readFile(new URL("./styles.css", import.meta.url), "utf8"),
  readFile(new URL("./analysis.js", import.meta.url), "utf8"),
  readFile(new URL("./storage.js", import.meta.url), "utf8"),
  readFile(new URL("./app.js", import.meta.url), "utf8"),
]);

const script = [
  analysis.replaceAll("export ", ""),
  storage.replaceAll("export ", ""),
  app.replace(/^import .*$/gm, ""),
].join("\n\n");

const standalone = html
  .replace('<link rel="stylesheet" href="./styles.css">', `<style>\n${css}\n</style>`)
  .replace('<script src="./xlsx.full.min.js"></script>', '<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>')
  .replace('<script type="module" src="./app.js"></script>', `<script type="module">\n${script}\n</script>`);

await writeFile(new URL("./standalone.html", import.meta.url), standalone, "utf8");
