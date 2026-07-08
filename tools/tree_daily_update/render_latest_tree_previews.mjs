import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const root = "D:/AI/宏观框架分析";
const outputDir = path.join(root, "codex_tmp", "latest_tree_preview");
await fs.mkdir(outputDir, { recursive: true });

const entries = await fs.readdir(root, { withFileTypes: true });
const files = [];
for (const entry of entries) {
  if (!entry.isFile()) continue;
  if (!entry.name.endsWith(".xlsx")) continue;
  if (!entry.name.includes("TREE")) continue;
  if (entry.name.startsWith("~$")) continue;
  const full = path.join(root, entry.name);
  const stat = await fs.stat(full);
  files.push({ full, name: entry.name, mtimeMs: stat.mtimeMs });
}
files.sort((a, b) => b.mtimeMs - a.mtimeMs);
if (!files.length) throw new Error("No TREE workbook found");

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(files[0].full));
const previews = [
  { sheetName: "重点策略跟踪情况(V3)", range: "F1:U90", name: "tree_v3_macro_preview.png" },
  { sheetName: "重点策略跟踪情况(V3.0)", range: "A1:T90", name: "tree_v30_macro_preview.png" },
];
const out = [];
for (const item of previews) {
  const preview = await workbook.render({
    sheetName: item.sheetName,
    range: item.range,
    scale: 1,
    format: "png",
  });
  const bytes = new Uint8Array(await preview.arrayBuffer());
  const outPath = path.join(outputDir, item.name);
  await fs.writeFile(outPath, bytes);
  out.push({ ...item, output: outPath, bytes: bytes.length });
}
console.log(JSON.stringify({ workbook: files[0].full, previews: out }, null, 2));
