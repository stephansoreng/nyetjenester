import fs from "node:fs";
import path from "node:path";
import ExcelJS from "exceljs";

const root = process.cwd();
const inputDir = path.join(root, "input");
const outputDir = path.join(root, "src", "data");
const outputFile = path.join(outputDir, "requests.json");

const requiredColumns = [
  "Number",
  "Status",
  "Title",
  "Assignment Group",
  "Hn Phase Placeholder",
];

const phaseLabels = {
  proposed: "Foreslåtte initiativ",
  intake: "Inntakskø",
  design: "Løsningsdesign",
  offer: "Tilbud",
  execution: "Gjennomføring",
  closing: "Avslutning",
  unclassified: "Uklassifisert",
};

function normalize(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function findLatestExcelFile() {
  const candidates = [];
  for (const dir of [inputDir, root]) {
    if (!fs.existsSync(dir)) continue;
    for (const file of fs.readdirSync(dir)) {
      if (file.startsWith("~$")) continue;
      if (!file.toLowerCase().endsWith(".xlsx")) continue;
      candidates.push(path.join(dir, file));
    }
    if (candidates.length > 0 && dir === inputDir) break;
  }

  return candidates
    .map((filePath) => ({ filePath, mtime: fs.statSync(filePath).mtimeMs }))
    .sort((a, b) => b.mtime - a.mtime)[0]?.filePath;
}

function derivePhase(status, phasePlaceholder) {
  const statusKey = normalize(status);
  const phaseKey = normalize(phasePlaceholder);

  if (phaseKey.startsWith("1.")) return phaseLabels.proposed;
  if (phaseKey.startsWith("2.") && statusKey === "inntakskø") return phaseLabels.intake;
  if (phaseKey.startsWith("2.")) return phaseLabels.design;
  if (phaseKey.startsWith("3.")) return phaseLabels.offer;
  if (phaseKey.startsWith("4.")) return phaseLabels.execution;
  if (phaseKey.startsWith("5.") || phaseKey.startsWith("6.")) return phaseLabels.closing;
  return phaseLabels.unclassified;
}

function findColumn(headers, wanted) {
  const wantedKey = normalize(wanted);
  return headers.findIndex((header) => normalize(header) === wantedKey);
}

function cellToText(value) {
  if (value == null) return "";
  if (value instanceof Date) return value.toISOString();
  if (typeof value !== "object") return String(value).trim();
  if ("text" in value) return String(value.text ?? "").trim();
  if ("result" in value) return String(value.result ?? "").trim();
  if ("richText" in value) {
    return value.richText.map((part) => part.text ?? "").join("").trim();
  }
  return String(value).trim();
}

const sourceFile = findLatestExcelFile();
if (!sourceFile) {
  throw new Error("Fant ingen .xlsx-fil i input/ eller prosjektroten.");
}

const workbook = new ExcelJS.Workbook();
await workbook.xlsx.readFile(sourceFile);
const sheet = workbook.worksheets[0];
const sheetName = sheet?.name;
if (!sheet) {
  throw new Error("Fant ingen arkfaner i Excel-filen.");
}

const headerRow = sheet.getRow(1);
const headers = [];
headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
  headers[colNumber - 1] = cellToText(cell.value);
});

const columnMap = Object.fromEntries(
  requiredColumns.map((column) => [column, findColumn(headers, column)])
);

const missingColumns = requiredColumns.filter((column) => columnMap[column] === -1);
if (missingColumns.length > 0) {
  throw new Error(`Mangler kolonner i Excel-arket: ${missingColumns.join(", ")}`);
}

const requests = [];
sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
  if (rowNumber === 1) return;

  const get = (column) => cellToText(row.getCell(columnMap[column] + 1).value);
  const number = get("Number");
  const title = get("Title");
  const status = get("Status");
  const assignmentGroup = get("Assignment Group") || "Uten eier";
  const phasePlaceholder = get("Hn Phase Placeholder");

  if (!number && !title) return;

  requests.push({
    id: number || `rad-${rowNumber}`,
    number,
    title,
    status,
    assignmentGroup,
    phasePlaceholder,
    derivedPhase: derivePhase(status, phasePlaceholder),
    sourceRow: rowNumber,
  });
});

const payload = {
  metadata: {
    sourceFile: path.basename(sourceFile),
    sourcePath: path.relative(root, sourceFile),
    sheetName,
    importedAt: new Date().toISOString(),
    total: requests.length,
  },
  requests,
};

fs.mkdirSync(outputDir, { recursive: true });
fs.writeFileSync(outputFile, `${JSON.stringify(payload, null, 2)}\n`, "utf8");

console.log(`Importerte ${requests.length} saker fra ${path.relative(root, sourceFile)}.`);
