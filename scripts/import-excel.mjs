import fs from "node:fs";
import path from "node:path";
import ExcelJS from "exceljs";

const root = process.cwd();
const inputDir = path.join(root, "input");
const outputDir = path.join(root, "src", "data");
const outputFile = path.join(outputDir, "requests.json");

const EPJ_FILE = path.join(inputDir, "features_EPJ.xlsx");
const EPJ_ASSIGNMENT_GROUP = "EPJ - Produktområdeledelse";

const requiredColumns = [
  "Number",
  "Status",
  "Title",
  "Assignment Group",
  "Hn Phase Placeholder",
];

const epjRequiredColumns = ["Feature_ID", "Feature navn", "Status"];

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
  const epjBasename = path.basename(EPJ_FILE).toLowerCase();
  const candidates = [];
  for (const dir of [inputDir, root]) {
    if (!fs.existsSync(dir)) continue;
    for (const file of fs.readdirSync(dir)) {
      if (file.startsWith("~$")) continue;
      if (!file.toLowerCase().endsWith(".xlsx")) continue;
      if (file.toLowerCase() === epjBasename) continue;
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

function deriveEpjPhase(status) {
  switch (normalize(status)) {
    case "blokkert":
    case "under arbeid":
      return phaseLabels.execution;
    case "klar for pi-planning":
      return phaseLabels.intake;
    case "under analyse":
      return phaseLabels.design;
    case "pi-backlog":
    case "ideer og ønsker":
    case "":
      return phaseLabels.proposed;
    default:
      return phaseLabels.proposed;
  }
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

async function importMainFile(sourceFile) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(sourceFile);
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error("Fant ingen arkfaner i Excel-filen.");

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
      source: "nye-tjenester",
      sourceRow: rowNumber,
    });
  });

  return { sheet, requests };
}

async function importEpjFile() {
  if (!fs.existsSync(EPJ_FILE)) {
    console.warn(`[import-excel] ${path.relative(root, EPJ_FILE)} ikke funnet — hopper over EPJ-import.`);
    return [];
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(EPJ_FILE);
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error("Fant ingen arkfaner i features_EPJ.xlsx.");

  const headerRow = sheet.getRow(1);
  const headers = [];
  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    headers[colNumber - 1] = cellToText(cell.value);
  });

  const columnMap = Object.fromEntries(
    epjRequiredColumns.map((column) => [column, findColumn(headers, column)])
  );

  const missingColumns = epjRequiredColumns.filter((column) => columnMap[column] === -1);
  if (missingColumns.length > 0) {
    throw new Error(`Mangler kolonner i features_EPJ.xlsx: ${missingColumns.join(", ")}`);
  }

  const piColIndex = findColumn(headers, "PI");

  const requests = [];
  sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;
    const get = (column) => cellToText(row.getCell(columnMap[column] + 1).value);
    const number = get("Feature_ID");
    const title = get("Feature navn");
    const status = get("Status");
    const pi = piColIndex >= 0 ? cellToText(row.getCell(piColIndex + 1).value) : "";

    if (!number && !title) return;
    if (normalize(status) === "kansellert") return;

    requests.push({
      id: number || `epj-rad-${rowNumber}`,
      number,
      title,
      status,
      pi,
      assignmentGroup: EPJ_ASSIGNMENT_GROUP,
      phasePlaceholder: "",
      derivedPhase: deriveEpjPhase(status),
      source: "epj",
      sourceRow: rowNumber,
    });
  });

  console.log(`Importerte ${requests.length} EPJ-saker fra ${path.relative(root, EPJ_FILE)}.`);
  return requests;
}

// --- Kjør import ---

const sourceFile = findLatestExcelFile();
if (!sourceFile) {
  console.warn(
    "[import-excel] Ingen .xlsx-fil funnet i input/ eller prosjektroten — " +
    "hopper over import. Bruker eksisterende src/data/requests.json."
  );
  process.exit(0);
}

const { sheet, requests: mainRequests } = await importMainFile(sourceFile);
const epjRequests = await importEpjFile();

const allRequests = [...mainRequests, ...epjRequests];

const payload = {
  metadata: {
    sourceFile: path.basename(sourceFile),
    sourcePath: path.relative(root, sourceFile),
    sheetName: sheet.name,
    epjSourceFile: fs.existsSync(EPJ_FILE) ? path.basename(EPJ_FILE) : null,
    importedAt: new Date().toISOString(),
    total: allRequests.length,
    totalMain: mainRequests.length,
    totalEpj: epjRequests.length,
  },
  requests: allRequests,
};

fs.mkdirSync(outputDir, { recursive: true });
fs.writeFileSync(outputFile, `${JSON.stringify(payload, null, 2)}\n`, "utf8");

console.log(`Importerte ${mainRequests.length} saker fra ${path.relative(root, sourceFile)}.`);
console.log(`Totalt ${allRequests.length} saker skrevet til ${path.relative(root, outputFile)}.`);
