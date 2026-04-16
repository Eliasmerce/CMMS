const defaultSheets = [
  {
    name: "Formulario",
    purpose:
      "Main user-facing form for maintenance reports, downtime timestamps, materials, and signatures.",
  },
  {
    name: "Base de Datos",
    purpose:
      "Historical row-level log of all maintenance events and downtime records.",
  },
  {
    name: "Calculo Indicadores",
    purpose:
      "KPI calculation engine for downtime, production loss, and availability metrics.",
  },
  {
    name: "Summary / Charts",
    purpose:
      "Visual dashboard for failures, area availability, critical equipment performance, and trends.",
  },
  {
    name: "Support Sheets",
    purpose:
      "Helper tables for dropdown lists, reference data, and intermediate summary ranges.",
  },
];

const defaultFlow = [
  "Operator fills Formulario.",
  "Validation/macro checks mandatory fields and time consistency.",
  "Record is appended into Base de Datos.",
  "Calculo Indicadores computes downtime and availability KPIs.",
  "Summary sheets aggregate by area, equipment, and period.",
  "Charts visualize trend and performance for operations review.",
];

const riskItems = [
  "Macros disabled -> save/refresh actions fail.",
  "XLOOKUP/dynamic arrays may break in legacy Excel versions.",
  "Hard-coded ranges can break when sheets/columns change.",
  "Manual datetime entry can produce negative or cross-day errors.",
  "No duplicate report-number control can affect KPI trust.",
];

const improvementItems = [
  "Use Excel Table for Base de Datos with structured references.",
  "Add validation and duplicate-report checks before save.",
  "Centralize KPI constants (cost/hour, target availability).",
  "Use robust VBA pattern: validate -> write -> verify -> notify.",
  "Implement compatibility fallback logic for lookup formulas.",
  "Protect formula sheets and keep only entry cells unlocked.",
];

const sheetList = document.getElementById("sheetList");
const sheetTemplate = document.getElementById("sheetTemplate");
const addSheetBtn = document.getElementById("addSheetBtn");
const flowList = document.getElementById("flowList");
const riskList = document.getElementById("riskList");
const improvementList = document.getElementById("improvementList");
const output = document.getElementById("output");
const generateBtn = document.getElementById("generateBtn");
const downloadBtn = document.getElementById("downloadBtn");

let generatedReport = "";

function addSheetCard(sheet = { name: "", purpose: "" }) {
  const frag = sheetTemplate.content.cloneNode(true);
  const card = frag.querySelector(".sheet-card");
  card.querySelector(".sheet-name").value = sheet.name;
  card.querySelector(".sheet-purpose").value = sheet.purpose;

  card.querySelector(".remove-sheet").addEventListener("click", () => {
    card.remove();
  });

  sheetList.appendChild(frag);
}

function addCheckList(container, items, prefix) {
  const ul = document.createElement("ul");
  items.forEach((item, index) => {
    const li = document.createElement("li");
    const id = `${prefix}-${index}`;
    li.innerHTML = `<label><input id="${id}" type="checkbox" checked /> ${item}</label>`;
    ul.appendChild(li);
  });
  container.appendChild(ul);
}

function getSelectedFrom(container) {
  return [...container.querySelectorAll("input[type='checkbox']")]
    .filter((box) => box.checked)
    .map((box) => box.parentElement.textContent.trim());
}

function getSheets() {
  return [...sheetList.querySelectorAll(".sheet-card")].map((card) => ({
    name: card.querySelector(".sheet-name").value.trim() || "Unnamed sheet",
    purpose: card.querySelector(".sheet-purpose").value.trim() || "No purpose provided.",
  }));
}

function renderDefaults() {
  defaultSheets.forEach(addSheetCard);
  flowList.innerHTML = defaultFlow.map((step) => `<li>${step}</li>`).join("");
  addCheckList(riskList, riskItems, "risk");
  addCheckList(improvementList, improvementItems, "improvement");
}

function generateReport() {
  const workbookName = document.getElementById("workbookName").value.trim() || "Workbook";
  const version = document.getElementById("excelVersion").value.trim() || "Not specified";
  const notes = document.getElementById("contextNotes").value.trim() || "No extra notes.";
  const sheets = getSheets();
  const selectedRisks = getSelectedFrom(riskList);
  const selectedImprovements = getSelectedFrom(improvementList);

  const sheetSection = sheets
    .map((s) => `- **${s.name}**: ${s.purpose}`)
    .join("\n");

  const flowSection = [...flowList.querySelectorAll("li")]
    .map((li, idx) => `${idx + 1}. ${li.textContent.trim()}`)
    .join("\n");

  const riskSection = selectedRisks.length
    ? selectedRisks.map((r) => `- ${r}`).join("\n")
    : "- No specific risks selected.";

  const improvementSection = selectedImprovements.length
    ? selectedImprovements.map((i) => `- ${i}`).join("\n")
    : "- No improvements selected.";

  generatedReport = `# Functional Workbook Assessment\n\n## Profile\n- **Workbook**: ${workbookName}\n- **Target Excel environment**: ${version}\n- **Context**: ${notes}\n\n## Sheet Roles\n${sheetSection}\n\n## Data Flow\n${flowSection}\n\n## Key Risks\n${riskSection}\n\n## Recommended Improvements\n${improvementSection}\n\n## Conclusion\nThis workbook behaves as a lightweight Excel-based maintenance management application. It can remain aligned to business intent while becoming safer, more reliable, and easier to operate by applying the selected improvements above.`;

  output.textContent = generatedReport;
  downloadBtn.disabled = false;
}

function downloadMarkdown() {
  if (!generatedReport) return;
  const blob = new Blob([generatedReport], { type: "text/markdown;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "workbook_assessment.md";
  link.click();
  URL.revokeObjectURL(link.href);
}

addSheetBtn.addEventListener("click", () => addSheetCard());
generateBtn.addEventListener("click", generateReport);
downloadBtn.addEventListener("click", downloadMarkdown);

renderDefaults();
