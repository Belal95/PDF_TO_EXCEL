import React, { useMemo, useState } from "react";
import * as pdfjsLib from "pdfjs-dist";
import workerSrc from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import ExcelJS from "exceljs";
import {
  Upload,
  FileSpreadsheet,
  FileText,
  PlusCircle,
  Download,
  AlertCircle,
  CheckCircle2,
} from "lucide-react";

pdfjsLib.GlobalWorkerOptions.workerSrc = workerSrc;

const AGENT_ORDER = {
  فوري: 1,
  خالص: 2,
  ضامن: 3,
  الاهلي: 4,
};

const OUTPUT_HEADERS = [
  "التاريخ",
  "وكيل البيع",
  "عدد التذاكر",
  "عمولة وكيل البيع",
  "عمولة ترانس اي تي",
  "عمولة الهيئة",
  "ايرادات التذاكر",
  "اجمالي ايراد الهيئة",
  "الاجمالي",
];

const REQUIRED_HEADER_MATCHES = {
  التاريخ: ["التاريخ"],
  "وكيل البيع": ["وكيل البيع", "الوكيل"],
  "عدد التذاكر": ["عدد التذاكر"],
  "عمولة وكيل البيع": ["عمولة وكيل البيع"],
  "عمولة ترانس اي تي": [
    "عمولة ترانس اي تي",
    "عموله ترنس اى تى",
    "عموله ترانس اي تي",
  ],
  "عمولة الهيئة": ["عمولة الهيئة", "عموله الهيئه", "عموله الهيئة"],
  "ايرادات التذاكر": ["ايرادات التذاكر", "ايرادات التذاكر"],
  "اجمالي ايراد الهيئة": [
    "اجمالي ايراد الهيئة",
    "اجمالى ايراد الهيئه",
    "اجمالى ايراد الهيئة",
  ],
  الاجمالي: ["الاجمالي", "الاجمالى"],
};

function normalizeArabicDigits(str) {
  if (!str) return "";
  const arabicDigits = "٠١٢٣٤٥٦٧٨٩";
  return String(str).replace(/[٠-٩]/g, (d) => String(arabicDigits.indexOf(d)));
}

function normalizeArabicText(str) {
  if (!str) return "";
  return normalizeArabicDigits(String(str))
    .normalize("NFKC")
    .replace(/\r/g, "\n")
    .replace(/\u00A0/g, " ")
    .replace(/[ـ]/g, "")
    .replace(/[–—]/g, "-")
    .replace(/[إأآ]/g, "ا")
    .replace(/ة/g, "ه")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeAgentName(raw) {
  const text = normalizeArabicText(raw);
  if (text.includes("فوري")) return "فوري";
  if (text.includes("خالص")) return "خالص";
  if (text.includes("ضامن")) return "ضامن";
  if (text.includes("الاهلي")) return "الاهلي";
  return text;
}

function parseArabicNumber(value) {
  if (value == null) return null;
  const normalized = normalizeArabicDigits(String(value))
    .replace(/[^\d.]/g, "")
    .trim();
  if (!normalized) return null;
  return Number(normalized);
}

function standardizeDate(dateStr) {
  const cleaned = normalizeArabicDigits(dateStr).replace(/[-.]/g, "/").trim();
  const [year, month, day] = cleaned.split("/");
  return `${day.padStart(2, "0")}/${month.padStart(2, "0")}/${year.padStart(4, "0")}`;
}

function extractDateFromText(text) {
  const normalized = normalizeArabicText(text);
  const sectionMatch = normalized.match(
    /من تاريخ\s*:?\s*(\d{4}[\/.-]\d{1,2}[\/.-]\d{1,2})\s*الي تاريخ/i,
  );
  if (sectionMatch?.[1]) return standardizeDate(sectionMatch[1]);
  const dates = normalized.match(/\d{4}[\/.-]\d{1,2}[\/.-]\d{1,2}/g);
  if (!dates?.length) throw new Error("تعذر العثور على تاريخ التقرير");
  return [...new Set(dates.map(standardizeDate))].sort()[0];
}

function sortRows(rows) {
  return [...rows].sort((a, b) => {
    if (a["التاريخ"] < b["التاريخ"]) return -1;
    if (a["التاريخ"] > b["التاريخ"]) return 1;
    const aOrder = AGENT_ORDER[a["وكيل البيع"]] ?? 999;
    const bOrder = AGENT_ORDER[b["وكيل البيع"]] ?? 999;
    if (aOrder !== bOrder) return aOrder - bOrder;
    return String(a["وكيل البيع"]).localeCompare(String(b["وكيل البيع"]), "ar");
  });
}

function buildDuplicateKey(row) {
  return `${row["التاريخ"]}||${normalizeAgentName(row["وكيل البيع"] || "")}`;
}

function dedupeRows(rows) {
  const seen = new Set();
  const kept = [];
  const duplicates = [];
  for (const row of rows) {
    const key = buildDuplicateKey(row);
    if (seen.has(key)) duplicates.push(row);
    else {
      seen.add(key);
      kept.push(row);
    }
  }
  return { kept, duplicates };
}

function isSummaryText(text) {
  const t = normalizeArabicText(text);
  return t.includes("اجمالي الفتره") || t.includes("اجمالي التقرير");
}

function reconstructRowsFromItems(items) {
  const rows = [];
  const yTolerance = 2.5;

  for (const item of items) {
    const str = String(item.str || "").trim();
    if (!str) continue;

    const x = item.transform?.[4] ?? 0;
    const y = item.transform?.[5] ?? 0;
    const width = item.width ?? 0;

    let row = rows.find((r) => Math.abs(r.y - y) <= yTolerance);
    if (!row) {
      row = { y, parts: [] };
      rows.push(row);
    }

    row.parts.push({ x, y, width, str, normalized: normalizeArabicText(str) });
  }

  return rows.sort((a, b) => b.y - a.y);
}

function findDataRows(rowGroups) {
  return rowGroups.filter((row) => {
    const joined = normalizeArabicText(row.parts.map((p) => p.str).join(" "));
    if (!joined || isSummaryText(joined)) return false;
    const hasAgent = row.parts.some((p) => {
      const t = p.normalized;
      return (
        t.includes("فوري") ||
        t.includes("خالص") ||
        t.includes("ضامن") ||
        t.includes("الاهلي")
      );
    });
    const numericParts = row.parts.filter((p) =>
      /\d/.test(normalizeArabicDigits(p.str)),
    );
    return hasAgent && numericParts.length >= 7;
  });
}

function getAgentPart(parts) {
  const agentParts = parts.filter((p) => {
    const t = p.normalized;
    return (
      t.includes("فوري") ||
      t.includes("خالص") ||
      t.includes("ضامن") ||
      t.includes("الاهلي")
    );
  });
  if (!agentParts.length) return null;
  return agentParts.reduce((best, current) =>
    current.x > best.x ? current : best,
  );
}

function getNumericParts(parts, agentX) {
  return parts
    .filter((p) => /\d/.test(normalizeArabicDigits(p.str)))
    .filter((p) => p.x < agentX - 5)
    .sort((a, b) => b.x - a.x);
}

function mapRowByColumns(row, reportDate) {
  const agentPart = getAgentPart(row.parts);
  if (!agentPart) throw new Error("تعذر العثور على اسم الوكيل في صف البيانات");

  const numericParts = getNumericParts(row.parts, agentPart.x);
  if (numericParts.length < 7) {
    const debugLine = row.parts
      .sort((a, b) => b.x - a.x)
      .map((p) => p.str)
      .join(" | ");
    throw new Error(`تعذر العثور على 7 أعمدة رقمية في الصف: ${debugLine}`);
  }

  const values = numericParts.slice(0, 7).map((p) => parseArabicNumber(p.str));
  const agentName = normalizeAgentName(agentPart.str);

  return {
    التاريخ: reportDate,
    "وكيل البيع": agentName,
    "عدد التذاكر": values[0],
    "عمولة وكيل البيع": values[1],
    "عمولة ترانس اي تي": values[2],
    "عمولة الهيئة": values[3],
    "ايرادات التذاكر": values[4],
    "اجمالي ايراد الهيئة": values[5],
    الاجمالي: values[6],
  };
}

async function extractRowsFromPdf(file) {
  const arrayBuffer = await file.arrayBuffer();
  const pdfDoc = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const extractedRows = [];

  for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum++) {
    const page = await pdfDoc.getPage(pageNum);
    const content = await page.getTextContent();
    const rowGroups = reconstructRowsFromItems(content.items);
    const pageText = rowGroups
      .map((row) =>
        row.parts
          .sort((a, b) => b.x - a.x)
          .map((p) => p.str)
          .join(" "),
      )
      .join("\n");
    const reportDate = extractDateFromText(pageText);

    const dataRows = findDataRows(rowGroups);
    for (const row of dataRows) {
      extractedRows.push(mapRowByColumns(row, reportDate));
    }
  }

  if (!extractedRows.length) {
    throw new Error("تعذر العثور على أي صف صالح للوكلاء داخل الملف");
  }

  const { kept, duplicates } = dedupeRows(extractedRows);
  return { kept, duplicates };
}

function downloadBuffer(buffer, fileName) {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

async function buildNewWorkbook(rows) {
  const workbook = new ExcelJS.Workbook();
  workbook.views = [{ rightToLeft: true }];

  const sheet = workbook.addWorksheet("data", {
    views: [{ rightToLeft: true, state: "frozen", ySplit: 1 }],
  });
  sheet.properties.defaultRowHeight = 20;

  sheet.columns = OUTPUT_HEADERS.map((header) => ({
    header,
    key: header,
    width: Math.max(16, header.length + 4),
    style: { alignment: { horizontal: "center", vertical: "middle" } },
  }));

  sortRows(rows).forEach((row) => {
    const cleanRow = {};
    for (const header of OUTPUT_HEADERS) cleanRow[header] = row[header];
    sheet.addRow(cleanRow);
  });

  sheet.getRow(1).font = { bold: true };
  sheet.getRow(1).alignment = {
    horizontal: "center",
    vertical: "middle",
    readingOrder: "rtl",
  };
  sheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: OUTPUT_HEADERS.length },
  };

  [
    "عدد التذاكر",
    "عمولة وكيل البيع",
    "عمولة ترانس اي تي",
    "عمولة الهيئة",
    "ايرادات التذاكر",
    "اجمالي ايراد الهيئة",
    "الاجمالي",
  ].forEach((key) => {
    const col = sheet.getColumn(key);
    col.numFmt = key === "عدد التذاكر" ? "0" : "0.00";
  });

  return workbook.xlsx.writeBuffer();
}

function getHeaderColumnMap(sheet) {
  const headerRow = sheet.getRow(1);
  const headerMap = {};
  headerRow.eachCell((cell, colNumber) => {
    const value = normalizeArabicText(String(cell.value || ""));
    if (value) headerMap[value] = colNumber;
  });

  const resolved = {};
  for (const [target, aliases] of Object.entries(REQUIRED_HEADER_MATCHES)) {
    const normalizedAliases = aliases.map(normalizeArabicText);
    const found = Object.entries(headerMap).find(([header]) =>
      normalizedAliases.includes(header),
    );
    if (!found)
      throw new Error(`الشيت data لا يحتوي على العمود المطلوب: ${target}`);
    resolved[target] = found[1];
  }
  return resolved;
}

async function appendToExistingWorkbook(file, extractedRows) {
  const buffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const sheet = workbook.getWorksheet("data") || workbook.getWorksheet("Data");
  if (!sheet) throw new Error('لم يتم العثور على الشيت باسم "data"');

  sheet.views = [{ rightToLeft: true, state: "frozen", ySplit: 1 }];
  const headerMap = getHeaderColumnMap(sheet);

  const existingRows = [];
  for (let i = 2; i <= sheet.rowCount; i++) {
    const row = sheet.getRow(i);
    const record = {};
    for (const header of OUTPUT_HEADERS) {
      const col = headerMap[header];
      const raw = row.getCell(col).value;
      record[header] =
        raw && typeof raw === "object" && raw.text ? raw.text : raw;
    }

    const normalized = {
      التاريخ: String(record["التاريخ"] || "").trim(),
      "وكيل البيع": normalizeAgentName(String(record["وكيل البيع"] || "")),
      "عدد التذاكر": Number(record["عدد التذاكر"] || 0),
      "عمولة وكيل البيع": Number(record["عمولة وكيل البيع"] || 0),
      "عمولة ترانس اي تي": Number(record["عمولة ترانس اي تي"] || 0),
      "عمولة الهيئة": Number(record["عمولة الهيئة"] || 0),
      "ايرادات التذاكر": Number(record["ايرادات التذاكر"] || 0),
      "اجمالي ايراد الهيئة": Number(record["اجمالي ايراد الهيئة"] || 0),
      الاجمالي: Number(record["الاجمالي"] || 0),
    };

    if (normalized["التاريخ"] && normalized["وكيل البيع"])
      existingRows.push(normalized);
  }

  const existingKeys = new Set(existingRows.map(buildDuplicateKey));
  const appended = [];
  const skippedDuplicates = [];

  for (const row of extractedRows) {
    const key = buildDuplicateKey(row);
    if (existingKeys.has(key)) skippedDuplicates.push(row);
    else {
      appended.push(row);
      existingKeys.add(key);
    }
  }

  const combined = sortRows([...existingRows, ...appended]);
  if (sheet.rowCount >= 2) sheet.spliceRows(2, sheet.rowCount - 1);

  combined.forEach((rowData) => {
    const row = sheet.addRow([]);
    for (const header of OUTPUT_HEADERS) {
      row.getCell(headerMap[header]).value = rowData[header];
    }
  });

  [
    "عدد التذاكر",
    "عمولة وكيل البيع",
    "عمولة ترانس اي تي",
    "عمولة الهيئة",
    "ايرادات التذاكر",
    "اجمالي ايراد الهيئة",
  ].forEach((key) => {
    const col = sheet.getColumn(headerMap[key]);
    col.numFmt = key === "عدد التذاكر" ? "0" : "0.00";
  });

  return {
    buffer: await workbook.xlsx.writeBuffer(),
    existingCount: existingRows.length,
    appendedCount: appended.length,
    skippedDuplicatesCount: skippedDuplicates.length,
    totalAfterAppend: combined.length,
  };
}

export default function SalesAgentsPdfToExcelApp() {
  const [pdfFiles, setPdfFiles] = useState([]);
  const [existingExcelFile, setExistingExcelFile] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [summary, setSummary] = useState(null);
  const [previewRows, setPreviewRows] = useState([]);
  const [errors, setErrors] = useState([]);

  const canExport = pdfFiles.length > 0;
  const canAppend = pdfFiles.length > 0 && existingExcelFile;

  async function runExtraction() {
    const collected = [];
    const extractionErrors = [];
    const batchDuplicates = [];

    for (const file of pdfFiles) {
      try {
        const { kept, duplicates } = await extractRowsFromPdf(file);
        collected.push(...kept);
        batchDuplicates.push(...duplicates);
      } catch (error) {
        extractionErrors.push({
          file: file.name,
          message: error.message || String(error),
        });
      }
    }

    const { kept, duplicates } = dedupeRows(collected);
    return {
      rows: sortRows(kept),
      duplicateRows: [...batchDuplicates, ...duplicates],
      extractionErrors,
    };
  }

  async function handleExportNew() {
    setIsProcessing(true);
    setErrors([]);
    try {
      const result = await runExtraction();
      setPreviewRows(result.rows);
      setErrors(result.extractionErrors);
      const buffer = await buildNewWorkbook(result.rows);
      downloadBuffer(buffer, `agents-report-${Date.now()}.xlsx`);
      setSummary({
        mode: "new",
        extracted: result.rows.length,
        appended: result.rows.length,
        skippedDuplicates: result.duplicateRows.length,
        errors: result.extractionErrors.length,
      });
    } catch (error) {
      setErrors([{ file: "عام", message: error.message || String(error) }]);
    } finally {
      setIsProcessing(false);
    }
  }

  async function handleAppendExisting() {
    setIsProcessing(true);
    setErrors([]);
    try {
      const result = await runExtraction();
      setPreviewRows(result.rows);
      setErrors(result.extractionErrors);
      const appendSummary = await appendToExistingWorkbook(
        existingExcelFile,
        result.rows,
      );
      downloadBuffer(
        appendSummary.buffer,
        `updated-${existingExcelFile.name.replace(/\.[^.]+$/, "")}.xlsx`,
      );
      setSummary({
        mode: "append",
        extracted: result.rows.length,
        appended: appendSummary.appendedCount,
        skippedDuplicates:
          appendSummary.skippedDuplicatesCount + result.duplicateRows.length,
        errors: result.extractionErrors.length,
        existingCount: appendSummary.existingCount,
        totalAfterAppend: appendSummary.totalAfterAppend,
      });
    } catch (error) {
      setErrors([{ file: "عام", message: error.message || String(error) }]);
    } finally {
      setIsProcessing(false);
    }
  }

  const sortedPreview = useMemo(() => sortRows(previewRows), [previewRows]);

  return (
    <div className="min-h-screen bg-slate-50 p-6 text-slate-900" dir="rtl">
      <div className="mx-auto max-w-6xl space-y-6">
        <div className="rounded-3xl bg-white p-6 shadow-sm ring-1 ring-slate-200">
          <div className="flex items-start justify-between gap-4">
            <div>
              <h1 className="text-2xl font-bold">
                تطبيق استخراج تقارير وكلاء البيع
              </h1>
              <p className="mt-2 text-sm text-slate-600">
                يستخرج بيانات الوكلاء من ملفات PDF العربية اعتمادًا على مواقع
                الأعمدة داخل الصفحة ويصدرها إلى Excel جديد أو يضيفها إلى ملف
                قائم مع فحص التكرار.
              </p>
            </div>
            <div className="rounded-2xl bg-slate-100 p-3">
              <FileSpreadsheet className="h-7 w-7" />
            </div>
          </div>
        </div>

        <div className="grid gap-6 lg:grid-cols-2">
          <div className="rounded-3xl bg-white p-6 shadow-sm ring-1 ring-slate-200">
            <div className="mb-4 flex items-center gap-3">
              <Upload className="h-5 w-5" />
              <h2 className="text-lg font-semibold">ملفات PDF</h2>
            </div>
            <input
              type="file"
              multiple
              accept=".pdf"
              onChange={(e) => setPdfFiles(Array.from(e.target.files || []))}
              className="mt-2 block w-full rounded-xl border border-slate-300 bg-white p-3 text-sm"
            />
            <div className="mt-4 space-y-2 text-sm">
              {pdfFiles.length === 0 ? (
                <div className="rounded-xl bg-slate-50 px-3 py-2 text-slate-500">
                  لم يتم اختيار ملفات PDF بعد.
                </div>
              ) : (
                pdfFiles.map((file) => (
                  <div
                    key={file.name}
                    className="flex items-center gap-2 rounded-xl bg-slate-50 px-3 py-2"
                  >
                    <FileText className="h-4 w-4" />
                    <span>{file.name}</span>
                  </div>
                ))
              )}
            </div>
          </div>

          <div className="rounded-3xl bg-white p-6 shadow-sm ring-1 ring-slate-200">
            <div className="mb-4 flex items-center gap-3">
              <PlusCircle className="h-5 w-5" />
              <h2 className="text-lg font-semibold">ملف Excel الحالي</h2>
            </div>
            <input
              type="file"
              accept=".xlsx,.xlsm,.xls"
              onChange={(e) =>
                setExistingExcelFile(e.target.files?.[0] || null)
              }
              className="mt-2 block w-full rounded-xl border border-slate-300 bg-white p-3 text-sm"
            />
            <div className="mt-4 rounded-xl bg-slate-50 px-3 py-2 text-sm text-slate-600">
              سيتم الإضافة إلى الشيت <span className="font-semibold">data</span>{" "}
              بدون تغيير هيكل الأعمدة، مع الحفاظ على اتجاه Excel من اليمين إلى
              اليسار.
            </div>
            {existingExcelFile && (
              <div className="mt-3 rounded-xl bg-slate-50 px-3 py-2 text-sm">
                الملف المختار:{" "}
                <span className="font-medium">{existingExcelFile.name}</span>
              </div>
            )}
          </div>
        </div>

        <div className="rounded-3xl bg-white p-6 shadow-sm ring-1 ring-slate-200">
          <div className="flex flex-wrap gap-3">
            <button
              onClick={handleExportNew}
              disabled={!canExport || isProcessing}
              className="inline-flex items-center gap-2 rounded-2xl bg-slate-900 px-5 py-3 text-sm font-medium text-white disabled:cursor-not-allowed disabled:opacity-50"
            >
              <Download className="h-4 w-4" />
              تصدير إلى ملف جديد
            </button>
            <button
              onClick={handleAppendExisting}
              disabled={!canAppend || isProcessing}
              className="inline-flex items-center gap-2 rounded-2xl bg-emerald-600 px-5 py-3 text-sm font-medium text-white disabled:cursor-not-allowed disabled:opacity-50"
            >
              <PlusCircle className="h-4 w-4" />
              إضافة إلى الملف الحالي
            </button>
          </div>
          <p className="mt-3 text-sm text-slate-500">
            {isProcessing
              ? "جارٍ المعالجة..."
              : "اختر ملفات PDF ثم اختر طريقة الإخراج."}
          </p>
        </div>

        {summary && (
          <div className="rounded-3xl bg-white p-6 shadow-sm ring-1 ring-slate-200">
            <div className="mb-4 flex items-center gap-3 text-emerald-700">
              <CheckCircle2 className="h-5 w-5" />
              <h2 className="text-lg font-semibold">ملخص العملية</h2>
            </div>
            <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-3">
              <SummaryCard
                label="نوع العملية"
                value={
                  summary.mode === "append"
                    ? "إضافة إلى ملف موجود"
                    : "تصدير ملف جديد"
                }
              />
              <SummaryCard
                label="عدد الصفوف المستخرجة"
                value={summary.extracted}
              />
              <SummaryCard
                label="عدد الصفوف المضافة"
                value={summary.appended}
              />
              <SummaryCard
                label="عدد الصفوف المتخطاة بسبب التكرار"
                value={summary.skippedDuplicates}
              />
              <SummaryCard label="عدد الأخطاء" value={summary.errors} />
              {summary.totalAfterAppend != null && (
                <SummaryCard
                  label="إجمالي الصفوف بعد الإضافة"
                  value={summary.totalAfterAppend}
                />
              )}
            </div>
          </div>
        )}

        <div className="grid gap-6 lg:grid-cols-2">
          <div className="rounded-3xl bg-white p-6 shadow-sm ring-1 ring-slate-200">
            <h2 className="mb-4 text-lg font-semibold">
              معاينة البيانات المستخرجة
            </h2>
            <div className="overflow-auto rounded-2xl border border-slate-200">
              <table className="min-w-full divide-y divide-slate-200 text-sm">
                <thead className="bg-slate-50">
                  <tr>
                    {OUTPUT_HEADERS.map((header) => (
                      <th
                        key={header}
                        className="whitespace-nowrap px-3 py-2 text-right font-semibold text-slate-700"
                      >
                        {header}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 bg-white">
                  {sortedPreview.length === 0 ? (
                    <tr>
                      <td
                        colSpan={OUTPUT_HEADERS.length}
                        className="px-3 py-6 text-center text-slate-500"
                      >
                        لا توجد بيانات مستخرجة حتى الآن.
                      </td>
                    </tr>
                  ) : (
                    sortedPreview.map((row, index) => (
                      <tr
                        key={`${row["التاريخ"]}-${row["وكيل البيع"]}-${index}`}
                      >
                        {OUTPUT_HEADERS.map((header) => (
                          <td
                            key={header}
                            className="whitespace-nowrap px-3 py-2 text-right"
                          >
                            {row[header] ?? ""}
                          </td>
                        ))}
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>

          <div className="rounded-3xl bg-white p-6 shadow-sm ring-1 ring-slate-200">
            <div className="mb-4 flex items-center gap-3 text-amber-700">
              <AlertCircle className="h-5 w-5" />
              <h2 className="text-lg font-semibold">الأخطاء والملاحظات</h2>
            </div>
            {errors.length === 0 ? (
              <div className="rounded-2xl bg-slate-50 px-4 py-5 text-sm text-slate-500">
                لا توجد أخطاء حالياً.
              </div>
            ) : (
              <div className="space-y-3">
                {errors.map((err, idx) => (
                  <div
                    key={`${err.file}-${idx}`}
                    className="rounded-2xl border border-rose-200 bg-rose-50 px-4 py-3 text-sm"
                  >
                    <div className="font-semibold text-rose-700">
                      {err.file}
                    </div>
                    <div className="mt-1 text-rose-600">{err.message}</div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

function SummaryCard({ label, value }) {
  return (
    <div className="rounded-2xl bg-slate-50 p-4 ring-1 ring-slate-200">
      <div className="text-sm text-slate-500">{label}</div>
      <div className="mt-2 text-2xl font-bold text-slate-900">{value}</div>
    </div>
  );
}
