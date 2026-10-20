export const AGENTS = ["فوري", "خالص", "ضامن", "الاهلي"];
export const OUTPUT_HEADERS = [
  "التاريخ", "وكيل البيع", "عدد التذاكر", "عمولة وكيل البيع",
  "عمولة ترانس اي تي", "عمولة الهيئة", "ايرادات التذاكر",
  "اجمالي ايراد الهيئة", "الاجمالي",
];
const DATE_TOKEN = String.raw`(?:\d{4}[/.-]\d{1,2}[/.-]\d{1,2}|\d{1,2}[/.-]\d{1,2}[/.-]\d{4})`;

export function normalizeArabicDigits(value) {
  return String(value ?? "")
    .replace(/[٠-٩]/g, (digit) => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit)))
    .replace(/[۰-۹]/g, (digit) => String("۰۱۲۳۴۵۶۷۸۹".indexOf(digit)));
}
export function normalizeArabicText(value) {
  return normalizeArabicDigits(value).normalize("NFKC")
    .replace(/[\u200e\u200f\u202a-\u202e\u2066-\u2069ـ]/g, "")
    .replace(/[إأآ]/g, "ا").replace(/ة/g, "ه").replace(/ى/g, "ي")
    .replace(/\s+/g, " ").trim();
}
export function normalizeAgentName(value) {
  const text = normalizeArabicText(value);
  return AGENTS.find((agent) => text.includes(agent)) ?? null;
}
export function standardizeDate(value) {
  const normalized = normalizeArabicDigits(value).trim();
  const parts = normalized.split(/[/.-]/).map(Number);
  if (parts.length !== 3) throw new Error("تاريخ التقرير غير صالح.");
  const [year, month, day] = /^\d{4}[/.-]/.test(normalized) ? parts : [parts[2], parts[1], parts[0]];
  const date = new Date(Date.UTC(year, month - 1, day));
  if (year < 1900 || year > 2200 || date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) throw new Error("تاريخ التقرير غير صالح.");
  return `${String(day).padStart(2, "0")}/${String(month).padStart(2, "0")}/${year}`;
}
function dateKey(value) { return value.split("/").reverse().join("-"); }
export function sortRows(rows) {
  return [...rows].sort((a, b) => dateKey(a["التاريخ"]).localeCompare(dateKey(b["التاريخ"]))
    || AGENTS.indexOf(a["وكيل البيع"]) - AGENTS.indexOf(b["وكيل البيع"]));
}
export function reconstructRowsFromItems(items) {
  const rows = [];
  for (const item of items) {
    const str = String(item.str ?? "").trim();
    if (!str) continue;
    const x = item.transform?.[4] ?? 0;
    const y = item.transform?.[5] ?? 0;
    let row = rows.find((candidate) => Math.abs(candidate.y - y) <= 2.5);
    if (!row) { row = { y, parts: [] }; rows.push(row); }
    row.parts.push({ x, str, normalized: normalizeArabicText(str) });
  }
  return rows.sort((a, b) => b.y - a.y);
}
function rowText(row) {
  return normalizeArabicText([...row.parts].sort((a, b) => b.x - a.x).map((part) => part.str).join(" "));
}
// Only the labelled report interval is authoritative; the header is a creation date.
export function extractReportDate(rowGroups) {
  const periodRows = rowGroups.filter((row) => /(?:من|الي)\s*تاريخ/.test(rowText(row)));
  if (!periodRows.length) return null;
  const dates = rowGroups.flatMap((row) => row.parts.flatMap((part) => {
    const match = part.normalized.match(new RegExp(DATE_TOKEN));
    return match ? [{ x: part.x, y: row.y, value: match[0] }] : [];
  }));
  function labelledDate(label) {
    for (const row of periodRows) {
      // Some PDFs place the label and date in a single text item/line.
      const inline = rowText(row).match(new RegExp(`${label}\\s*[:：]?\\s*(${DATE_TOKEN})`));
      if (inline) return standardizeDate(inline[1]);
      const parts = [...row.parts].sort((a, b) => b.x - a.x);
      for (let index = 0; index < parts.length; index++) {
        for (let length = 1; length <= 3; length++) {
          if (!new RegExp(`^${label}`).test(parts.slice(index, index + length).map((part) => part.normalized).join(" "))) continue;
          // Arabic labels sit to the right of their value. Their baselines can
          // differ, so associate nearby coordinates instead of taking a header date.
          const nearby = dates.filter((date) => Math.abs(date.y - row.y) <= 12 && date.x < parts[index].x && parts[index].x - date.x <= 250)
            .sort((a, b) => Math.abs(a.y - row.y) - Math.abs(b.y - row.y) || b.x - a.x);
          if (nearby.length) return standardizeDate(nearby[0].value);
        }
      }
    }
    return null;
  }
  const start = labelledDate(String.raw`من\s*تاريخ`);
  const end = labelledDate(String.raw`الي\s*تاريخ`);
  if (!start || !end) throw new Error("تعذر قراءة «من تاريخ» و«إلى تاريخ» في شروط التقرير. استخدم تقرير PDF اليومي الأصلي.");
  if (start !== end) throw new Error("هذا التقرير يغطي أكثر من يوم. ارفع تقريرًا يوميًا حتى تكون قيمة التاريخ دقيقة.");
  return start;
}
function parseNumber(value) {
  const text = normalizeArabicDigits(value).replace(/[٬,\s\u200e\u200f]/g, "").replace(/٫/g, ".");
  if (!/^-?\d+(?:\.\d+)?$/.test(text) || !Number.isFinite(Number(text))) throw new Error("تعذر قراءة إحدى القيم الرقمية في صف الوكيل.");
  return Number(text);
}
export function extractPageRows(rowGroups, reportDate) {
  const rows = [];
  for (const row of rowGroups) {
    if (/اجمالي\s*(?:الفتره|التقرير)/.test(rowText(row))) continue;
    const agentPart = [...row.parts].sort((a, b) => b.x - a.x).find((part) => normalizeAgentName(part.str));
    if (!agentPart) continue;
    const numericParts = row.parts.filter((part) => /\d/.test(normalizeArabicDigits(part.str)) && part.x < agentPart.x - 5).sort((a, b) => b.x - a.x);
    if (numericParts.length !== 7) throw new Error(`تعذر قراءة الأعمدة السبعة للوكيل ${normalizeAgentName(agentPart.str)}. لم يتم استبدال قيمه بأصفار.`);
    const record = { التاريخ: reportDate, "وكيل البيع": normalizeAgentName(agentPart.str) };
    numericParts.forEach((part, index) => { record[OUTPUT_HEADERS[index + 2]] = parseNumber(part.str); });
    rows.push(record);
  }
  return rows;
}
export async function extractRowsFromDocument(pdfDocument) {
  const rows = [];
  const reportDates = new Set();
  let reportDate = null;
  for (let pageNumber = 1; pageNumber <= pdfDocument.numPages; pageNumber++) {
    const page = await pdfDocument.getPage(pageNumber);
    const content = await page.getTextContent();
    const groups = reconstructRowsFromItems(content.items);
    const pageDate = extractReportDate(groups);
    if (pageDate) reportDate = pageDate;
    if (!reportDate) throw new Error("تعذر العثور على تاريخ فترة التقرير. لن يتم استخدام تاريخ إنشاء الملف.");
    const pageRows = extractPageRows(groups, reportDate);
    rows.push(...pageRows);
    if (pageRows.length) reportDates.add(reportDate);
    page.cleanup?.();
  }
  if (!rows.length) throw new Error("لم يتم العثور على صفوف للوكلاء. تأكد من رفع تقرير PDF أصلي يحتوي على نص.");
  return { rows, reportDates: [...reportDates] };
}
export function completeAgentRows(rows, reportDates) {
  const byKey = new Map();
  let duplicateCount = 0;
  for (const row of rows) {
    const key = `${row["التاريخ"]}|${row["وكيل البيع"]}`;
    const existing = byKey.get(key);
    if (existing) {
      if (OUTPUT_HEADERS.slice(2).some((header) => existing[header] !== row[header])) throw new Error(`توجد قيم مختلفة للوكيل ${row["وكيل البيع"]} بتاريخ ${row["التاريخ"]}. أزل التقرير المكرر المختلف ثم أعد المحاولة.`);
      duplicateCount++;
    } else byKey.set(key, row);
  }
  let addedZeroRows = 0;
  for (const reportDate of new Set(reportDates)) {
    for (const agent of AGENTS) {
      const key = `${reportDate}|${agent}`;
      if (!byKey.has(key)) {
        byKey.set(key, { التاريخ: reportDate, "وكيل البيع": agent, ...Object.fromEntries(OUTPUT_HEADERS.slice(2).map((header) => [header, 0])) });
        addedZeroRows++;
      }
    }
  }
  return { rows: sortRows([...byKey.values()]), addedZeroRows, duplicateCount };
}
