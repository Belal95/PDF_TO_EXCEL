import test from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "exceljs";
import { AGENTS, OUTPUT_HEADERS, completeAgentRows, extractPageRows, extractReportDate, extractRowsFromDocument, reconstructRowsFromItems, sortRows, standardizeDate } from "../src/report-data.js";
import { buildWorkbook } from "../src/excel-export.js";

const item = (str, x, y) => ({ str, transform: [1, 0, 0, 1, x, y] });
const interval = (from = "٢٠٢٦/٠٩/٣٠", to = from) => [
  item("التاريخ: ٢٠٢٦/١٠/٠١", 530, 780),
  item("من تاريخ:", 503, 660), item("الى تاريخ:", 174, 660),
  item(from, 386, 655), item(to, 58, 655),
];
const agentItems = (agent = "خالص – اي فاينانس", y = 500) => [
  item(agent, 530, y),
  ...["١٢", "١٠٫٥٠", "٢٫٥٠", "٢٠٫٠٠", "١٬٠٠٠٫٠٠", "١٠٢٠٫٠٠", "١٠٣٠٫٥٠"].map((value, index) => item(value, 460 - index * 65, y)),
];
const sourceRow = (agent = "خالص", date = "30/09/2026") => ({ التاريخ: date, "وكيل البيع": agent, ...Object.fromEntries(OUTPUT_HEADERS.slice(2).map((header, index) => [header, index + 1])) });
const document = (pages) => ({ numPages: pages.length, getPage: async (number) => ({ getTextContent: async () => ({ items: pages[number - 1] }) }) });

test("uses report-period coordinates when labels and dates have different baselines", () => {
  assert.equal(extractReportDate(reconstructRowsFromItems(interval())), "30/09/2026");
});
test("supports inline date labels and ignores creation dates", () => {
  const groups = reconstructRowsFromItems([
    item("التاريخ: 2026/09/01", 0, 800),
    item("من تاريخ: 2026/09/30 إلى تاريخ: 2026/09/30", 0, 600),
  ]);
  assert.equal(extractReportDate(groups), "30/09/2026");
});
test("never falls back to a creation date when the period is missing", async () => {
  assert.equal(extractReportDate(reconstructRowsFromItems([item("التاريخ: 2026/10/01", 0, 800)])), null);
  await assert.rejects(extractRowsFromDocument(document([[item("التاريخ: 2026/10/01", 0, 800), ...agentItems()]])), /تاريخ فترة التقرير/);
});
test("rejects multi-day aggregate reports and incomplete period labels", () => {
  assert.throws(() => extractReportDate(reconstructRowsFromItems(interval("2026/09/29", "2026/09/30"))), /أكثر من يوم/);
  assert.throws(() => extractReportDate(reconstructRowsFromItems([item("من تاريخ:", 500, 660), item("التاريخ: 2026/10/01", 400, 800)])), /تعذر قراءة/);
});
test("normalizes Arabic dates and validates calendar days", () => {
  assert.equal(standardizeDate("۲۰۲۶/۰۹/۳۰"), "30/09/2026");
  assert.equal(standardizeDate("30-09-2026"), "30/09/2026");
  assert.throws(() => standardizeDate("2026/09/31"), /غير صالح/);
});
test("preserves parsed values and fills all seven missing-agent values with numeric zeros", () => {
  const parsed = extractPageRows(reconstructRowsFromItems(agentItems()), "30/09/2026");
  assert.equal(parsed[0]["عدد التذاكر"], 12);
  assert.equal(parsed[0]["ايرادات التذاكر"], 1000);
  assert.equal(parsed[0]["عمولة وكيل البيع"], 10.5);
  const result = completeAgentRows(parsed, ["30/09/2026"]);
  assert.equal(result.rows.length, 4);
  assert.deepEqual(result.rows.map((row) => row["وكيل البيع"]), AGENTS);
  for (const row of result.rows.filter((row) => row["وكيل البيع"] !== "خالص")) {
    assert.ok(OUTPUT_HEADERS.slice(2).every((header) => row[header] === 0));
  }
  assert.equal(result.rows[1]["ايرادات التذاكر"], 1000);
});
test("combines genuine rows before padding, and ignores identical duplicates", () => {
  const rows = [sourceRow("خالص"), sourceRow("خالص"), sourceRow("فوري")];
  const result = completeAgentRows(rows, ["30/09/2026", "30/09/2026"]);
  assert.equal(result.rows.length, 4);
  assert.equal(result.rows[0]["عدد التذاكر"], 1);
  assert.equal(result.addedZeroRows, 2);
  assert.equal(result.duplicateCount, 1);
});
test("does not silently choose between conflicting duplicates", () => {
  assert.throws(() => completeAgentRows([sourceRow(), { ...sourceRow(), "عدد التذاكر": 999 }], ["30/09/2026"]), /قيم مختلفة/);
});
test("sorts by calendar date across month and year boundaries", () => {
  const rows = sortRows([sourceRow("خالص", "01/01/2027"), sourceRow("خالص", "01/10/2026"), sourceRow("خالص", "30/09/2026")]);
  assert.deepEqual(rows.map((row) => row["التاريخ"]), ["30/09/2026", "01/10/2026", "01/01/2027"]);
});
test("does not hide malformed agent rows behind zero rows", () => {
  assert.throws(() => extractPageRows(reconstructRowsFromItems(agentItems().slice(0, -1)), "30/09/2026"), /لم يتم استبدال قيمه بأصفار/);
});
test("reads continuation pages, excludes totals, and rejects unextractable PDFs", async () => {
  const result = await extractRowsFromDocument(document([
    [...interval(), ...agentItems()],
    [...agentItems("ضامن"), ...agentItems("اجمالي الفتره", 400)],
  ]));
  assert.equal(result.rows.length, 2);
  assert.ok(result.rows.every((row) => row["التاريخ"] === "30/09/2026"));
  await assert.rejects(extractRowsFromDocument(document([interval()])), /لم يتم العثور على صفوف/);
});
test("serialized Excel has four rows, correct dates, numeric zeros, and original columns", async () => {
  const rows = completeAgentRows([sourceRow()], ["30/09/2026"]).rows;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await buildWorkbook(rows));
  const sheet = workbook.getWorksheet("data");
  assert.equal(sheet.rowCount, 5);
  assert.deepEqual(sheet.getRow(1).values.slice(1), OUTPUT_HEADERS);
  assert.equal(sheet.getCell("A2").value, "30/09/2026");
  assert.equal(sheet.getCell("B2").value, "فوري");
  for (let column = 3; column <= 9; column++) assert.equal(sheet.getRow(2).getCell(column).value, 0);
  assert.equal(sheet.views[0].rightToLeft, true);
});
