import ExcelJS from "exceljs";
import { OUTPUT_HEADERS, sortRows } from "./report-data.js";

export async function buildWorkbook(rows) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "ENR Reports";
  const sheet = workbook.addWorksheet("data", {
    views: [{ rightToLeft: true, state: "frozen", ySplit: 1 }],
  });
  sheet.properties.defaultRowHeight = 23;
  sheet.columns = OUTPUT_HEADERS.map((header, index) => ({
    header, key: header, width: index < 2 ? 17 : Math.max(18, header.length + 3),
    style: { font: { name: "Arial", size: 11 }, alignment: { horizontal: "center", vertical: "middle" } },
  }));
  sortRows(rows).forEach((record) => {
    const row = sheet.addRow(OUTPUT_HEADERS.map((header) => record[header]));
    if (row.number % 2 === 0) row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF4F6FA" } };
  });
  const header = sheet.getRow(1);
  header.height = 34;
  header.font = { name: "Arial", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
  header.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF182A43" } };
  header.alignment = { horizontal: "center", vertical: "middle", wrapText: true, readingOrder: "rtl" };
  OUTPUT_HEADERS.slice(2).forEach((key) => { sheet.getColumn(key).numFmt = key === "عدد التذاكر" ? "#,##0" : "#,##0.00"; });
  sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: OUTPUT_HEADERS.length } };
  return workbook.xlsx.writeBuffer();
}
export function downloadWorkbook(buffer, fileName) {
  const url = URL.createObjectURL(new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
