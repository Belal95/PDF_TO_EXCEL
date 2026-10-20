import * as pdfjsLib from "pdfjs-dist";
import workerSrc from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import { extractRowsFromDocument } from "./report-data.js";

pdfjsLib.GlobalWorkerOptions.workerSrc = workerSrc;

export async function extractRowsFromPdf(file) {
  const task = pdfjsLib.getDocument({ data: await file.arrayBuffer() });
  try { return await extractRowsFromDocument(await task.promise); }
  finally { await task.destroy(); }
}
