import { useRef, useState } from "react";
import { ArrowDownToLine, ArrowLeft, Check, CheckCircle2, FileSpreadsheet, FileText, FileUp, LoaderCircle, LockKeyhole, Plus, X } from "lucide-react";
import { completeAgentRows } from "./report-data.js";
import "./App.css";

const fileKey = (file) => `${file.name}|${file.size}|${file.lastModified}`;

export default function App() {
  const inputRef = useRef(null);
  const busyRef = useRef(false);
  const [files, setFiles] = useState([]);
  const [results, setResults] = useState([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isDragging, setIsDragging] = useState(false);
  const [progress, setProgress] = useState(0);
  const [download, setDownload] = useState(null);
  const [message, setMessage] = useState("");

  async function processFiles(nextFiles) {
    if (busyRef.current) return;
    setFiles(nextFiles);
    setDownload(null);
    setResults([]);
    setMessage("");
    setProgress(0);
    if (!nextFiles.length) return;
    busyRef.current = true;
    setIsProcessing(true);
    const rows = [], dates = [], completed = [];
    try {
      const { extractRowsFromPdf } = await import("./pdf-extractor.js");
      for (let index = 0; index < nextFiles.length; index++) {
        const file = nextFiles[index];
        try {
          const report = await extractRowsFromPdf(file);
          rows.push(...report.rows);
          dates.push(...report.reportDates);
          completed.push({ key: fileKey(file), ok: true, dates: report.reportDates.join("، ") });
        } catch (error) {
          const detail = /[\u0600-\u06ff]/.test(error.message ?? "") ? error.message
            : error.name === "PasswordException" ? "الملف محمي بكلمة مرور. ارفع نسخة غير محمية."
              : "تعذر قراءة ملف PDF. استخدم نسخة أصلية تحتوي على نص.";
          completed.push({ key: fileKey(file), ok: false, error: detail });
        }
        setResults([...completed]);
        setProgress(Math.round(((index + 1) / nextFiles.length) * 100));
      }
      if (rows.length) {
        const report = completeAgentRows(rows, dates);
        const { buildWorkbook } = await import("./excel-export.js");
        const buffer = await buildWorkbook(report.rows);
        setDownload({ buffer, count: report.rows.length, days: new Set(dates).size, failed: completed.filter((result) => !result.ok).length });
      } else setMessage("لم نتمكن من استخراج بيانات من الملفات. راجع الملاحظات أدناه وأضف تقارير PDF أصلية.");
    } catch (error) {
      setMessage(error.message || "تعذر تجهيز ملف Excel. حاول مرة أخرى.");
    } finally {
      busyRef.current = false;
      setIsProcessing(false);
    }
  }

  function addFiles(selected) {
    if (busyRef.current) return;
    const valid = Array.from(selected ?? []).filter((file) => /\.pdf$/i.test(file.name));
    if (!valid.length) { setMessage("اختر ملفات PDF فقط."); return; }
    const unique = new Map([...files, ...valid].map((file) => [fileKey(file), file]));
    processFiles([...unique.values()]);
  }

  return (
    <div className="app-shell" dir="rtl">
      <header className="site-header">
        <a className="brand" href="/" aria-label="تقارير وكلاء البيع - الرئيسية">
          <span className="brand-icon"><FileSpreadsheet size={22} strokeWidth={1.8} /></span>
          <span>تقارير<span className="brand-dot">.</span></span>
        </a>
        <span className="privacy-note"><LockKeyhole size={14} /> ملفاتك تبقى على جهازك</span>
      </header>
      <main>
        <div className="intro">
          <span className="eyebrow"><span /> تقارير وكلاء البيع</span>
          <h1>من التقرير،<br className="mobile-break" /> إلى <span dir="ltr">Excel.</span></h1>
          <p>ارفع تقارير PDF وحمّل بياناتك في جدول واحد مرتب.</p>
        </div>
        <section className="converter" aria-label="تحويل تقارير PDF إلى Excel">
          <div className="steps" aria-hidden="true">
            <span className="step"><span className="step-number">١</span> ارفع التقارير</span>
            <span className="step-line" />
            <span className={`step ${download ? "step-ready" : "step-muted"}`}><span className="step-number">٢</span> حمّل Excel</span>
          </div>
          <div className={`drop-zone ${isDragging ? "is-dragging" : ""} ${files.length ? "has-files" : ""}`}
            onDragOver={(event) => { event.preventDefault(); if (!isProcessing) setIsDragging(true); }}
            onDragLeave={(event) => { if (!event.currentTarget.contains(event.relatedTarget)) setIsDragging(false); }}
            onDrop={(event) => { event.preventDefault(); setIsDragging(false); addFiles(event.dataTransfer.files); }}>
            <div className="upload-art" aria-hidden="true"><FileUp size={32} strokeWidth={1.5} /><span className="upload-plus"><Plus size={12} strokeWidth={3} /></span></div>
            <h2>{files.length ? "أضف تقارير أخرى" : "اسحب تقاريرك هنا"}</h2>
            <p>ملفات PDF، تقرير واحد أو أكثر</p>
            <button className="browse-button" type="button" disabled={isProcessing} onClick={() => inputRef.current?.click()}><Plus size={16} /> اختر ملفات PDF</button>
            <input ref={inputRef} className="file-input" type="file" multiple accept=".pdf,application/pdf" aria-label="اختر ملفات PDF" disabled={isProcessing}
              onChange={(event) => { addFiles(event.target.files); event.target.value = ""; }} />
          </div>
          {files.length > 0 && <div className="selected-files">
            <div className="files-heading"><span>التقارير المضافة <b>{files.length}</b></span><button disabled={isProcessing} className="text-button" onClick={() => processFiles([])}>مسح الكل</button></div>
            <ul className="file-list">{files.map((file) => {
              const result = results.find((item) => item.key === fileKey(file));
              return <li className={`file-row ${result && !result.ok ? "file-error" : ""}`} key={fileKey(file)}>
                <span className="file-icon"><FileText size={20} /></span>
                <div className="file-detail"><span className="file-name" dir="auto">{file.name}</span><span className="file-caption">{result ? (result.ok ? <><Check size={12} /> فترة التقرير: <bdi>{result.dates}</bdi></> : result.error) : "في انتظار المعالجة…"}</span></div>
                <button className="remove-button" disabled={isProcessing} aria-label={`إزالة ${file.name}`} onClick={() => processFiles(files.filter((item) => fileKey(item) !== fileKey(file)))}><X size={16} /></button>
              </li>;
            })}</ul>
          </div>}
          <div className="download-area">
            <div className="conversion-status" role="status" aria-live="polite">
              {isProcessing ? <><LoaderCircle size={17} className="spin" /><span>جارٍ تجهيز التقارير… {progress}%</span></>
                : download ? <><CheckCircle2 size={17} /><span>جاهز للتحميل · {download.days} يوم · {download.count} صف</span></>
                  : <span>خطوتان فقط. بدون إعدادات إضافية.</span>}
            </div>
            {isProcessing && <progress className="progress-bar" value={progress} max="100" aria-label="تقدم معالجة التقارير" />}
            <button className="download-button" type="button" disabled={!download || isProcessing}
              onClick={async () => { const { downloadWorkbook } = await import("./excel-export.js"); downloadWorkbook(download.buffer, "تقارير-وكلاء-البيع.xlsx"); }}>
              <span>{isProcessing ? "جارٍ تجهيز الملف…" : "تحميل ملف Excel"}</span><ArrowDownToLine size={20} strokeWidth={1.8} />
            </button>
            {download?.failed > 0 && <p className="warning-message" role="alert">ملف Excel يشمل {results.filter((result) => result.ok).length} تقريرًا ناجحًا فقط. تعذر تحويل {download.failed} ملف؛ راجع الملاحظات بجوار أسمائها.</p>}
            {message && <p className="error-message" role="alert">{message}</p>}
            <span className="download-hint"><FileSpreadsheet size={13} /> ملف واحد بصيغة XLSX</span>
          </div>
        </section>
        <div className="afterword"><span className="afterword-line" /><span>تقارير يومية. بيانات مرتبة. وقت أقل.</span><span className="afterword-line" /></div>
      </main>
      <footer><span>أداة تقارير وكلاء البيع</span><span dir="ltr">PDF <ArrowLeft size={12} /> Excel</span></footer>
    </div>
  );
}
