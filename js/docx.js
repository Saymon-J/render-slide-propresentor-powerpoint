// Чтение .docx/.txt → [(текст абзаца, спаны)] — порт _docx_paragraphs/_txt_paragraphs
// parser.py. Глобал JSZip — из vendor/jszip.min.js.
// ponytail: document.xml разбирается регулярками (Word-файлы регулярны, DOMParser
// в node отсутствует); потолок — экзотические rPr-структуры вне Word-генерации.

import { span, WD_HL } from "./core.js";

function xmlUnescape(s) {
  return s.replace(/&(amp|lt|gt|quot|apos|#x?[0-9a-fA-F]+);/g, (m, e) => {
    if (e === "amp") return "&";
    if (e === "lt") return "<";
    if (e === "gt") return ">";
    if (e === "quot") return '"';
    if (e === "apos") return "'";
    const hex = e[1] === "x" || e[1] === "X";
    return String.fromCodePoint(parseInt(e.slice(hex ? 2 : 1), hex ? 16 : 10));
  });
}

// содержимое рана: w:t + w:tab(\t) + w:br/w:cr(\n), как r.text у python-docx
function runText(rXml) {
  let out = "";
  const re = /<w:(t|tab|br|cr)\b[^>]*?(?:\/>|>([\s\S]*?)<\/w:\1>)/g;
  let m;
  while ((m = re.exec(rXml))) {
    if (m[1] === "t") out += xmlUnescape(m[2] ?? "");
    else if (m[1] === "tab") out += "\t";
    else out += "\n";
  }
  return out;
}

const ON = new Set([undefined, "", "1", "true", "on"]);
// <w:b/> — вкл; <w:b w:val="0|false|off"/> — выкл (как r.bold у python-docx)
const flag = (rPr, tag) => {
  const m = rPr && rPr.match(new RegExp(`<w:${tag}(?:\\s+w:val="([^"]*)")?\\s*/?>`));
  return !!m && ON.has(m[1]);
};

function runColor(rPr) {
  const m = rPr && rPr.match(/<w:color w:val="([0-9A-Fa-f]{6})"\s*\/?>/);
  return m ? m[1].toUpperCase() : null;  // theme/auto → нет цвета, как в parser.py
}

function runHighlight(rPr) {
  const m = rPr && rPr.match(/<w:highlight w:val="(\w+)"\s*\/?>/);
  return m ? (WD_HL[m[1]] ?? null) : null;
}

export async function docxParagraphs(bytes) {
  // bytes: ArrayBuffer|Uint8Array документа .docx → [(текст, [Span])]
  const zip = await JSZip.loadAsync(bytes);
  const xml = await zip.file("word/document.xml").async("string");
  const out = [];
  const paras = xml.match(/<w:p\b[\s\S]*?<\/w:p>|<w:p\b[^>]*\/>/g) || [];
  for (const p of paras) {
    const spans = [];
    const runs = p.match(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>|<w:r(?:\s[^>]*)?\/>/g) || [];
    for (const r of runs) {
      const text = runText(r);
      if (!text) continue;
      const rPr = (r.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [])[0];
      spans.push(span(text, flag(rPr, "b"), flag(rPr, "i"), runColor(rPr), runHighlight(rPr)));
    }
    out.push([spans.map((s) => s.text).join(""), spans]);
  }
  return out;
}

// python str.splitlines: \r\n \n \r \v \f \x1c \x1d \x1e \x85 \u2028 \u2029
const LINE_SPLIT = /\r\n|[\n\r\v\f\x1c\x1d\x1e\x85\u2028\u2029]/;

export function txtParagraphs(text) {
  // .txt/вставленный текст: строка → [(строка, [Span])], BOM срезается (utf-8-sig)
  return (text[0] === "\uFEFF" ? text.slice(1) : text).split(LINE_SPLIT)
    .map((line) => [line, line.trim() ? [span(line)] : []]);
}
