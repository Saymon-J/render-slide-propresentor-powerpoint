// Порт parser.py и не-protobuf-части export.py: модель, разбор конспекта,
// раскрой стихов на слайды, RTF writer ProPresenter 7.
// Изменять синхронно с parser.py / export.py (эталон) — проверяется test_core.mjs.

import { parseRef, refForms } from "./books.js";

// Word-маркер (highlight) → RGB
export const WD_HL = {
  BLACK: "000000", BLUE: "0000FF", CYAN: "00FFFF", DARK_BLUE: "000080",
  DARK_GREEN: "006400", DARK_RED: "800000", DARK_YELLOW: "808000",
  GREEN: "00FF00", LIGHT_GRAY: "C0C0C0", MAGENTA: "FF00FF",
  RED: "FF0000", TURQUOISE: "40E0D0", VIOLET: "EE82EE", WHITE: "FFFFFF",
  YELLOW: "FFFF00",
};

// Span — plain-объект {text, bold, italic, color, highlight}
export const span = (text, bold = false, italic = false, color = null, highlight = null) =>
  ({ text, bold, italic, color, highlight });
export const passage = (screen, label, paragraphs = []) => ({ screen, label, paragraphs });
export const sermon = (title, passages = []) => ({ title, passages });

const SEP_RE = /^[\s*•·—–-]+$/;
// абзац начинается с номера стиха («12 Итак, …») — mybible кладёт каждый стих
// с новой строки, в теле отрывка они должны течь одним абзацем
const VERSE_PAR = /^\d{1,3}(\s|$)/;

export function cutSpans(spans, n) {
  // Отрезать первые n символов из последовательности спанов.
  const out = [];
  for (const sp of spans) {
    if (n <= 0) out.push({ ...sp });
    else if (sp.text.length <= n) n -= sp.text.length;
    else { out.push({ ...sp, text: sp.text.slice(n) }); n = 0; }
  }
  while (out.length && !out[0].text.trim()) out.shift();
  if (out.length) out[0] = { ...out[0], text: out[0].text.replace(/^\s+/, "") };
  return out;
}

export function mergeSpans(spans) {
  const out = [];
  for (const sp of spans) {
    if (!sp.text) continue;
    const last = out[out.length - 1];
    if (last && last.bold === sp.bold && last.italic === sp.italic
        && last.color === sp.color && last.highlight === sp.highlight) {
      last.text += sp.text;
    } else out.push({ ...sp });
  }
  return out;
}

export function splitLines(paragraphs) {
  // Переносы строк внутри спанов (\n от <br> в вставленном HTML) режут на абзацы.
  const out = [];
  for (const [text, spans] of paragraphs) {
    if (!spans.some((sp) => sp.text.includes("\n") || sp.text.includes("\r"))) {
      out.push([text, spans]);
      continue;
    }
    let bufTxt = "", bufSpans = [];
    for (const sp of spans) {
      const parts = sp.text.split(/\r\n|\r|\n/);
      parts.forEach((part, i) => {
        if (i) { out.push([bufTxt, bufSpans]); bufTxt = ""; bufSpans = []; }
        bufTxt += part;
        if (part) bufSpans.push({ ...sp, text: part });
      });
    }
    out.push([bufTxt, bufSpans]);
  }
  return out;
}

export function parseSermonParagraphs(paragraphs, fallbackTitle = "") {
  // [(текст абзаца, спаны)] → Sermon. Источник — файл или вставленный текст.
  // Название проповеди — первая непустая строка, если она не ссылка.
  paragraphs = splitLines(paragraphs);
  let title = "", seenFirst = false;
  let cur = null;
  const passages = [];

  for (const [text, spans] of paragraphs) {
    const stripped = text.trim();

    if (!stripped) { cur = null; continue; }  // пустая строка закрывает тело
    if (SEP_RE.test(stripped)) { cur = null; continue; }  // разделитель секции

    const ref = parseRef(text);
    if (!seenFirst) {
      seenFirst = true;
      if (!ref) title = stripped;
    }

    if (ref) {
      const [book, refstr, consumed] = ref;
      // в шапке слайда — полная форма: «Бытие 4:1-11», «2-Паралипоменон 4:1-11»
      const [, label] = refForms(book, refstr);
      cur = passage(label, label);
      passages.push(cur);
      const rest = cutSpans(spans, consumed).filter((sp) => sp.text.trim());
      if (rest.length) {
        cur.paragraphs.push(mergeSpans(rest));
        cur = null;  // однострочный отрывок, тело закрыто
      }
      continue;
    }

    if (cur) {
      const clean = mergeSpans(spans.filter((sp) => sp.text.trim()));
      if (!clean.length) continue;
      if (cur.paragraphs.length && VERSE_PAR.test(clean[0].text)) {
        const prev = cur.paragraphs[cur.paragraphs.length - 1];
        if (!prev[prev.length - 1].text.endsWith(" ")) {
          prev[prev.length - 1] = { ...prev[prev.length - 1], text: prev[prev.length - 1].text + " " };
        }
        prev.push(...clean);
      } else cur.paragraphs.push(clean);
    }
  }

  return sermon(title || fallbackTitle, passages.filter((p) => p.paragraphs.length));
}

// ---------- раскрой стихов (export.py) ----------

// граница стиха внутри текста: знак препинания/кавычка, пробелы (возможно с
// тире реплики NRT), номер стиха; просто «40 дней» границей не считается
const VERSE_END = /(?<=[.,;:!?»])\s+(?:—\s+)?(?=\d{1,3}\s)/g;
const VERSE_START = /^\d{1,3}(\s|$)/;

export function lighten(rgb) {
  // Тёмные цвета конспекта осветлить — иначе нечитаемы на тёмном фоне.
  let c = [0, 2, 4].map((i) => parseInt(rgb.slice(i, i + 2), 16));
  const lum = (t) => 0.2126 * t[0] + 0.7152 * t[1] + 0.0722 * t[2];
  while (lum(c) < 110) {  // ponytail: порог на глаз; дальше — настройка темы
    c = c.map((v) => Math.min(255, v + Math.floor((255 - v) / 3) + 8));
  }
  return c;
}

function verseUnits(para) {
  // Абзац → стихи: режем перед номером стиха после знака препинания внутри
  // спана и перед спаном, начинающимся с номера.
  const units = [], tail = [];
  for (const sp of para) {
    const parts = sp.text.split(VERSE_END);
    parts.forEach((part, i) => {
      if (!part) return;
      if (tail.length && (i > 0 || VERSE_START.test(part))) {
        units.push(tail.splice(0));
      }
      tail.push({ ...sp, text: part });
    });
  }
  if (tail.length) units.push(tail);
  return units.length ? units : [para];
}

function flushChunk(items, joinVerses = true) {
  // Стихи одного слайда → абзацы. joinVerses: стихи текут одним абзацем,
  // иначе каждый стих — отдельный абзац.
  if (!joinVerses) return items.map(([, unit]) => unit);
  const paras = [];
  let curPi = -1, cur = [];
  for (const [pi, unit] of items) {
    const startsVerse = cur.length > 0 && VERSE_START.test(unit[0].text);
    if (cur.length && pi !== curPi && !startsVerse) {
      paras.push(cur); cur = [];
    } else if (cur.length && !cur[cur.length - 1].text.endsWith(" ")) {
      cur[cur.length - 1] = { ...cur[cur.length - 1], text: cur[cur.length - 1].text + " " };
    }
    curPi = pi;
    cur.push(...unit);
  }
  if (cur.length) paras.push(cur);
  return paras;
}

export function splitParagraphs(paragraphs, maxLines, lineChars, joinVerses = true) {
  // Разбить тело отрывка на порции, не превышающие экран (~4-5 стихов);
  // длинный абзац режется по номерам стихов.
  const out = [];
  let cur = [], used = 0;
  paragraphs.forEach((para, pi) => {
    for (const unit of verseUnits(para)) {
      const n = Math.max(1, Math.ceil(unit.reduce((a, sp) => a + sp.text.length, 0) / lineChars));
      if (cur.length && used + n > maxLines) {
        out.push(flushChunk(cur, joinVerses));
        cur = []; used = 0;
      }
      cur.push([pi, unit]);
      used += n;
    }
  });
  if (cur.length) out.push(flushChunk(cur, joinVerses));
  return out;
}

// ---------- RTF (export.py) ----------

function esc(text) {
  let out = "";
  for (const ch of text) {
    const cp = ch.codePointAt(0);
    if ("{}\\".includes(ch)) out += "\\" + ch;
    else if (cp < 128) out += ch;
    else out += `\\u${cp} ?`;
  }
  return out;
}

const hexToPct = (rgb) => rgb.map((v) => `\\c${Math.round(v * 100000 / 255)}`).join("");
const rgbEq = (a, b) => a[0] === b[0] && a[1] === b[1] && a[2] === b[2];
const WHITE = [255, 255, 255];

export function buildRtf(paragraphs, sizePt, align = "left", baseBold = false, font, theme) {
  // RTF в формате writer'а ProPresenter 7 (кириллица через \uNNNN ?).
  font = font || theme.font;
  const colors = [WHITE];  // RTF-индекс 1: белый текст
  const idx = new Map();

  const colorId = (sp) => (sp.color ? lighten(sp.color) : WHITE);
  const hlId = (sp) => {
    if (!sp.highlight) return 0;  // 0 = без фона; PP игнорирует альфу в expandedcolortbl
    const key = `hl|${sp.highlight}`;
    if (!idx.has(key)) {
      colors.push(lighten(sp.highlight));
      idx.set(key, colors.length);  // RTF-индекс: colors[0] → 1
    }
    return idx.get(key);
  };
  const cfId = (sp) => {
    const c = colorId(sp);
    if (rgbEq(c, WHITE)) return 1;
    const key = `cf|${c.join(",")}`;
    if (!idx.has(key)) {
      colors.push(c);
      idx.set(key, colors.length);
    }
    return idx.get(key);
  };

  const fs = sizePt * 2;
  const par = "\\pard\\li0\\fi0\\ri0\\" + (align === "center" ? "qc" : "ql");
  // пробел в конце head и после каждой числовой команды — обязательный
  // разделитель RTF: «\b1» + текст «2-…» иначе склеивается в «\b12» и ест цифру
  const head = ("\\sb0\\sa0\\sl240\\slmult1\\slleading0\\f0\\b0\\i0\\ul0\\strike0"
                + `\\fs${fs}\\expnd0\\expndtw0\\cf1\\strokewidth0\\strokec1\\nosupersub\\ulc0 `);

  // сначала тело — оно наполняет colortbl цветами спанов, потом преамбула
  const body = [];
  paragraphs.forEach((para, pi) => {
    const seg = [`${par}${head}`];
    const cur = { b: false, i: false, cf: 1, hl: 0 };
    for (const sp of para) {
      const b = sp.bold || baseBold, i = sp.italic;
      const cf = cfId(sp), hl = hlId(sp);
      if (b !== cur.b) { seg.push(`\\b${b ? 1 : 0} `); cur.b = b; }
      if (i !== cur.i) { seg.push(`\\i${i ? 1 : 0} `); cur.i = i; }
      if (cf !== cur.cf) { seg.push(`\\cf${cf} `); cur.cf = cf; }
      if (hl !== cur.hl) { seg.push(`\\highlight${hl}\\cb${hl} `); cur.hl = hl; }
      seg.push(esc(sp.text));
    }
    if (pi < paragraphs.length - 1) seg.push("\\par");
    body.push(seg.join(""));
  });

  const colortbl = colors.map(([r, g, b]) => `\\red${r}\\green${g}\\blue${b}`).join(";");
  const expanded = colors.map((c) => "\\csgenericrgb" + hexToPct(c) + "\\c100000").join("");
  const parts = [
    `{\\rtf0\\ansi\\ansicpg1252{\\fonttbl\\f0\\fnil ${font};}`,
    "{\\colortbl;" + colortbl + ";}",
    "{\\*\\expandedcolortbl" + expanded + ";}",
    "{\\*\\listtable}{\\*\\listoverridetable}\\uc1\\paperw36000\\margl0\\margr0\\margt0\\margb0",
    ...body,
    "}",
  ];
  return parts.join("");
}
