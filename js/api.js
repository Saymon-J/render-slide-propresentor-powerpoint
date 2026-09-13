// Мост UI ↔ JS-ядро: те же формы запросов/ответов, что были у Flask-эндпоинтов
// (api/parse, api/parse/text, api/relayout, api/export). Вся логика — на устройстве.
import { span, splitParagraphs, parseSermonParagraphs } from "./core.js";
import { initPro, theme, buildPresentation } from "./pro.js";
import { buildPptx } from "./pptx.js";
import { docxParagraphs, txtParagraphs } from "./docx.js";

export class ApiError extends Error {}

let ready = null;
export function init() {
  // одноразовая загрузка схемы protobuf и шаблона темы
  if (!ready) {
    ready = (async () => {
      const [schema, tmpl] = await Promise.all([
        fetch("./js/pro-schema.json").then((r) => r.json()),
        fetch("./assets/theme.pro").then((r) => r.arrayBuffer()),
      ]);
      initPro(schema, new Uint8Array(tmpl));
    })();
  }
  return ready;
}

const parasJson = (paras) => paras.map((par) => par.map((sp) =>
  ({ t: sp.text, b: sp.bold, i: sp.italic, c: sp.color, h: sp.highlight })));

// «Onest-Medium» → «Onest Medium»: превью подставляет имя семейства, как pptx.js
const displayFont = (name) => name.replace(/-Regular$/, "").replace("-", " ");

function sermonToJson(s, layout) {
  const t = theme();
  const join = layout !== "lines";
  const passages = [];
  for (const p of s.passages) {
    for (const chunk of splitParagraphs(p.paragraphs, t.max_lines, t.line_chars, join)) {
      passages.push({ screen: p.screen, label: p.label, paragraphs: parasJson(chunk) });
    }
  }
  return {
    title: s.title,
    theme: { name: t.name, fonts: { ref: displayFont(t.font), body: displayFont(t.font) } },
    passages,
  };
}

function jsonToSermon(d) {
  return {
    title: d.title || "",
    passages: (d.passages || []).map((p) => ({
      screen: p.screen || "",
      label: p.label || "",
      paragraphs: (p.paragraphs || []).map((par) =>
        par.filter((sp) => sp.t).map((sp) => span(sp.t, !!sp.b, !!sp.i, sp.c ?? null, sp.h ?? null))),
    })),
  };
}

const NO_PASSAGES = "В конспекте не найдено ни одного отрывка со ссылкой (например «Ин 3:16»)";

export async function parseFile(file, layout = "flow") {
  await init();
  const name = file.name.toLowerCase();
  let paragraphs;
  if (name.endsWith(".doc")) {
    throw new ApiError("Не удалось открыть .doc. Пересохраните файл как .docx (Word: Файл → Сохранить как)");
  } else if (name.endsWith(".txt")) {
    paragraphs = txtParagraphs(await file.text());
  } else if (name.endsWith(".docx")) {
    paragraphs = await docxParagraphs(await file.arrayBuffer());
  } else {
    throw new ApiError(`Формат ${name.slice(name.lastIndexOf("."))} не поддерживается. Нужен .docx или .txt`);
  }
  const s = parseSermonParagraphs(paragraphs, file.name.replace(/\.[^.]+$/, ""));
  if (!s.passages.length) throw new ApiError(NO_PASSAGES);
  return sermonToJson(s, layout);
}

export async function parseText(payload) {
  await init();
  let paragraphs;
  if (Array.isArray(payload.paragraphs)) {
    paragraphs = payload.paragraphs.map((par) => {
      const spans = par.filter((sp) => sp.t)
        .map((sp) => span(sp.t, !!sp.b, !!sp.i, sp.c ?? null, sp.h ?? null));
      return [spans.map((sp) => sp.text).join(""), spans];
    });
  } else if (typeof payload.text === "string") {
    paragraphs = txtParagraphs(payload.text);
  } else {
    throw new ApiError("Ожидается {text: ...} или {paragraphs: [[...]]}");
  }
  const s = parseSermonParagraphs(paragraphs);
  if (!s.passages.length) throw new ApiError(NO_PASSAGES);
  return sermonToJson(s, payload.layout || "flow");
}

export async function relayout(d) {
  await init();
  const t = theme();
  const join = (d.layout || "flow") !== "lines";
  const s = jsonToSermon(d);
  const out = [];
  for (const p of s.passages) {
    for (const chunk of splitParagraphs(p.paragraphs, t.max_lines, t.line_chars, join)) {
      out.push({ screen: p.screen, label: p.label, paragraphs: parasJson(chunk) });
    }
  }
  if (!out.length) throw new ApiError("Нет ни одного слайда");
  return { passages: out };
}

export async function exportFile(d) {
  await init();
  const s = jsonToSermon(d);
  if (d.format === "pptx") {
    return { blob: await buildPptx(s, theme(), "blob"), ext: "pptx" };
  }
  return {
    blob: new Blob([buildPresentation(s)], { type: "application/octet-stream" }),
    ext: "pro",
  };
}
