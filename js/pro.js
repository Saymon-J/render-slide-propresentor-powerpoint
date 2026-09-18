// Порт theme.py + protobuf-части export.py: тема из шаблона .pro,
// сборка .pro (ProPresenter 7). Глобал protobuf — из vendor/protobuf.min.js.
// Изменять синхронно с export.py (эталон) — проверяется test_core.mjs.

import { span, buildRtf } from "./core.js";

let root = null;      // protobufjs Root (после initPro)
let THEME = null;     // параметры темы (порт theme.CURRENT без слайдов-сообщений)
let DEF = null;       // параметры spbcoc по умолчанию — снимок шаблона

const T = () => {
  if (!THEME) throw new Error("initPro() не вызван");
  return THEME;
};

const clone = (Type, msg) => Type.decode(Type.encode(msg).finish());

export function initPro(schemaJson, templateBytes) {
  // schemaJson — содержимое js/pro-schema.json; templateBytes — шаблон .pro (Uint8Array)
  root = protobuf.Root.fromJSON(schemaJson);
  const Presentation = root.lookupType("rv.data.Presentation");
  const p = Presentation.decode(templateBytes);
  const slides = p.cues.map((c) => c.actions[0].slide.presentation.base_slide);

  const box = (el) => {
    const b = el.bounds;
    return [b.origin.x, b.origin.y, b.size.width, b.size.height].map(Math.trunc);
  };
  const titleEl = slides[0].elements[0].element;
  const [refEl, bodyEl] = slides[1].elements.map((e) => e.element)
    .sort((a, b) => a.bounds.origin.y - b.bounds.origin.y);
  const bg = slides[1].background_color;
  THEME = {
    name: "spbcoc",
    // слайды-шаблоны для .pro-экспорта (клонируются с заменой текста)
    title_slide: slides[0],
    ref_slide: slides[1],
    // производные параметры для RTF/pptx/превью
    title_font: titleEl.text.attributes.font.name,
    title_size: titleEl.text.attributes.font.size,
    title_box: box(titleEl),
    font: refEl.text.attributes.font.name,
    ref_size: refEl.text.attributes.font.size,
    ref_box: box(refEl),
    body_font: bodyEl.text.attributes.font.name,
    body_size: bodyEl.text.attributes.font.size,
    body_box: box(bodyEl),
    // пункты конспекта: слайд-заголовок тем же боксом, что титул, но
    // спокойнее — Onest-Medium 96 (титульный Forum 200 слишком крупный)
    point_font: "Onest-Medium",
    point_size: 96,
    slide_bg: [bg.red, bg.green, bg.blue],
    group_name: "Отрывки",
    // дробление длинного отрывка: строк «высотой» ~line_chars символов
    line_chars: 42,
    max_lines: 12,
  };
  DEF = {
    title_font: THEME.title_font, title_size: THEME.title_size,
    font: THEME.font, ref_size: THEME.ref_size,
    body_font: THEME.body_font, body_size: THEME.body_size,
    point_font: THEME.point_font, point_size: THEME.point_size,
    line_chars: THEME.line_chars, max_lines: THEME.max_lines,
  };
  return THEME;
}

export function applyTheme(o = {}) {
  // Тема пользователя: шрифты/размеры (пустое — вернуть spbcoc). Дробление
  // пересчитывается линейно под размер тела — ponytail: линейная прикидка,
  // не точная метрика шрифта под бокс темы
  const t = T();
  const clean = (s) => (s || "").replace(/[;{}\\]/g, "").trim();  // не ломать RTF fonttbl
  t.title_font = clean(o.title_font) || DEF.title_font;
  t.title_size = +o.title_size || DEF.title_size;
  t.font = clean(o.font) || DEF.font;
  t.ref_size = +o.ref_size || DEF.ref_size;
  t.body_font = clean(o.body_font) || DEF.body_font;
  t.body_size = +o.body_size || DEF.body_size;
  t.point_font = clean(o.point_font) || DEF.point_font;
  t.point_size = +o.point_size || DEF.point_size;
  t.line_chars = Math.max(8, Math.round(DEF.line_chars * DEF.body_size / t.body_size));
  t.max_lines = Math.max(2, Math.round(DEF.max_lines * DEF.body_size / t.body_size));
  return t;
}

export const theme = T;

export const defaultUid = () => ({ string: crypto.randomUUID() });

function slideFrom(tmpl, items, uid) {
  // Клон слайда-шаблона темы с заменой текста (элементы сверху вниз);
  // items: [{rtf, font, size}] — атрибуты элемента синхронны с RTF
  const PS = root.lookupType("rv.data.PresentationSlide");
  const Slide = root.lookupType("rv.data.Slide");
  const Element = root.lookupType("rv.data.Slide.Element");
  const ps = PS.create();
  ps.base_slide = clone(Slide, tmpl);
  ps.base_slide.uuid = uid();
  const wrapped = [...ps.base_slide.elements]
    .sort((a, b) => a.element.bounds.origin.y - b.element.bounds.origin.y);
  ps.base_slide.elements = [];
  const enc = new TextEncoder();
  for (const w of wrapped) {
    const cell = clone(Element, w);
    cell.element.uuid = uid();
    if (items.length) {
      const it = items.shift();
      cell.element.text.rtf_data = enc.encode(it.rtf);
      cell.element.text.attributes.font.name = it.font;
      cell.element.text.attributes.font.size = it.size;
    }
    ps.base_slide.elements.push(cell);
  }
  ps.notes = { rtf_data: enc.encode("{\\rtf0\\ansi\\ansicpg1252}") };
  return ps;
}

export function buildPresentation(sermon, uid = defaultUid) {
  const t = T();
  const PP = root.lookupType("rv.data.Presentation");
  const Cue = root.lookupType("rv.data.Cue");
  const enumVal = (name, key) => root.lookupEnum(name).values[key];

  const p = PP.create();
  p.application_info = {
    platform: enumVal("rv.data.ApplicationInfo.Platform", "PLATFORM_WINDOWS"),
    application: enumVal("rv.data.ApplicationInfo.Application", "APPLICATION_PROPRESENTER"),
    application_version: { major_version: 7, minor_version: 8, patch_version: 2 },
  };
  p.uuid = uid();
  p.name = sermon.title;
  p.background = { color: { red: t.slide_bg[0], green: t.slide_bg[1], blue: t.slide_bg[2], alpha: 1.0 } };

  const cues = [];
  const cueIds = [];
  const addCue = (label, slide) => {
    const cue = Cue.create();
    cue.uuid = uid();
    cue.isEnabled = true;  // без этого cue по умолчанию false — слайды в PP7 выключены
    cue.completion_target_uuid = { string: "00000000-0000-0000-0000-000000000000" };
    cue.completion_action_type = enumVal("rv.data.Cue.CompletionActionType", "COMPLETION_ACTION_TYPE_LAST");
    cue.completion_action_uuid = { string: "00000000-0000-0000-0000-000000000000" };
    cue.actions = [{
      uuid: uid(),
      isEnabled: true,
      type: enumVal("rv.data.Action.ActionType", "ACTION_TYPE_PRESENTATION_SLIDE"),
      label: { text: label },
      slide: { presentation: slide },
    }];
    cues.push(cue);
    cueIds.push(cue.uuid);
  };

  // заголовок капсом текстом: атрибут Capitalization=ALL_CAPS PP при импорте
  // не отрисовывает, рендер идёт по RTF
  const titleRtf = buildRtf([[span(sermon.title.toUpperCase())]], t.title_size, "center", true, t.title_font, t);
  addCue(sermon.title, slideFrom(t.title_slide,
    [{ rtf: titleRtf, font: t.title_font, size: t.title_size }], uid));
  for (const passage of sermon.passages) {
    if (passage.point) {
      // пункт конспекта — слайд-заголовок в боксе титула, шрифт пункта
      const rtf = buildRtf([[span(passage.label)]], t.point_size, "center", false, t.point_font, t);
      addCue(passage.label, slideFrom(t.title_slide,
        [{ rtf, font: t.point_font, size: t.point_size }], uid));
      continue;
    }
    const refRtf = buildRtf([[span(passage.screen)]], t.ref_size, "left", false, t.font, t);  // координаты — начертанием шрифта, не жирным
    const bodyRtf = buildRtf(passage.paragraphs, t.body_size, "left", false, t.body_font, t);
    addCue(passage.label, slideFrom(t.ref_slide, [
      { rtf: refRtf, font: t.font, size: t.ref_size },
      { rtf: bodyRtf, font: t.body_font, size: t.body_size },
    ], uid));
  }

  p.cues = cues;
  p.cue_groups = [{
    group: { uuid: uid(), name: t.group_name },
    cue_identifiers: cueIds,
  }];
  return PP.encode(p).finish();
}
