// Порт theme.py + protobuf-части export.py: тема из шаблона .pro,
// сборка .pro (ProPresenter 7). Глобал protobuf — из vendor/protobuf.min.js.
// Изменять синхронно с export.py (эталон) — проверяется test_core.mjs.

import { span, buildRtf } from "./core.js";

let root = null;      // protobufjs Root (после initPro)
let THEME = null;     // параметры темы (порт theme.CURRENT без слайдов-сообщений)

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
    body_size: bodyEl.text.attributes.font.size,
    body_box: box(bodyEl),
    slide_bg: [bg.red, bg.green, bg.blue],
    group_name: "Отрывки",
    // дробление длинного отрывка: строк «высотой» ~line_chars символов
    line_chars: 42,
    max_lines: 12,
  };
  return THEME;
}

export const theme = T;

export const defaultUid = () => ({ string: crypto.randomUUID() });

function slideFrom(tmpl, texts, uid) {
  // Клон слайда-шаблона темы с заменой текста (элементы сверху вниз).
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
    if (texts.length) cell.element.text.rtf_data = enc.encode(texts.shift());
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

  const titleRtf = buildRtf([[span(sermon.title)]], t.title_size, "center", true, t.title_font, t);
  addCue(sermon.title, slideFrom(t.title_slide, [titleRtf], uid));
  for (const passage of sermon.passages) {
    const refRtf = buildRtf([[span(passage.screen)]], t.ref_size, "left", true, t.font, t);
    const bodyRtf = buildRtf(passage.paragraphs, t.body_size, "left", false, t.font, t);
    addCue(passage.label, slideFrom(t.ref_slide, [refRtf, bodyRtf], uid));
  }

  p.cues = cues;
  p.cue_groups = [{
    group: { uuid: uid(), name: t.group_name },
    cue_identifiers: cueIds,
  }];
  return PP.encode(p).finish();
}
