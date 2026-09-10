// Порт export_pptx.py → PptxGenJS. Глобалы PptxGenJS и JSZip — из vendor/
// (jszip подключается раньше pptxgen.bundle.js). Тема — параметры initPro().
// Изменять синхронно с export_pptx.py (эталон) — проверяется test_core.mjs.

import { span, lighten } from "./core.js";

const EMU_PER_PX = 6350;  // слайд 12192000×6858000 EMU == 1920×1080 единиц темы
const PT_SCALE = 0.5;     // высота слайда 540pt против 1080 единиц темы
const IN_PER_EMU = 1 / 914400;
const inch = (px) => px * EMU_PER_PX * IN_PER_EMU;

const family = (name) => name.split("-")[0];  // «Onest-Regular» → «Onest»
const hex = (rgb) => rgb.map((v) => v.toString(16).padStart(2, "0").toUpperCase()).join("");

// paragraphs → ранги PptxGenJS; конец абзаца — breakLine
function toRuns(paragraphs, sizePt, baseBold, font) {
  const items = [];
  for (const para of paragraphs) {
    if (!para.length) {
      items.push({ text: "", options: { breakLine: true, fontFace: font, fontSize: sizePt * PT_SCALE } });
      continue;
    }
    para.forEach((sp, j) => {
      const o = {
        fontFace: font, fontSize: sizePt * PT_SCALE,
        bold: !!(sp.bold || baseBold), italic: !!sp.italic,
        color: sp.color ? hex(lighten(sp.color)) : "FFFFFF",  // белый, как у .pro
      };
      if (sp.highlight) o.highlight = hex(lighten(sp.highlight));
      items.push({ text: sp.text, options: { ...o, breakLine: j === para.length - 1 } });
    });
  }
  return items;
}

function addText(slide, box, sizePt, align, valignMid, paragraphs, baseBold, font) {
  slide.addText(toRuns(paragraphs, sizePt, baseBold, font), {
    x: inch(box[0]), y: inch(box[1]), w: inch(box[2]), h: inch(box[3]),
    align, valign: valignMid ? "middle" : "top", wrap: true,
    fontFace: font, fontSize: sizePt * PT_SCALE,
  });
}

export async function buildPptx(sermon, theme, outputType = "blob") {
  const t = theme;
  const titleFont = family(t.title_font);
  const font = family(t.font);
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "PP", width: inch(1920), height: inch(1080) });
  pptx.layout = "PP";
  const bg = hex(t.slide_bg.map((v) => Math.round(v * 255)));

  const addSlide = () => {
    const s = pptx.addSlide();
    s.background = { color: bg };
    return s;
  };

  let s = addSlide();
  addText(s, t.title_box, t.title_size, "center", true, [[span(sermon.title)]], true, titleFont);
  for (const passage of sermon.passages) {
    s = addSlide();
    addText(s, t.ref_box, t.ref_size, "left", true, [[span(passage.screen)]], true, font);
    addText(s, t.body_box, t.body_size, "left", false, passage.paragraphs, false, font);
  }
  return pptx.write({ outputType });
}
