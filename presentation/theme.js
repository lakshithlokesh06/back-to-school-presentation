// Shared theme values and layout helpers. Tune these in later design phases.
const pptxgen = require("pptxgenjs");

const theme = {
  colors: {
    background: "F7F4EE",
    ink: "172A3A",
    muted: "526474",
    accent: "D36A45",
    accentSoft: "E9D9C8",
    surface: "FFFFFF",
    line: "D8D4CC",
  },
  fonts: {
    heading: "Aptos Display",
    body: "Aptos",
  },
  page: {
    width: 13.333,
    height: 7.5,
    marginX: 0.72,
    marginY: 0.58,
  },
};

function createPresentation() {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Communication Skills Assignment";
  pptx.subject = "Back to School — OBC presentation scaffold";
  pptx.company = "St. Joseph Pre-University College, Bangalore";
  pptx.lang = "en-IN";
  pptx.theme = {
    headFontFace: theme.fonts.heading,
    bodyFontFace: theme.fonts.body,
    lang: "en-IN",
  };
  return pptx;
}

function addBase(slide, index, section) {
  slide.background = { color: theme.colors.background };
  slide.addShape("line", {
    x: theme.page.marginX,
    y: 6.98,
    w: theme.page.width - theme.page.marginX * 2,
    h: 0,
    line: { color: theme.colors.line, width: 0.7 },
  });
  slide.addText(section.toUpperCase(), {
    x: theme.page.marginX,
    y: 7.05,
    w: 4.5,
    h: 0.18,
    fontFace: theme.fonts.body,
    fontSize: 8,
    bold: true,
    charSpacing: 1.2,
    color: theme.colors.muted,
    margin: 0,
  });
  slide.addText(`${String(index).padStart(2, "0")} / 03`, {
    x: 11.7,
    y: 7.05,
    w: 0.9,
    h: 0.18,
    fontFace: theme.fonts.body,
    fontSize: 8,
    color: theme.colors.muted,
    align: "right",
    margin: 0,
  });
}

function addPlaceholder(slide, label, opts = {}) {
  const { x, y, w, h, accent = false } = opts;
  slide.addShape("roundRect", {
    x, y, w, h,
    rectRadius: 0.08,
    fill: { color: accent ? theme.colors.accentSoft : theme.colors.surface, transparency: 22 },
    line: { color: theme.colors.line, width: 1, dash: "dash" },
  });
  slide.addText(label, {
    x: x + 0.18,
    y: y + h / 2 - 0.16,
    w: w - 0.36,
    h: 0.32,
    fontFace: theme.fonts.body,
    fontSize: 11,
    color: theme.colors.muted,
    align: "center",
    valign: "mid",
    margin: 0.02,
  });
}

module.exports = { pptxgen, theme, createPresentation, addBase, addPlaceholder };
