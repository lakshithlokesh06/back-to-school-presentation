const path = require("node:path");
const content = require("./content");
const { theme, createPresentation, addBase, addPlaceholder } = require("./theme");

const pptx = createPresentation();

function addOpening(data) {
  const slide = pptx.addSlide();
  addBase(slide, 1, "Opening · OBC");
  slide.addShape("ellipse", {
    x: 9.15, y: 1.1, w: 2.9, h: 2.9,
    fill: { color: theme.colors.accentSoft, transparency: 14 },
    line: { color: theme.colors.accentSoft, transparency: 100 },
  });
  addPlaceholder(slide, data.visualPlaceholder, { x: 8.45, y: 4.35, w: 3.95, h: 1.25, accent: true });
  slide.addText(data.title, {
    x: theme.page.marginX, y: 1.28, w: 7.45, h: 2.35,
    fontFace: theme.fonts.heading, fontSize: 34, bold: true,
    color: theme.colors.ink, breakLine: false, margin: 0, valign: "mid",
    fit: "shrink",
  });
  slide.addText(data.subtitle, {
    x: theme.page.marginX, y: 3.98, w: 6.8, h: 0.38,
    fontFace: theme.fonts.body, fontSize: 17, color: theme.colors.accent, bold: true, margin: 0,
  });
  slide.addText(data.school, {
    x: theme.page.marginX, y: 4.49, w: 6.8, h: 0.32,
    fontFace: theme.fonts.body, fontSize: 12, color: theme.colors.muted, margin: 0,
  });
}

function addBody(data) {
  const slide = pptx.addSlide();
  addBase(slide, 2, "Body · 3 takeaways");
  slide.addText(data.title, {
    x: theme.page.marginX, y: 0.62, w: 10.5, h: 0.65,
    fontFace: theme.fonts.heading, fontSize: 27, bold: true,
    color: theme.colors.ink, margin: 0,
  });
  const top = 1.72;
  const gap = 0.23;
  const cardH = 1.38;
  const cardW = 11.9;
  data.points.forEach((point, i) => {
    const y = top + i * (cardH + gap);
    slide.addShape("roundRect", {
      x: theme.page.marginX, y, w: cardW, h: cardH,
      rectRadius: 0.08,
      fill: { color: theme.colors.surface },
      line: { color: theme.colors.line, width: 0.8 },
    });
    slide.addShape("ellipse", {
      x: 1.04, y: y + 0.27, w: 0.82, h: 0.82,
      fill: { color: i === 1 ? theme.colors.accentSoft : "E9EEF0" },
      line: { color: "FFFFFF", transparency: 100 },
    });
    slide.addText(String(i + 1).padStart(2, "0"), {
      x: 1.04, y: y + 0.52, w: 0.82, h: 0.22,
      fontFace: theme.fonts.heading, fontSize: 12, bold: true,
      color: theme.colors.ink, align: "center", margin: 0,
    });
    slide.addText(point.title, {
      x: 2.15, y: y + 0.27, w: 6.7, h: 0.38,
      fontFace: theme.fonts.heading, fontSize: 19, bold: true,
      color: theme.colors.ink, margin: 0,
    });
    slide.addText(point.description, {
      x: 2.15, y: y + 0.76, w: 8.4, h: 0.3,
      fontFace: theme.fonts.body, fontSize: 13, color: theme.colors.muted, margin: 0,
    });
    addPlaceholder(slide, "ICON", { x: 10.8, y: y + 0.38, w: 0.82, h: 0.62, accent: i === 1 });
  });
}

function addClosing(data) {
  const slide = pptx.addSlide();
  addBase(slide, 3, "Close · Your next step");
  slide.addShape("ellipse", {
    x: 9.35, y: 1.2, w: 2.35, h: 2.35,
    fill: { color: theme.colors.accentSoft, transparency: 16 },
    line: { color: theme.colors.accentSoft, transparency: 100 },
  });
  addPlaceholder(slide, data.visualPlaceholder, { x: 8.45, y: 4.35, w: 3.85, h: 1.1, accent: true });
  slide.addText(data.title, {
    x: theme.page.marginX, y: 0.82, w: 7.5, h: 0.62,
    fontFace: theme.fonts.heading, fontSize: 29, bold: true,
    color: theme.colors.ink, margin: 0,
  });
  slide.addText(data.prompt, {
    x: theme.page.marginX, y: 2.05, w: 7.9, h: 1.15,
    fontFace: theme.fonts.heading, fontSize: 27, bold: true,
    color: theme.colors.ink, margin: 0, fit: "shrink",
  });
  data.sequence.forEach((word, i) => {
    const x = theme.page.marginX + i * 1.72;
    slide.addShape("roundRect", {
      x, y: 3.65, w: 1.47, h: 0.55, rectRadius: 0.08,
      fill: { color: i === 2 ? theme.colors.accent : theme.colors.accentSoft },
      line: { color: theme.colors.accentSoft, transparency: 100 },
    });
    slide.addText(word, {
      x, y: 3.81, w: 1.47, h: 0.2,
      fontFace: theme.fonts.body, fontSize: 12, bold: true,
      color: i === 2 ? "FFFFFF" : theme.colors.ink, align: "center", margin: 0,
    });
    if (i < data.sequence.length - 1) {
      slide.addText("→", {
        x: x + 1.48, y: 3.77, w: 0.24, h: 0.24,
        fontFace: theme.fonts.body, fontSize: 14, color: theme.colors.muted, align: "center", margin: 0,
      });
    }
  });
  slide.addText(data.thankYou, {
    x: theme.page.marginX, y: 5.03, w: 3, h: 0.4,
    fontFace: theme.fonts.heading, fontSize: 17, bold: true,
    color: theme.colors.accent, margin: 0,
  });
}

addOpening(content.slides[0]);
addBody(content.slides[1]);
addClosing(content.slides[2]);

const out = path.resolve(__dirname, "..", "back-to-school-scaffold.pptx");
pptx.writeFile({ fileName: out });
