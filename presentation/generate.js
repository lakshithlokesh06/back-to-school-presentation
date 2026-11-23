const path = require("node:path");
const content = require("./content");
const { theme, createPresentation, addBase, addPlaceholder } = require("./theme");

const pptx = createPresentation();

function addOpening(data) {
  const slide = pptx.addSlide();
  addBase(slide, 1, "Opening · OBC");

  // One composed illustration: a quiet classroom in the background, a desk
  // and open notebook in the foreground, and a path that moves forward.
  slide.addShape("roundRect", {
    x: 8.22, y: 0.86, w: 4.38, h: 5.52,
    rectRadius: 0.12,
    fill: { color: theme.colors.surface },
    line: { color: theme.colors.line, width: 0.8 },
  });
  slide.addShape("rect", {
    x: 8.62, y: 1.22, w: 3.58, h: 2.35,
    fill: { color: "E9EEF0" },
    line: { color: "E9EEF0", transparency: 100 },
  });
  slide.addShape("rect", {
    x: 8.91, y: 1.53, w: 2.96, h: 1.66,
    fill: { color: "F7F4EE" },
    line: { color: theme.colors.line, width: 0.8 },
  });
  // Simple board marks suggest a real lesson without adding more slide copy.
  slide.addShape("line", {
    x: 9.25, y: 2.03, w: 1.15, h: 0,
    line: { color: theme.colors.accent, width: 2.4, beginArrowType: "none", endArrowType: "none" },
  });
  slide.addShape("line", {
    x: 9.25, y: 2.42, w: 2.05, h: 0,
    line: { color: "B9C4CA", width: 1.6 },
  });
  slide.addShape("line", {
    x: 9.25, y: 2.72, w: 1.46, h: 0,
    line: { color: "B9C4CA", width: 1.6 },
  });
  slide.addShape("rect", {
    x: 8.62, y: 3.57, w: 3.58, h: 0.14,
    fill: { color: theme.colors.accentSoft },
    line: { color: theme.colors.accentSoft, transparency: 100 },
  });
  // The dotted path visually links the former classroom to the next step.
  slide.addShape("line", {
    x: 9.08, y: 4.31, w: 2.72, h: 0,
    line: { color: theme.colors.accent, width: 2.2, dash: "dash", endArrowType: "triangle" },
  });
  slide.addShape("ellipse", {
    x: 9.06, y: 4.2, w: 0.22, h: 0.22,
    fill: { color: theme.colors.accent },
    line: { color: theme.colors.accent, transparency: 100 },
  });
  // Desk and open notebook form the foreground focal point.
  slide.addShape("line", {
    x: 8.85, y: 5.62, w: 3.16, h: 0,
    line: { color: theme.colors.ink, width: 3 },
  });
  slide.addShape("line", {
    x: 9.14, y: 5.62, w: -0.22, h: 0.53,
    line: { color: theme.colors.ink, width: 2.1 },
  });
  slide.addShape("line", {
    x: 11.72, y: 5.62, w: 0.22, h: 0.53,
    line: { color: theme.colors.ink, width: 2.1 },
  });
  slide.addShape("chevron", {
    x: 9.75, y: 4.87, w: 0.76, h: 0.49,
    rotate: 180,
    fill: { color: theme.colors.accentSoft },
    line: { color: theme.colors.accentSoft, transparency: 100 },
  });
  slide.addShape("chevron", {
    x: 10.47, y: 4.87, w: 0.76, h: 0.49,
    fill: { color: theme.colors.accentSoft },
    line: { color: theme.colors.accentSoft, transparency: 100 },
  });
  slide.addShape("line", {
    x: 10.48, y: 4.87, w: 0, h: 0.48,
    line: { color: theme.colors.accent, width: 1.2 },
  });

  slide.addText(data.title, {
    x: theme.page.marginX, y: 1.38, w: 6.88, h: 2.55,
    fontFace: theme.fonts.heading, fontSize: 35, bold: true,
    color: theme.colors.ink, breakLine: false, margin: 0, valign: "mid",
    fit: "shrink",
  });
  slide.addText(data.subtitle, {
    x: theme.page.marginX, y: 4.34, w: 6.7, h: 0.42,
    fontFace: theme.fonts.body, fontSize: 17, color: theme.colors.accent, bold: true, margin: 0,
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
