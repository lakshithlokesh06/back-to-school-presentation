const path = require("node:path");
const content = require("./content");
const { theme, createPresentation, addBase, addPlaceholder } = require("./theme");

const pptx = createPresentation();

function addOpening(data) {
  const slide = pptx.addSlide();
  addBase(slide, 1, "Opening · OBC");
  slide.addImage({
    path: path.resolve(__dirname, "..", "assets", "slide1-student-classroom.png"),
    x: 7.58,
    y: 0.90,
    w: 5.05,
    h: 5.72,
    sizing: { type: "cover", x: 7.58, y: 0.90, w: 5.05, h: 5.72 },
    altText: "Student writing at a classroom desk and looking out toward the city",
    shadow: { type: "outer", color: "172A3A", opacity: 0.12, blur: 5, offset: 2, angle: 45 },
  });

  slide.addText(data.title, {
    x: theme.page.marginX, y: 1.38, w: 6.45, h: 2.55,
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
    x: theme.page.marginX, y: 0.55, w: 10.5, h: 0.72,
    fontFace: theme.fonts.heading, fontSize: 30, bold: true,
    color: theme.colors.ink, margin: 0,
  });

  const cardY = 1.58;
  const cardW = 3.88;
  const cardH = 4.92;
  const cardGap = 0.28;
  const cardX = theme.page.marginX;

  function addIcon(index, x, y) {
    const stroke = { color: theme.colors.ink, width: 1.7, beginArrowType: "none", endArrowType: "none" };
    const accentLine = { color: theme.colors.accent, width: 2.2, beginArrowType: "none", endArrowType: "none" };
    if (index === 0) {
      // A marked paper and pencil: results are part of school, not identity.
      slide.addShape("roundRect", { x: x + 0.18, y: y + 0.13, w: 0.92, h: 1.05, rectRadius: 0.07, fill: { color: theme.colors.surface }, line: stroke });
      slide.addShape("line", { x: x + 0.38, y: y + 0.41, w: 0.46, h: 0, line: accentLine });
      slide.addShape("line", { x: x + 0.38, y: y + 0.65, w: 0.52, h: 0, line: { color: theme.colors.line, width: 1.4 } });
      slide.addShape("line", { x: x + 0.38, y: y + 0.83, w: 0.40, h: 0, line: { color: theme.colors.line, width: 1.4 } });
      slide.addShape("line", { x: x + 0.85, y: y + 0.94, w: 0.42, h: -0.76, line: { color: theme.colors.accent, width: 2.5, beginArrowType: "none", endArrowType: "none" } });
      slide.addShape("line", { x: x + 0.85, y: y + 0.94, w: -0.09, h: 0.16, line: { color: theme.colors.accent, width: 2.1 } });
    } else if (index === 1) {
      // A diverging path with a compass point signals choice and direction.
      slide.addShape("line", { x: x + 0.64, y: y + 1.12, w: 0, h: -0.43, line: stroke });
      slide.addShape("line", { x: x + 0.64, y: y + 0.73, w: -0.36, h: -0.46, line: stroke });
      slide.addShape("line", { x: x + 0.64, y: y + 0.73, w: 0.36, h: -0.46, line: stroke });
      slide.addShape("line", { x: x + 0.64, y: y + 1.10, w: -0.37, h: 0, line: accentLine });
      slide.addShape("line", { x: x + 0.64, y: y + 1.10, w: 0.37, h: 0, line: accentLine });
      slide.addShape("ellipse", { x: x + 0.52, y: y + 0.60, w: 0.24, h: 0.24, fill: { color: theme.colors.accent }, line: { color: theme.colors.accent, transparency: 100 } });
      slide.addShape("chevron", { x: x + 0.20, y: y + 0.19, w: 0.20, h: 0.20, rotate: 270, fill: { color: theme.colors.accent }, line: { color: theme.colors.accent, transparency: 100 } });
      slide.addShape("chevron", { x: x + 0.89, y: y + 0.19, w: 0.20, h: 0.20, rotate: 90, fill: { color: theme.colors.accent }, line: { color: theme.colors.accent, transparency: 100 } });
    } else {
      // Three growing blocks and a rising line represent skills built over time.
      slide.addShape("rect", { x: x + 0.20, y: y + 0.77, w: 0.29, h: 0.35, fill: { color: theme.colors.surface }, line: stroke });
      slide.addShape("rect", { x: x + 0.49, y: y + 0.57, w: 0.29, h: 0.55, fill: { color: theme.colors.accentSoft }, line: stroke });
      slide.addShape("rect", { x: x + 0.78, y: y + 0.35, w: 0.29, h: 0.77, fill: { color: theme.colors.surface }, line: stroke });
      slide.addShape("line", { x: x + 0.18, y: y + 1.16, w: 0.96, h: 0, line: stroke });
      slide.addShape("line", { x: x + 0.22, y: y + 0.60, w: 0.83, h: -0.47, line: accentLine });
      slide.addShape("chevron", { x: x + 0.94, y: y + 0.04, w: 0.20, h: 0.20, rotate: 315, fill: { color: theme.colors.accent }, line: { color: theme.colors.accent, transparency: 100 } });
    }
  }

  data.points.forEach((point, i) => {
    const x = cardX + i * (cardW + cardGap);
    slide.addShape("roundRect", {
      x, y: cardY, w: cardW, h: cardH,
      rectRadius: 0.12,
      fill: { color: theme.colors.surface },
      line: { color: theme.colors.line, width: 0.9 },
    });
    slide.addShape("roundRect", {
      x: x + 0.20, y: cardY + 0.22, w: cardW - 0.40, h: 1.48,
      rectRadius: 0.1,
      fill: { color: i === 1 ? "F3E8DE" : "E9EEF0" },
      line: { color: i === 1 ? "F3E8DE" : "E9EEF0", transparency: 100 },
    });
    slide.addText(`0${i + 1}`, {
      x: x + 0.37, y: cardY + 0.35, w: 0.46, h: 0.24,
      fontFace: theme.fonts.body, fontSize: 10, bold: true,
      charSpacing: 1, color: theme.colors.muted, margin: 0,
    });
    addIcon(i, x + 1.32, cardY + 0.32);
    slide.addText(point.title, {
      x: x + 0.30, y: cardY + 2.05, w: cardW - 0.60, h: 0.92,
      fontFace: theme.fonts.heading, fontSize: 22, bold: true,
      color: theme.colors.ink, margin: 0, valign: "mid", fit: "shrink",
    });
    slide.addText(point.description, {
      x: x + 0.30, y: cardY + 3.13, w: cardW - 0.60, h: 1.05,
      fontFace: theme.fonts.body, fontSize: 15, color: theme.colors.muted,
      margin: 0, valign: "top", breakLine: false, fit: "shrink",
    });
    slide.addShape("line", {
      x: x + 0.30, y: cardY + 4.38, w: cardW - 0.60, h: 0,
      line: { color: theme.colors.line, width: 0.8 },
    });
    slide.addShape("ellipse", {
      x: x + 0.30, y: cardY + 4.57, w: 0.14, h: 0.14,
      fill: { color: theme.colors.accent },
      line: { color: theme.colors.accent, transparency: 100 },
    });
    slide.addShape("line", {
      x: x + 0.52, y: cardY + 4.64, w: 0.62, h: 0,
      line: { color: theme.colors.accentSoft, width: 2.4 },
    });
  });
}

function addClosing(data) {
  const slide = pptx.addSlide();
  addBase(slide, 3, "Close · Your next step");

  // A single path-and-student illustration carries the closing idea: one
  // deliberate first step, then a clear route forward.
  slide.addShape("roundRect", {
    x: 8.43, y: 0.83, w: 4.16, h: 5.63,
    rectRadius: 0.12,
    fill: { color: theme.colors.surface },
    line: { color: theme.colors.line, width: 0.8 },
  });
  slide.addShape("ellipse", {
    x: 9.44, y: 1.24, w: 2.12, h: 2.12,
    fill: { color: "E9EEF0" },
    line: { color: "E9EEF0", transparency: 100 },
  });
  // Narrowing path leads toward the horizon; dashed center line suggests
  // progress without introducing a second visual motif.
  slide.addShape("trapezoid", {
    x: 9.16, y: 3.35, w: 2.74, h: 2.68,
    fill: { color: "E9EEF0" },
    line: { color: "E9EEF0", transparency: 100 },
  });
  slide.addShape("line", {
    x: 10.53, y: 5.73, w: 0.01, h: -2.12,
    line: { color: theme.colors.accent, width: 1.6, dash: "dash" },
  });
  // Single student silhouette standing at the beginning of the path.
  slide.addShape("ellipse", {
    x: 10.34, y: 4.32, w: 0.38, h: 0.38,
    fill: { color: theme.colors.ink },
    line: { color: theme.colors.ink, transparency: 100 },
  });
  slide.addShape("roundRect", {
    x: 10.28, y: 4.72, w: 0.50, h: 0.88,
    rectRadius: 0.12,
    fill: { color: theme.colors.accent },
    line: { color: theme.colors.accent, transparency: 100 },
  });
  slide.addShape("line", {
    x: 10.39, y: 5.52, w: -0.16, h: 0.47,
    line: { color: theme.colors.ink, width: 2.4 },
  });
  slide.addShape("line", {
    x: 10.67, y: 5.52, w: 0.16, h: 0.47,
    line: { color: theme.colors.ink, width: 2.4 },
  });
  slide.addShape("line", {
    x: 10.29, y: 4.86, w: -0.28, h: 0.35,
    line: { color: theme.colors.ink, width: 2.1 },
  });
  slide.addShape("line", {
    x: 10.77, y: 4.86, w: 0.27, h: -0.25,
    line: { color: theme.colors.ink, width: 2.1 },
  });

  slide.addText(data.title, {
    x: theme.page.marginX, y: 0.82, w: 7.3, h: 0.62,
    fontFace: theme.fonts.heading, fontSize: 30, bold: true,
    color: theme.colors.ink, margin: 0,
  });
  slide.addText(data.prompt, {
    x: theme.page.marginX, y: 2.02, w: 7.25, h: 1.72,
    fontFace: theme.fonts.heading, fontSize: 31, bold: true,
    color: theme.colors.ink, margin: 0, valign: "mid", fit: "shrink",
  });
  slide.addText(data.sequence.join("  →  "), {
    x: theme.page.marginX, y: 4.18, w: 6.9, h: 0.44,
    fontFace: theme.fonts.body, fontSize: 17, bold: true,
    color: theme.colors.accent, margin: 0,
  });
  slide.addText(data.thankYou, {
    x: theme.page.marginX, y: 5.51, w: 3, h: 0.38,
    fontFace: theme.fonts.heading, fontSize: 15, bold: true,
    color: theme.colors.muted, margin: 0,
  });
}

addOpening(content.slides[0]);
addBody(content.slides[1]);
addClosing(content.slides[2]);

const out = path.resolve(__dirname, "..", "back-to-school-scaffold.pptx");
pptx.writeFile({ fileName: out });
