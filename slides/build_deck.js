const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/61998724-611c-41d8-94f2-03edd7a4fb28_67db12f2-f87f-4e7d-8318-bd68def116d6/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "/home/user/ai-metascience-2026/slides/class-metascience-plan.pptx";

const THEME = {
  name: "Psych 251 plain",
  headFontFace: "News Gothic MT",
  bodyFontFace: "News Gothic MT",
  colors: {
    dk1: "000000", lt1: "FFFFFF", dk2: "595959", lt2: "F2F2F2",
    accent1: "000000", accent2: "7F7F7F", accent3: "A6A6A6",
    accent4: "404040", accent5: "BFBFBF", accent6: "262626",
    hlink: "000000", folHlink: "595959",
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625 in
pres.author = "Michael C. Frank";
pres.title = "Class metascience 2026: proposal and choice points";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

const FOOTER = "Psych 251 · class metascience 2026";

pres.defineSlideMaster({
  title: "TITLE",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.2, w: 8.8, h: 1.9,
        fontSize: 40, bold: true, color: C.text1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 3.35, w: 8.8, h: 1.2,
        fontSize: 18, color: C.text2, valign: "top", margin: 0 }, text: "" } },
  ],
});

pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.3, w: 8.8, h: 0.7,
        fontSize: 30, bold: true, color: C.text1, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { text: { text: FOOTER, options: { x: 0.6, y: 5.2, w: 6, h: 0.25, fontSize: 10,
        color: C.text2, margin: 0, isTextBox: true } } },
  ],
  slideNumber: { x: 8.9, y: 5.2, w: 0.5, h: 0.25, fontSize: 10, color: C.text2, align: "right", margin: 0 },
});

// Thin black outline boxes are the deck's one visual motif.
const box = (slide, x, y, w, h, name) =>
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, objectName: name,
    fill: { color: C.background1 }, line: { color: C.text1, width: 1 } });

// ---------------------------------------------------------------- 1. Title
pres.addSection({ title: "Proposal" });
{
  const s = pres.addSlide({ masterName: "TITLE", sectionTitle: "Proposal" });
  s.addText("Do best-practices extensions improve replications?", { placeholder: "title" });
  s.addText([
    { text: "Psych 251 class metascience project, 2026", options: { breakLine: true } },
    { text: "Proposal and choice points · Monday, November 2" },
  ], { placeholder: "body" });
  s.addNotes("Most of you have opted in to the metascience project. Today: what we propose to study with your projects, and the choices we still have to make before we register the plan on November 13.");
}

// ---------------------------------------------------------------- 2. Design
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Proposal" });
  s.addText("The proposal", { placeholder: "title" });

  const steps = [
    ["Original study", "The published effect you chose"],
    ["Phase 1: replication", "As close to the original as you can make it"],
    ["Phase 2: extension", "1–3 best-practices changes: a rescue if Phase 1 failed, an upgrade if it worked"],
  ];
  const bx = [0.6, 3.7, 6.8], by = 1.25, bw = 2.6, bh = 1.5;
  steps.forEach(([head, body], i) => {
    box(s, bx[i], by, bw, bh, `Step ${i + 1} box`);
    s.addText([
      { text: head, options: { bold: true, fontSize: 16, breakLine: true } },
      { text: body, options: { fontSize: 13, color: C.text2 } },
    ], { x: bx[i] + 0.15, y: by + 0.12, w: bw - 0.3, h: bh - 0.24, valign: "top",
         margin: 0, paraSpaceAfter: 4, isTextBox: true, objectName: `Step ${i + 1} text` });
    if (i < 2) {
      s.addShape(pres.shapes.LINE, { x: bx[i] + bw + 0.08, y: by + bh / 2, w: 0.34, h: 0,
        line: { color: C.text1, width: 1.5, endArrowType: "triangle" }, objectName: `Arrow ${i + 1}` });
    }
  });

  s.addText([
    { text: "The question", options: { bold: true, fontSize: 16, breakLine: true } },
    { text: "Do changes drawn from Experimentology (measurement, design, sample, checks) make a replication more successful, and recover more of the original effect?", options: { fontSize: 14 } },
  ], { x: 0.6, y: 3.0, w: 4.5, h: 1.9, valign: "top", margin: 0, paraSpaceAfter: 6, isTextBox: true, objectName: "Question" });

  s.addText([
    { text: "Who's in", options: { bold: true, fontSize: 16, breakLine: true } },
    { text: "Opted-in projects: extra budget, and every opted-in student is an author", options: { fontSize: 14, bullet: true, breakLine: true } },
    { text: "Track 1 only; scientific extensions (Track 2) are excluded", options: { fontSize: 14, bullet: true, breakLine: true } },
    { text: "We expect about 15–20 projects", options: { fontSize: 14, bullet: true } },
  ], { x: 5.4, y: 3.0, w: 4.0, h: 1.9, valign: "top", margin: 0, paraSpaceAfter: 6, isTextBox: true, objectName: "Who" });

  s.addNotes("Every project has three studies: the original, your Phase 1 replication, and your Phase 2 extension. Projects that opt in are analysed together. Track 2 projects rerun the original plus a new condition, which is a different kind of study, so they are not included.");
}

// ---------------------------------------------------------------- 3. Outcomes
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Proposal" });
  s.addText("What we will measure", { placeholder: "title" });

  s.addText([
    { text: "1  Replication success", options: { bold: true, fontSize: 18, breakLine: true } },
    { text: "Two TAs independently rate each phase against the original, from 0 to 1. A score of .75 or more counts as a success.", options: { fontSize: 14 } },
  ], { x: 0.6, y: 1.2, w: 5.2, h: 1.25, valign: "top", margin: 0, paraSpaceAfter: 4, isTextBox: true, objectName: "Outcome 1" });

  s.addText([
    { text: "2  Effect size ratio", options: { bold: true, fontSize: 18, breakLine: true } },
    { text: "Your effect divided by the original effect, both as standardized mean differences, for Phase 1 and Phase 2. We compute it by re-running your analysis on the data in your repository.", options: { fontSize: 14 } },
  ], { x: 0.6, y: 2.55, w: 5.2, h: 1.55, valign: "top", margin: 0, paraSpaceAfter: 4, isTextBox: true, objectName: "Outcome 2" });

  s.addText("Also: precision (how much the standard error shrank, beyond what sample size alone explains).", {
    x: 0.6, y: 4.25, w: 5.2, h: 0.6, fontSize: 12, color: C.text2, valign: "top", margin: 0, isTextBox: true, objectName: "Secondary" });

  box(s, 6.3, 1.2, 3.1, 3.65, "Context box");
  s.addText([
    { text: "5 of 17", options: { bold: true, fontSize: 48, breakLine: true } },
    { text: "rescues in Boyce et al. (2024) mostly or fully replicated (29%)", options: { fontSize: 14, breakLine: true } },
    { text: "Different students, timeline, budget and tools, so our comparison with 2024 is descriptive only.", options: { fontSize: 12, color: C.text2 } },
  ], { x: 6.5, y: 1.4, w: 2.7, h: 3.25, valign: "middle", margin: 0, paraSpaceAfter: 10, isTextBox: true, objectName: "Context" });

  s.addNotes("Both outcomes follow Veronica Boyce's 2024 rescue paper so that the numbers are comparable. The ratio is 1 if you recover the original effect size exactly, 0 if you find nothing, and negative if the effect goes the other way.");
}

// ---------------------------------------------------------------- 4. Caveats
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Proposal" });
  s.addText("Two things to know about this design", { placeholder: "title" });

  s.addChart(pres.charts.BAR, [
    { name: "Raw change", labels: ["Rescues", "Upgrades"], values: [0.13, -0.22] },
    { name: "Corrected for regression to the mean", labels: ["Rescues", "Upgrades"], values: [-0.02, -0.02] },
  ], {
    x: 0.6, y: 1.15, w: 4.9, h: 3.8, objectName: "Regression to the mean chart",
    barDir: "col", barGapWidthPct: 60,
    chartColors: ["000000", "A6A6A6"],
    showTitle: true, title: "Change in effect size ratio when Phase 2 changes nothing (simulated)",
    titleFontSize: 12, titleColor: "000000", titleFontFace: "+mn-lt",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "+0.00;-0.00;0.00",
    dataLabelFontSize: 11, dataLabelColor: "000000", dataLabelFontFace: "+mn-lt",
    valAxisMinVal: -0.3, valAxisMaxVal: 0.2, valAxisMajorUnit: 0.1, valAxisLabelFormatCode: "+0.0;-0.0;0.0",
    valAxisLabelFontSize: 10, valAxisLabelColor: "595959", valAxisLabelFontFace: "+mn-lt",
    catAxisLabelFontSize: 12, catAxisLabelColor: "000000", catAxisLabelFontFace: "+mn-lt",
    catAxisLabelPos: "low",
    valGridLine: { color: "D9D9D9", size: 0.5 }, catGridLine: { style: "none" },
    showLegend: true, legendPos: "b", legendFontSize: 11, legendColor: "000000", legendFontFace: "+mn-lt",
  });

  s.addText([
    { text: "Regression to the mean", options: { bold: true, fontSize: 18, breakLine: true } },
    { text: "Rescues are chosen for low Phase 1 results, so they tend to rise anyway, and upgrades tend to fall. The main test pools all projects, where this cancels; within type we correct for it.", options: { fontSize: 14 } },
  ], { x: 5.85, y: 1.2, w: 3.55, h: 1.95, valign: "top", margin: 0, paraSpaceAfter: 4, isTextBox: true, objectName: "Caveat 1" });

  s.addText([
    { text: "Modest power", options: { bold: true, fontSize: 18, breakLine: true } },
    { text: "With about 20 projects, a 30% larger effect in Phase 2 (say, 65% of the original instead of 50%) is detected only 39% of the time. We report an estimate with an interval, not just yes or no.", options: { fontSize: 14 } },
  ], { x: 5.85, y: 3.25, w: 3.55, h: 1.75, valign: "top", margin: 0, paraSpaceAfter: 4, isTextBox: true, objectName: "Caveat 2" });

  s.addNotes("Numbers are from 1,000 simulated classes of 20 projects in the planning document. The black bars show the bias you would see if you just compared Phase 2 with Phase 1 within rescues or within upgrades; the grey bars show that the correction removes it.");
}

// ---------------------------------------------------------------- 5. Choice points
pres.addSection({ title: "Choice points" });
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Choice points" });
  s.addText("Choice points", { placeholder: "title" });

  const cards = [
    ["When do TAs rate Phase 1?", "Proposed: at the midterms this week, before any Phase 2 data exist."],
    ["Rescue or upgrade: who decides?", "Proposed: the TA rating of Phase 1 (.75 or more = upgrade); your declared track is recorded too."],
    ["What if Phase 2 changes the design?", "Between → within changes the metric. Proposed: recompute on the Phase 1 scale and flag it."],
    ["How ambiguous was the original?", "Proposed: count differences and unstated choices in your report, plus an agent count a person checks."],
    ["Agents checking the original papers", "Proposed: an agent flags statistical inconsistencies; a TA verifies each flag after Phase 2 data are in."],
    ["Studying us, not the experiments", "Attitudes, agent logs, quizzes: need a new IRB protocol and committee sign-off. Opt-in, never graded."],
  ];
  const cw = 2.8, ch = 1.55, gx = 0.2, gy = 0.2, x0 = 0.6, y0 = 1.15;
  cards.forEach(([q, p], i) => {
    const x = x0 + (i % 3) * (cw + gx), y = y0 + Math.floor(i / 3) * (ch + gy);
    box(s, x, y, cw, ch, `Choice ${i + 1} box`);
    s.addText([
      { text: q, options: { bold: true, fontSize: 14, breakLine: true } },
      { text: p, options: { fontSize: 12, color: C.text2 } },
    ], { x: x + 0.12, y: y + 0.1, w: cw - 0.24, h: ch - 0.2, valign: "top", margin: 0,
         paraSpaceAfter: 4, isTextBox: true, objectName: `Choice ${i + 1} text` });
  });

  s.addText("Plan registered on OSF by Nov 13 · Phase 2 data collected the week of Nov 16", {
    x: 0.6, y: 4.62, w: 8.8, h: 0.35, fontSize: 14, bold: true, valign: "middle", margin: 0, isTextBox: true, objectName: "Timeline" });

  s.addNotes("These are the open decisions from the planning document (D2–D5, D8–D10). The first one is urgent: if the TAs are to rate Phase 1 before Phase 2 exists, they rate at the midterms on Wednesday and Friday.");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
