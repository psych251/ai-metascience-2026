const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/61998724-611c-41d8-94f2-03edd7a4fb28_67db12f2-f87f-4e7d-8318-bd68def116d6/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "/home/user/ai-metascience-2026/slides/ai-policy-and-metascience.pptx";

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
pres.title = "AI policy and the class metascience project";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

pres.defineSlideMaster({
  title: "TITLE",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.0, w: 8.8, h: 1.9,
        fontSize: 40, bold: true, color: C.text1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 3.15, w: 8.8, h: 1.0,
        fontSize: 18, color: C.text2, valign: "top", margin: 0 }, text: "" } },
  ],
});

// One content layout per section, differing only in the footer, so each section stands alone.
const contentLayout = (title, footer) => pres.defineSlideMaster({
  title,
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.3, w: 8.8, h: 0.7,
        fontSize: 30, bold: true, color: C.text1, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { text: { text: footer, options: { x: 0.6, y: 5.2, w: 6, h: 0.25, fontSize: 10,
        color: C.text2, margin: 0, isTextBox: true } } },
  ],
  slideNumber: { x: 8.9, y: 5.2, w: 0.5, h: 0.25, fontSize: 10, color: C.text2, align: "right", margin: 0 },
});
contentLayout("CONTENT_POLICY", "Psych 251 · AI policy");
contentLayout("CONTENT_META", "Psych 251 · class metascience 2026");

// Thin black outline boxes are the deck's one visual motif.
const box = (slide, x, y, w, h, name) =>
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, objectName: name,
    fill: { color: C.background1 }, line: { color: C.text1, width: 1 } });

// ---------------------------------------------------------------- Title
pres.addSection({ title: "AI policy" });
{
  const s = pres.addSlide({ masterName: "TITLE", sectionTitle: "AI policy" });
  s.addText("AI policy and the class metascience project", { placeholder: "title" });
  s.addText([
    { text: "Psych 251 · Monday, November 2", options: { breakLine: true } },
    { text: "1  Suggested AI policies and norms    2  The class metascience project" },
  ], { placeholder: "body" });
  s.addText("AI use: these slides were drafted with Claude Code (Anthropic) from our metascience planning document and the post “My lab’s AI policies and norms” (October 2026), at Mike’s direction.", {
    x: 0.6, y: 4.55, w: 8.8, h: 0.55, fontSize: 11, color: C.text2, valign: "top", margin: 0,
    isTextBox: true, objectName: "AI disclosure" });
  s.addNotes("Two separate topics today. First, some suggestions for how to use AI in your work, adapted from the policies and norms I wrote up for my lab. Second, the class metascience project that most of you have opted in to.");
}

// Rows of a bold lead-in and an explanation: the layout for both policy slides.
const ruleRows = (s, rows, prefix) => rows.forEach(([lead, text], i) => {
  const y = 1.2 + i * 0.76;
  s.addText(lead, { x: 0.6, y, w: 2.75, h: 0.62, fontSize: 15, bold: true, valign: "top",
    margin: 0, isTextBox: true, objectName: `${prefix} ${i + 1} lead` });
  s.addText(text, { x: 3.55, y, w: 5.85, h: 0.62, fontSize: 14, valign: "top",
    margin: 0, isTextBox: true, objectName: `${prefix} ${i + 1} text` });
});

// ---------------------------------------------------------------- Policies
{
  const s = pres.addSlide({ masterName: "CONTENT_POLICY", sectionTitle: "AI policy" });
  s.addText("Suggested AI policies", { placeholder: "title" });
  ruleRows(s, [
    ["You write the text", "Your proposal and report are written by you. Agents can write code, documentation and admin."],
    ["Always disclose", "Every submission has an AI use statement, even if you used none: what for, and which model and version."],
    ["Agent review last", "After your own read-through, have an agent review your work. Consider every suggestion; keep what you can defend."],
    ["Protect participants", "No Prolific IDs or other identifiers in a commercial model. Free text, audio and video only under a no-training agreement."],
    ["Respect others’ work", "Don’t put a manuscript you are reviewing, or a classmate’s work, into a model unless that is allowed."],
  ], "Policy");
  s.addNotes("These are adapted from my lab's policies. The first two are already course policy. Agent review is new: do it after your own review, not instead of it. On data: your anonymized CSVs are fine; anything with identifiers is not.");
}

// ---------------------------------------------------------------- Norms
{
  const s = pres.addSlide({ masterName: "CONTENT_POLICY", sectionTitle: "AI policy" });
  s.addText("Suggested AI norms", { placeholder: "title" });
  ruleRows(s, [
    ["Engineer it", "Version control, tests for reusable code, an organized project, clear documentation (Poldrack, Better Code, Better Science)."],
    ["Mark AI text you share", "Quote it when you paste it into Slack, so nobody mistakes it for your own words."],
    ["Don’t send slop", "Condense before you share, by hand or with AI. If something is unpolished, say so."],
    ["Read the paper first", "Struggle with it yourself before asking a model about it; you will remember more."],
    ["Avoid the slop cliff", "Keep understanding your own project: check details, ask follow-ups, and treat a surprising answer as a reason to dig in."],
  ], "Norm");
  s.addNotes("Norms are softer than policies: habits that keep you in control of your work. The slop cliff is the point where a project has moved faster than your understanding of it, and you can no longer explain or defend what it does.");
}

// ---------------------------------------------------------------- Metascience: design
pres.addSection({ title: "Metascience" });
{
  const s = pres.addSlide({ masterName: "CONTENT_META", sectionTitle: "Metascience" });
  s.addText("Class metascience: the proposal", { placeholder: "title" });

  const steps = [
    ["Original study", "The published effect you chose"],
    ["Phase 1: replication", "As close to the original as you can make it"],
    ["Phase 2: extension", "1–3 best-practices changes, whether Phase 1 worked or not"],
  ];
  const bx = [0.6, 3.7, 6.8], by = 1.25, bw = 2.6, bh = 1.3;
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
  ], { x: 0.6, y: 2.9, w: 4.5, h: 2.0, valign: "top", margin: 0, paraSpaceAfter: 6, isTextBox: true, objectName: "Question" });

  s.addText([
    { text: "Who's in", options: { bold: true, fontSize: 16, breakLine: true } },
    { text: "Opted-in projects: extra budget, and every opted-in student is an author", options: { fontSize: 14, bullet: true, breakLine: true } },
    { text: "Track 1 only; scientific extensions (Track 2) are excluded", options: { fontSize: 14, bullet: true, breakLine: true } },
    { text: "We expect about 15–20 projects", options: { fontSize: 14, bullet: true } },
  ], { x: 5.4, y: 2.9, w: 4.0, h: 2.0, valign: "top", margin: 0, paraSpaceAfter: 6, isTextBox: true, objectName: "Who" });

  s.addNotes("Every project has three studies: the original, your Phase 1 replication, and your Phase 2 extension. Whether we call Phase 2 a rescue or an upgrade, it draws on the same menu of changes, so we analyse all extensions together. Track 2 projects rerun the original plus a new condition, which is a different kind of study, so they are not included.");
}

// ---------------------------------------------------------------- 3. Outcomes
{
  const s = pres.addSlide({ masterName: "CONTENT_META", sectionTitle: "Metascience" });
  s.addText("What we will measure", { placeholder: "title" });

  s.addText([
    { text: "Replication success", options: { bold: true, fontSize: 18, breakLine: true } },
    { text: "Course staff rate each phase against the original, from 0 to 1. A score of .75 or more counts as a success.", options: { fontSize: 14 } },
  ], { x: 0.6, y: 1.2, w: 4.2, h: 1.6, valign: "top", margin: 0, paraSpaceAfter: 4, isTextBox: true, objectName: "Outcome 1" });

  s.addText([
    { text: "Effect size ratio", options: { bold: true, fontSize: 18, breakLine: true } },
    { text: "Your effect divided by the original effect, both as standardized mean differences. We compute it by re-running your analysis on the data in your repository.", options: { fontSize: 14 } },
  ], { x: 5.2, y: 1.2, w: 4.2, h: 1.6, valign: "top", margin: 0, paraSpaceAfter: 4, isTextBox: true, objectName: "Outcome 2" });

  s.addText([
    { text: "How we compare", options: { bold: true, fontSize: 18, breakLine: true } },
    { text: "Phase 2 against Phase 1, pooled across every project. Boyce et al. (2024) found that 5 of 17 rescues of failed class replications (29%) mostly or fully replicated; we will set our results beside theirs, but the studies differ too much for a formal comparison.", options: { fontSize: 14 } },
  ], { x: 0.6, y: 3.05, w: 8.8, h: 1.5, valign: "top", margin: 0, paraSpaceAfter: 4, isTextBox: true, objectName: "Comparison" });

  s.addText("Also: precision (how much the standard error shrank, beyond what sample size alone explains).", {
    x: 0.6, y: 4.6, w: 8.8, h: 0.35, fontSize: 12, color: C.text2, valign: "top", margin: 0, isTextBox: true, objectName: "Secondary" });

  s.addNotes("Both outcomes follow Veronica Boyce's 2024 rescue paper so that the numbers are comparable. The ratio is 1 if you recover the original effect size exactly, 0 if you find nothing, and negative if the effect goes the other way. We do not compare rescues with upgrades: a project is a rescue because its Phase 1 came out low, so it would tend to look better in Phase 2 even if nothing changed.");
}

// ---------------------------------------------------------------- 4. Choice points
{
  const s = pres.addSlide({ masterName: "CONTENT_META", sectionTitle: "Metascience" });
  s.addText("Choice points", { placeholder: "title" });

  const cards = [
    ["Staff ratings of replication success", "All four of us rate every project: during presentations, or from your report if the data were not in yet."],
    ["Rescue or upgrade: who decides?", "Is there really a difference?"],
    ["Ambiguity ratings", "How systematic should they be, and what system do we use?"],
    ["Agent reproducibility audits", "When do they happen, and who runs and checks them?"],
    ["Attitudes, chat logs, comprehension quiz", "These study you rather than your experiments, so they require a new IRB protocol."],
  ];
  const cw = 2.8, ch = 1.55, gx = 0.2, gy = 0.2, x0 = 0.6, y0 = 1.15;
  cards.forEach(([q, p], i) => {
    const row = Math.floor(i / 3), inRow = row === 0 ? 3 : cards.length - 3;
    const x = x0 + ((3 - inRow) * (cw + gx)) / 2 + (i % 3) * (cw + gx), y = y0 + row * (ch + gy);
    box(s, x, y, cw, ch, `Choice ${i + 1} box`);
    s.addText([
      { text: q, options: { bold: true, fontSize: 14, breakLine: true } },
      { text: p, options: { fontSize: 12, color: C.text2 } },
    ], { x: x + 0.12, y: y + 0.1, w: cw - 0.24, h: ch - 0.2, valign: "top", margin: 0,
         paraSpaceAfter: 4, isTextBox: true, objectName: `Choice ${i + 1} text` });
  });

  s.addText("Plan registered on OSF by Nov 13 · Phase 2 data collected the week of Nov 16", {
    x: 0.6, y: 4.62, w: 8.8, h: 0.35, fontSize: 14, bold: true, valign: "middle", margin: 0, isTextBox: true, objectName: "Timeline" });

  s.addNotes("These are the decisions still open before we register. Ratings: every project gets a rating for each phase from all four of us, made as soon as that phase's data exist. The last card is separate from the main study: anything that studies you directly needs its own IRB protocol and consent, and would never be graded.");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
