import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const workspaceDir = "C:\\Users\\serba\\Documents\\ChatGPT\\PhD";
const TMP_DIR = path.join(workspaceDir, "tym_deck_build");
const SKILL_DIR = "C:\\Users\\serba\\.codex\\plugins\\cache\\openai-primary-runtime\\presentations\\26.909.12148\\skills\\presentations";
const FINAL_PPTX = path.join(workspaceDir, "output", "TYM_MLNET_Synthetic_Research_Roadmap_2027_RoTimeBank_Annotation_Pilot.pptx");
const coverPath = path.join(TMP_DIR, "cover-aurora.jpg");
const { resolvePresentationFont } = await import(pathToFileURL(path.join(SKILL_DIR, "container_tools/artifact_tool_utils.mjs")).href);
const FONT = resolvePresentationFont({ fontFamily: "Arial" });
const C = {
  bg: "#202522",
  bg2: "#272D29",
  bg3: "#303833",
  white: "#F3F6F4",
  mint: "#62DDA8",
  mint2: "#9AEBCA",
  muted: "#AAB8B0",
  line: "#47534B",
  amber: "#F0C36A",
  red: "#E58E81",
  black: "#111714",
};
const W = 1280;
const H = 720;
const presentation = Presentation.create({ slideSize: { width: W, height: H } });

function addText(slide, text, x, y, w, h, size = 22, color = C.white, bold = false, align = "left", name = "Text") {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name,
    position: { left: x, top: y, width: w, height: h },
    fill: "none",
    line: { fill: "none", width: 0 },
  });
  shape.text = text;
  shape.text.style = {
    typeface: FONT,
    fontSize: size,
    bold,
    color,
    alignment: align,
    verticalAlignment: "top",
    autoFit: "shrinkText",
    wrap: "square",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  };
  return shape;
}
function addRect(slide, x, y, w, h, fill, line = "none", lineWidth = 0, geometry = "rect", name = "Diagram shape") {
  return slide.shapes.add({
    geometry,
    name,
    position: { left: x, top: y, width: w, height: h },
    fill,
    line: { fill: line, width: lineWidth },
  });
}
function addRule(slide, x, y, w, color = C.line, h = 1) {
  return addRect(slide, x, y, w, h, color, "none", 0, "rect", "Section rule");
}
function baseSlide(title, section, index) {
  const slide = presentation.slides.add();
  slide.background.fill = C.bg;
  addText(slide, "TYM  /  " + section.toUpperCase(), 64, 27, 900, 22, 15, C.mint, true, "left", "Section label");
  addText(slide, title, 64, 61, 1150, 56, 41, C.white, true, "left", "Slide title");
  addRule(slide, 64, 137, 1152, C.line, 1);
  addRule(slide, 64, 674, 1152, C.line, 1);
  addText(slide, "SYNTHETIC DATA  /  UNSUPERVISED DISCOVERY  /  HUMAN EVALUATION", 64, 689, 950, 18, 12, C.muted, false, "left", "Footer");
  addText(slide, String(index).padStart(2, "0"), 1160, 687, 56, 22, 14, C.mint2, true, "right", "Page number");
  return slide;
}
function setNotes(slide, notes) {
  slide.speakerNotes.textFrame.setText(notes);
}
function addDiagramNode(slide, x, y, w, h, number, label, detail) {
  addRect(slide, x, y, w, h, C.bg2, C.line, 1, "roundRect", "Editable workflow node");
  addText(slide, number, x + 16, y + 15, 35, 22, 15, C.mint, true, "left", "Step number");
  addText(slide, label, x + 16, y + 42, w - 30, 30, 21, C.white, true, "left", "Step label");
  addText(slide, detail, x + 16, y + 77, w - 30, h - 86, 18, C.muted, false, "left", "Step detail");
}
function addConnector(slide, x1, y1, x2, y2, color = C.mint) {
  const left = Math.min(x1, x2);
  const top = Math.min(y1, y2);
  const width = Math.max(2, Math.abs(x2 - x1));
  const height = Math.max(2, Math.abs(y2 - y1));
  if (Math.abs(y2 - y1) < 2) return addRect(slide, left, y1 - 1, width, 2, color, "none", 0, "rect", "Diagram connector");
  return addRect(slide, x1 - 1, top, 2, height, color, "none", 0, "rect", "Diagram connector");
}

// 1. Cover
{
  const slide = presentation.slides.add();
  slide.background.fill = C.black;
  const photo = new Uint8Array(await fs.readFile(coverPath));
  slide.images.add({ blob: photo, contentType: "image/jpeg", alt: "Aurora above a mountain lake at night", fit: "cover", position: { left: 0, top: 0, width: W, height: H } });
  addRect(slide, 0, 0, W, H, "#101612/62", "none", 0, "rect", "Photo contrast overlay");
  addText(slide, "TYM  /  RESEARCH ROADMAP", 72, 62, 760, 22, 16, C.mint2, true, "left", "Cover eyebrow");
  addText(slide, "Synthetic data for\nnarrative time", 72, 252, 820, 150, 58, C.white, true, "left", "Cover title");
  addText(slide, "ML.NET training, unlabeled discovery, and human-gold evaluation", 74, 431, 790, 74, 26, C.mint2, false, "left", "Cover subtitle");
  addRule(slide, 74, 548, 190, C.mint, 3);
  addText(slide, "A data-centric empirical program for the 2027 NLP conference cycle", 74, 574, 900, 36, 20, C.white, false, "left", "Cover scope");
  addText(slide, "Conference deadlines checked 1 October 2026", 74, 651, 720, 22, 15, C.muted, false, "left", "Cover date");
  setNotes(slide, "Research and engineering roadmap for the TYM project. The cover photograph was provided by the user. Conference dates and deadlines reflect official pages checked on 1 October 2026.");
}

// 2. Research thesis
{
  const slide = baseSlide("Research thesis", "Research question", 2);
  addText(slide, "Can governed synthetic examples improve temporal-semantic coverage when human labels are scarce?", 64, 166, 1120, 78, 34, C.white, true, "left", "Research question");
  addText(slide, "Unlabeled narratives guide what to review. Human adjudication establishes labels. Synthetic descendants expand training coverage.", 64, 257, 1090, 70, 22, C.muted, false, "left", "Research framing");
  const xs = [64, 361, 658, 955];
  const labels = [
    ["01", "Unlabeled text", "Narrative pool"],
    ["02", "Discovery strata", "Clusters and outliers"],
    ["03", "Human review", "Adjudicated examples"],
    ["04", "Synthetic training", "Controlled descendants"],
  ];
  for (let i = 0; i < xs.length; i++) addDiagramNode(slide, xs[i], 372, 245, 126, labels[i][0], labels[i][1], labels[i][2]);
  for (let i = 0; i < 3; i++) addConnector(slide, xs[i] + 249, 435, xs[i + 1] - 5, 435, C.mint);
  addText(slide, "Testable claim", 64, 546, 230, 26, 18, C.mint, true, "left", "Claim label");
  addText(slide, "Synthetic gains should survive held-out chapters and unseen generator patterns, with the same human-gold test set.", 295, 540, 840, 52, 22, C.white, false, "left", "Testable claim");
  setNotes(slide, "This is a proposed research claim, not a result. Synthetic examples will be evaluated only as additions to training data. Unlabeled clusters will guide sampling and will never be treated as TYM labels. See the implementation and evaluation guide in the repository.");
}

// 3. TYM target
{
  const slide = baseSlide("What the model must recover", "Task definition", 3);
  addText(slide, "Narrative order and story time can diverge. TYM represents both the text sequence and the temporal structure inferred from it.", 64, 165, 1100, 58, 22, C.muted, false);
  const x = [64, 448, 832];
  const widths = [340, 340, 384];
  const heads = ["EVENTS IN TEXT", "TIME SEGMENTS", "TEMPORAL LINKS"];
  const examples = [
    "EV1  “grew up”\nEV2  “remembered”\nEV3  “found”",
    "TS labels:\nNAR  /  REM  /  SUP\nGEN  /  FIC",
    "EV2 BEFORE EV3\nTS2 BEFORE TS3\nTIMEX3 anchors where explicit",
  ];
  for (let i = 0; i < 3; i++) {
    addText(slide, heads[i], x[i], 267, widths[i], 24, 16, C.mint, true);
    addText(slide, examples[i], x[i], 309, widths[i], 132, 26, C.white, false);
    if (i < 2) addRect(slide, x[i] + widths[i] + 20, 263, 1, 195, C.line);
  }
  addRule(slide, 64, 487, 1152, C.line, 1);
  addText(slide, "Output layers", 64, 520, 190, 28, 18, C.mint, true);
  addText(slide, "The API emits TYM tracks, segments, endpoints, and relations alongside TimeML-style EVENT, TIMEX3, SIGNAL, MAKEINSTANCE, and TLINK annotations.", 270, 514, 875, 70, 22, C.white, false);
  setNotes(slide, "The supplied paper describes time segments with NAR, REM, SUP, GEN, and FIC labels, plus TYM time-yard relations and a TimeML-style annotation layer. Source: user-provided PDF, especially the model and annotation description. Repository representation: Tym.Api response models and README.");
}

// 4. Paper baseline
{
  const slide = baseSlide("Paper baseline and measurement caveat", "Evidence", 4);
  const cols = [64, 453, 842];
  const nums = ["58.9%", "58.7%", "93"];
  const labs = ["accuracy", "macro F1", "time segments produced"];
  for (let i = 0; i < 3; i++) {
    addText(slide, nums[i], cols[i], 191, 330, 88, 68, C.mint, true);
    addText(slide, labs[i], cols[i], 284, 330, 38, 24, C.white, false);
  }
  addText(slide, "Reference evaluation: 2,000 event instances", 64, 390, 720, 34, 22, C.white, true);
  addRule(slide, 64, 450, 1152, C.line, 1);
  addText(slide, "Interpretation", 64, 482, 180, 25, 18, C.amber, true);
  addText(slide, "The paper reports these results. The repository has no measured performance yet. The PDF reports corpus-level totals separately; their relation to the 2,000-event evaluation remains unclear in the reported sections.", 270, 474, 888, 102, 21, C.white, false);
  addText(slide, "Keep the paper score separate from the API baseline.", 64, 604, 930, 30, 19, C.muted, false);
  setNotes(slide, "The user-provided PDF reports 58.9% accuracy and approximately 58.7% macro-F1 for 2,000 event instances, and reports 93 produced time segments (PDF pp. 10–11). It also gives corpus-level annotation totals on p. 10. The evaluation confusion matrix and corpus counts are presented in separate contexts and their denominator relationship is unclear, so this deck keeps them distinct. These numbers have not been reproduced by this branch.");
}

// 5. Current repository
{
  const slide = baseSlide("What the repository can currently prove", "Implementation baseline", 5);
  addText(slide, "67", 64, 183, 190, 80, 70, C.mint, true);
  addText(slide, "synthetic demo rows checked in", 66, 270, 330, 30, 20, C.white, false);
  addText(slide, "43 segment-type rows   /   24 event-temporal rows", 66, 311, 490, 28, 18, C.muted, false);
  addRule(slide, 594, 183, 1, C.line, 360);
  addText(slide, "Supervised path", 642, 181, 270, 26, 18, C.mint, true);
  addText(slide, "ML.NET FeaturizeText + SDCA maximum-entropy.\nGrouped chapter/document evaluation accepts only adjudicated human gold. Synthetic rows can enter training only.", 642, 222, 520, 112, 21, C.white, false);
  addText(slide, "Unsupervised path", 642, 376, 290, 25, 18, C.mint, true);
  addText(slide, "ML.NET text n-grams + K-Means. The command reports lexical clusters, distances, and exemplars. Measure TYM performance on human-held-out gold.", 642, 414, 520, 98, 21, C.white, false);
  addRule(slide, 64, 555, 1152, C.line, 1);
  addText(slide, "Evidence boundary", 64, 579, 190, 28, 18, C.amber, true);
  addText(slide, "No adjudicated TYM corpus is checked into this repository. Current seed-trained artifacts are pipeline demos, not research baselines.", 270, 573, 880, 54, 21, C.white, false);
  setNotes(slide, "Repository inspection on 1 October 2026: data/seed-examples.jsonl contains 43 segment_type records and 24 event_temporal records, marked synthetic_seed/unreviewed. The modeling project implements FeaturizeText plus SdcaMaximumEntropy classification, grouped evaluation with human_gold/adjudicated-only metrics, and K-Means clustering over FeaturizeText features. No gold corpus is checked into the repo. Sources: README.md, tools/Tym.Modeling/Program.cs, data/seed-examples.jsonl.");
}

// 6. Synthetic curriculum
{
  const slide = baseSlide("Synthetic examples should target contrastive coverage", "Synthetic-data design", 6);
  addText(slide, "Generate examples around a labeled distinction, then change the surface form independently.", 64, 161, 1100, 34, 22, C.muted);
  const rows = [
    ["Narrative frame", "Contrast narration with remembered, anticipated, imagined, or habitual event framing."],
    ["Tense and anchor", "Vary tense, aspect, temporal adverbials, and reference time without making one cue identify the label."],
    ["Event expression", "Include verbal events and nominalized events such as arrival, departure, or discovery."],
    ["Relation distance", "Place linked events in separate sentences or paragraphs; vary relation direction and cue placement."],
    ["Boundary ambiguity", "Include missing, conflicting, or weak cues and mark them for uncertainty or human review."],
  ];
  rows.forEach((r, i) => {
    const y = 220 + i * 82;
    addText(slide, r[0].toUpperCase(), 64, y, 285, 26, 16, C.mint, true);
    addText(slide, r[1], 360, y - 3, 800, 58, 20, C.white, false);
    addRule(slide, 64, y + 64, 1152, C.line, 1);
  });
  setNotes(slide, "This is a proposed synthetic curriculum for TYM's narrative temporal and segment-label tasks. It builds on documented temporal annotation gaps: nominalized events and long-distance links in TIMELINE, and dense event-pair annotation in NarrativeTime. Sources: https://aclanthology.org/2023.emnlp-main.1016/ and https://aclanthology.org/2024.lrec-main.1054/. All generated labels remain candidates until checked against the annotation guideline.");
}

// 7. Data quality controls
{
  const slide = baseSlide("Generation controls reduce label shortcuts", "Synthetic-data design", 7);
  const left = 64;
  const right = 673;
  const items = [
    ["Minimal contrasts", "Keep the event content stable while changing one temporal or narrative cue."],
    ["Surface variation", "Change syntax, names, verbs, cue location, and context without changing the intended phenomenon."],
    ["Lineage", "Record parent, phenomenon, transformation, generator version, and review state for every row."],
    ["Human audit", "Check labels and span boundaries on a sampled batch before adding a generator family to training."],
  ];
  items.forEach((item, i) => {
    const col = i % 2 === 0 ? left : right;
    const row = Math.floor(i / 2);
    const y = 198 + row * 193;
    addText(slide, "0" + (i + 1), col, y, 46, 28, 17, C.mint, true);
    addText(slide, item[0], col + 60, y - 2, 480, 31, 25, C.white, true);
    addText(slide, item[1], col + 60, y + 43, 480, 88, 20, C.muted, false);
    if (row === 0) addRule(slide, col, y + 147, 500, C.line, 1);
  });
  addText(slide, "A generator label is provenance, not a quality guarantee.", 64, 607, 1040, 30, 20, C.amber, true);
  setNotes(slide, "Recommended controls are hypotheses and protocol design, not implemented data generation. Recent synthetic-data studies report distribution-shift and diversity risks. For example, Kambhatla et al. find that fine-grained persona detail adds little lexical-diversity gain beyond a length cutoff in their setting; Geng et al. show pseudo-label alignment and distribution shift concerns for synthetic QE. Sources: https://aclanthology.org/2025.findings-emnlp.1146/ and https://aclanthology.org/2025.acl-long.373/. These results motivate measurement and human auditing, not direct transfer of their findings to TYM.");
}

// 8. Unsupervised discovery
{
  const slide = baseSlide("Unlabeled discovery prioritizes annotation", "Unsupervised learning", 8);
  addText(slide, "Current method", 64, 178, 200, 28, 18, C.mint, true);
  addText(slide, "Narrative text", 64, 234, 250, 42, 25, C.white, true);
  addRule(slide, 64, 288, 275, C.mint, 2);
  addText(slide, "FeaturizeText\nword and character n-grams", 64, 312, 330, 68, 21, C.white, false);
  addRule(slide, 64, 399, 275, C.mint, 2);
  addText(slide, "K-Means\nexploratory groupings", 64, 425, 330, 65, 21, C.white, false);
  addRule(slide, 64, 511, 275, C.mint, 2);
  addText(slide, "Inspect exemplars,\nrare groups, and outliers", 64, 538, 330, 54, 20, C.white, false);
  addRect(slide, 465, 183, 1, 430, C.line);
  addText(slide, "Use clusters to", 525, 183, 290, 26, 18, C.mint, true);
  addText(slide, "Find lexical modes missing from the gold set.\nSample annotation batches by cluster and outlier status.\nChoose phenomena for the next synthetic contrast set.", 525, 225, 620, 145, 22, C.white, false);
  addText(slide, "Human interpretation", 525, 410, 260, 27, 18, C.amber, true);
  addText(slide, "Use cluster IDs as lexical strata for review. Annotators decide whether a group captures a TYM phenomenon. Evaluate model quality against held-out human labels.", 525, 450, 620, 126, 22, C.white, false);
  setNotes(slide, "ML.NET describes clustering as unsupervised learning. The repository's CLI uses FeaturizeText features with K-Means and emits assignments, distances, Davies-Bouldin index, and exemplars. N-gram clusters are lexical groupings; inspect them and have annotators validate any interpretation. Cluster IDs are not TYM labels. Sources: https://learn.microsoft.com/en-us/dotnet/machine-learning/resources/tasks and https://learn.microsoft.com/en-us/dotnet/machine-learning/resources/metrics.");
}

// 9. Annotation loop
{
  const slide = baseSlide("Unlabeled text becomes an annotation sampling frame", "Unsupervised learning", 9);
  const x = [64, 358, 652, 946];
  const data = [
    ["01", "Narrative pool", "Collect unlabeled text under clear rights and privacy rules."],
    ["02", "Cluster review", "Inspect representative and outlier examples; name candidate strata."],
    ["03", "Human annotation", "Apply written TYM guidelines and adjudicate disagreement."],
    ["04", "Training candidates", "Generate targeted descendants from train-split parents only."],
  ];
  for (let i = 0; i < x.length; i++) addDiagramNode(slide, x[i], 230, 244, 160, data[i][0], data[i][1], data[i][2]);
  for (let i = 0; i < x.length - 1; i++) addConnector(slide, x[i] + 248, 310, x[i + 1] - 4, 310, C.mint);
  addRule(slide, 64, 456, 1152, C.line, 1);
  addText(slide, "Keep two ledgers", 64, 483, 220, 26, 18, C.mint, true);
  addText(slide, "Annotation ledger: source, annotator, guideline version, adjudication, and document split.\nGeneration ledger: parent, phenomenon, transformation, generator version, and review status.", 64, 527, 1092, 84, 22, C.white, false);
  setNotes(slide, "Proposed human-in-the-loop data workflow. The repository supports JSONL provenance fields and parent checks but does not yet implement a dedicated annotation platform or synthetic generator. Cluster assignments should be used for sampling only. Keep licensed or private narrative text outside public commits.");
}

// 10. Split and leakage control
{
  const slide = baseSlide("Split human gold before generating descendants", "Evaluation design", 10);
  addText(slide, "Partition documents or chapters first.", 64, 171, 840, 38, 27, C.white, true);
  addText(slide, "All synthetic descendants inherit the parent group and stay in train.", 64, 218, 920, 32, 21, C.muted, false);
  const groups = [
    { x: 64, title: "TRAIN", color: C.mint, body: "Adjudicated human train\n+ synthetic descendants\n+ allowed seed demos" },
    { x: 456, title: "DEV", color: C.amber, body: "Human gold only\nModel and ratio selection\nNo synthetic descendants" },
    { x: 848, title: "TEST", color: C.red, body: "Human gold only\nLocked until final scoring\nNo generator or chapter leakage" },
  ];
  for (const g of groups) {
    addRect(slide, g.x, 313, 345, 206, C.bg2, g.color, 2, "rect", "Evaluation partition");
    addText(slide, g.title, g.x + 20, 333, 290, 29, 19, g.color, true);
    addText(slide, g.body, g.x + 20, 384, 300, 110, 22, C.white, false);
  }
  addText(slide, "Hold out generator families as a separate stress test.", 64, 574, 1060, 34, 21, C.mint2, true);
  setNotes(slide, "The ML.NET evaluator creates or reads chapter/document-level group manifests. Evaluation rows are restricted to source_type=human_gold and review_status=adjudicated. Optional --train-seeds rows enter only the training partition, and provided parent IDs are checked against split membership. For stronger generator generalization claims, keep a generator family out of training and report it separately. See tools/Tym.Modeling/Program.cs and README.md.");
}

// 11. Controlled experiment
{
  const slide = baseSlide("Ablations isolate the value of synthetic data", "Evaluation design", 11);
  addText(slide, "Keep the held-out gold documents and evaluation code identical across runs.", 64, 164, 1100, 33, 22, C.muted);
  const widths = [260, 260, 260, 260];
  const x = [64, 350, 636, 922];
  const ratios = ["0 : 1", "0.25 : 1", "0.5 : 1", "1 : 1"];
  const names = ["Human only", "Low synthetic", "Medium synthetic", "High synthetic"];
  for (let i = 0; i < x.length; i++) {
    addText(slide, ratios[i], x[i], 234, widths[i], 62, 43, C.mint, true);
    addText(slide, "synthetic : human", x[i], 302, widths[i], 28, 17, C.muted, false);
    addText(slide, names[i], x[i], 355, widths[i], 36, 22, C.white, true);
    addRule(slide, x[i], 412, widths[i] - 18, i === 0 ? C.line : C.mint, 2);
  }
  addText(slide, "Repeat by phenomenon and generator family.", 64, 472, 475, 27, 18, C.mint, true);
  addText(slide, "Report label balance, deduplication, and synthetic-to-gold ratio for every run.", 64, 516, 535, 68, 20, C.white, false);
  addText(slide, "Decision question", 662, 472, 250, 27, 18, C.amber, true);
  addText(slide, "Does augmentation improve macro-F1 and targeted error slices without harming held-out human examples?", 662, 516, 515, 70, 20, C.white, false);
  setNotes(slide, "The ratios are a proposed starting grid, not established optimal values. Define the denominator at the example level within each task and report class and phenomenon distributions. Run human-only and augmented conditions on the same train/dev/test document manifest. Keep tuning away from the final held-out test.");
}

// 12. Metrics
{
  const slide = baseSlide("Report label, span, relation, and graph evidence separately", "Evaluation design", 12);
  const sections = [
    ["AVAILABLE IN CLI", "Accuracy, macro-F1, per-class precision/recall/F1, support, confusion matrix, log loss.\nCluster average distance and Davies-Bouldin index."],
    ["NEXT EVALUATORS", "Event-span boundary precision/recall/F1.\nTemporal-relation macro-F1 and direction-confusion analysis.\nWhole-graph consistency checks."],
    ["DATA QUALITY", "Annotator agreement before adjudication.\nCalibration by phenomenon and confidence band.\nError counts by chapter, genre, and generator family."],
  ];
  const xs = [64, 455, 846];
  for (let i = 0; i < sections.length; i++) {
    addText(slide, sections[i][0], xs[i], 203, 340, 26, 16, i === 1 ? C.amber : C.mint, true);
    addText(slide, sections[i][1], xs[i], 249, 340, 188, 21, C.white, false);
    if (i < 2) addRect(slide, xs[i] + 361, 202, 1, 280, C.line);
  }
  addRule(slide, 64, 514, 1152, C.line, 1);
  addText(slide, "Current boundary", 64, 545, 220, 27, 18, C.amber, true);
  addText(slide, "The CLI covers grouped classification and clustering. Event-span and end-to-end graph evaluators are next implementation steps.", 290, 537, 860, 62, 21, C.white, false);
  setNotes(slide, "ML.NET's metrics reference documents multiclass classification and clustering metrics. The current CLI reports grouped classification metrics and internal K-Means measures; it does not implement event-span or graph-level evaluation. Annotation agreement and richer TYM metrics are future additions. Source: https://learn.microsoft.com/en-us/dotnet/machine-learning/resources/metrics and repository pocs/model-evaluation/README.md.");
}

// 13. ML.NET architecture
{
  const slide = baseSlide("ML.NET separates research runs from API inference", "Engineering", 13);
  const boxes = [
    [64, "JSONL SOURCES", "Human gold\nSynthetic descendants\nUnlabeled narratives"],
    [354, "OFFLINE ML.NET", "Train / grouped eval\nK-Means discovery\nProvenance sidecars"],
    [644, "VERSIONED OUTPUTS", "Model .zip\nSHA-256 manifest\nSeed and lineage counts"],
    [934, "TYM API + REACT", "Saved artifact loading\nPer-field provenance\nInteractive review"],
  ];
  for (const b of boxes) {
    addRect(slide, b[0], 241, 244, 209, C.bg2, C.line, 1, "rect", "ML.NET workflow stage");
    addText(slide, b[1], b[0] + 16, 261, 213, 24, 15, C.mint, true);
    addText(slide, b[2], b[0] + 16, 310, 211, 118, 22, C.white, false);
  }
  for (let i = 0; i < boxes.length - 1; i++) addConnector(slide, boxes[i][0] + 247, 345, boxes[i + 1][0] - 5, 345, C.mint);
  addRule(slide, 64, 503, 1152, C.line, 1);
  addText(slide, "Deployment setting", 64, 534, 200, 28, 18, C.mint, true);
  addText(slide, "TYM_MODEL_DIR loads both English artifacts. Without saved artifacts, the API uses its built-in seed baselines. Romanian extraction remains rule-based.", 280, 527, 875, 73, 21, C.white, false);
  setNotes(slide, "Current implementation path: tools/Tym.Modeling trains ML.NET FeaturizeText + SdcaMaximumEntropy models and writes model manifests. The API accepts saved English artifacts through TYM_MODEL_DIR; absent that configuration, it retains seed classifiers. The UI uses React 18 UMD scripts and calls the API. Sources: repository README.md, Program.cs, tools/Tym.Modeling/Program.cs, ui/Tym.Ui/wwwroot/index.html.");
}

// 14. React workbench
{
  const slide = baseSlide("React workbench supports interactive inspection", "Engineering", 14);
  addText(slide, "ANALYZE", 64, 203, 220, 25, 16, C.mint, true);
  addText(slide, "Select English or Romanian text.\nGenerate SVG, JSON, XML, or TimeML output.", 64, 244, 285, 96, 21, C.white, false);
  addRule(slide, 64, 378, 316, C.line, 1);
  addText(slide, "FILTER", 64, 402, 220, 24, 16, C.mint, true);
  addText(slide, "Narrow events by temporal category, actor, and confidence.", 64, 443, 285, 71, 21, C.white, false);
  addRule(slide, 427, 198, 1, 407, C.line);
  addText(slide, "INSPECT", 487, 203, 220, 25, 16, C.mint, true);
  addText(slide, "Select an event or segment to highlight its source span. Open the detail view for anchors, actors, relation cues, classifier confidence, and provenance.", 487, 244, 628, 105, 23, C.white, false);
  addRule(slide, 487, 378, 630, C.line, 1);
  addText(slide, "REVIEW", 487, 402, 220, 24, 16, C.mint, true);
  addText(slide, "Use the evidence panel to find low-confidence or under-covered phenomena and draft the next annotation batch.", 487, 443, 628, 72, 22, C.white, false);
  addText(slide, "The workbench exposes model evidence for review and leaves annotation and adjudication to researchers.", 64, 576, 1080, 31, 19, C.amber, true);
  setNotes(slide, "UI features observed in the local browser on 1 October 2026: React workbench, language examples, API URL, diagram width, analysis tab, category/actor/min-confidence filters, event and segment selection, source text jump, evidence details, downloadable SVG/JSON/XML, and dark-mode toggle. This is product functionality, not a user-study result.");
}

// 15. Corpus plan
{
  const slide = baseSlide("Corpus plan: one target dataset, separate comparators", "Corpus design", 15);
  addText(slide, "Use existing corpora to test compatible layers; build TYM labels on rights-cleared target narratives.", 64, 158, 1130, 42, 22, C.muted, false);
  const cards = [
    [64, "TYM HUMAN-GOLD CORE", "Original narratives in declared genres and languages.\n\nAnnotate event spans, segment/frame labels, temporal anchors, and relations with written guidelines, independent annotators, and adjudication.\n\nOnly this set defines TYM gold."],
    [456, "PUBLIC TASK COMPARATORS", "NarrativeTime / TimeBank-Dense: dense links, news domain.\n\nTIMELINE: long-distance links and nominalized events, news domain.\n\nCaTeRS / ROCStories: short stories with temporal and causal links.\n\nMEANTIME: four-language parallel news; target-language labels are partly projected."],
    [848, "UNLABELED DISCOVERY POOL", "Collect rights-cleared text that matches target genres and languages; retain source, author/work, date, genre, and language metadata.\n\nUse ML.NET n-gram clusters to inspect strata, rare groups, and outliers.\n\nCompare random and cluster-guided samples at the same human annotation budget. Clusters never become TYM labels."],
  ];
  for (const [x, head, body] of cards) {
    addRect(slide, x, 218, 368, 344, C.bg2, C.line, 1, "rect", "Corpus role card");
    addRule(slide, x, 218, 368, C.mint, 3);
    addText(slide, head, x + 18, 240, 330, 24, 15, C.mint, true);
    addText(slide, body, x + 18, 283, 330, 257, 17, C.white, false);
  }
  addText(slide, "Do not map news or short-story labels onto NAR / REM / SUP / GEN / FIC without a documented schema crosswalk and human validation.", 64, 586, 1148, 48, 18, C.amber, true);
  setNotes(slide, "Corpus roles are intentionally separate. NarrativeTime provides dense timeline annotations over TimeBank-Dense news and reports two independent expert annotations; its guidelines, tools, and annotations are available under MIT according to the paper. TIMELINE targets exhaustive news relations, including long-distance and nominalized events. CaTeRS annotates 320 five-sentence ROCStories with causal and temporal relations, useful as a short-story diagnostic but not as TYM labels. MEANTIME contains 120 English news articles and translations in Spanish, Italian, and Dutch; its non-English annotations were partly projected from English and should not be treated as independent target-language gold without audit. The user's two custom TYM XMLs are one parallel story, not two independent evaluation documents. Ro-TimeBank is a different ISO-TimeML event/time corpus. Confirm source text, annotation, and model derivative rights before using or releasing any corpus. Sources: https://aclanthology.org/2024.lrec-main.1054/ ; https://aclanthology.org/2023.emnlp-main.1016/ ; https://aclanthology.org/W16-1007/ ; https://aclanthology.org/L16-1699/ ; https://aclanthology.org/L12-1451/.");
}

// 16. Supplied annotations and source distinction
{
  const slide = baseSlide("Your annotations: a bilingual pilot, not a benchmark", "Annotation evidence", 16);
  const cards = [
    [64, "ONE PAIRED STORY", "375 RO segments\n376 EN segments\n375 shared span intervals\n\nThe English and Romanian files share nearly every narrative unit. Keep the pair in one work-level fold."],
    [456, "TYM LABELS", "382 temporal links per file\n\nFour segment types are present: NAR, REM, FIC, GEN. SUP is absent. One Romanian relation points to a missing segment and is excluded."],
    [848, "SEPARATE TIMEML DATA", "Ro-TimeBank: 183 Romanian texts, 181 EN–RO alignments.\n\nIt adds event, TIMEX3, signal, and TLINK/SLINK/ALINK labels for news, not TYM narrative-frame labels."],
  ];
  for (const [x, head, body] of cards) {
    addRect(slide, x, 174, 368, 344, C.bg2, C.line, 1, "rect", "Supplied annotation evidence card");
    addRule(slide, x, 174, 368, C.mint, 3);
    addText(slide, head, x + 18, 197, 330, 24, 15, C.mint, true);
    addText(slide, body, x + 18, 241, 330, 252, 19, C.white, false);
  }
  addText(slide, "Trainable now: private ML.NET SDCA pilot classifiers. Next research target: multilingual XLM-R plus document-level temporal graph reasoning.", 64, 548, 1148, 44, 19, C.white, true);
  addText(slide, "One story cannot establish a state-of-the-art result. The thesis's 99.18% figure is corrected annotation-transfer coverage, not model accuracy.", 64, 604, 1148, 37, 16, C.amber, true);
  setNotes(slide, "Counts were computed from the user's two supplied TYM XML files with a secure importer. 375 Romanian TS and 376 English TS segments have 375 shared span intervals. Both files contain 382 TREL tags; the importer skips one Romanian TREL with an unresolved endpoint and maps a source spelling variant SIMMULTANEOUS to SIMULTANEOUS. The parser generated 1,515 private training rows across segment_type and temporal_relation for both languages. No test metrics were generated because both editions are one story. The Ro-TimeBank archive contains 183 Romanian texts and 181 EN-RO aligned pairs, but it uses ISO-TimeML event, time, signal, and link annotations. The user-provided 2011 thesis and Forăscu & Tufiș (2012) report a 99.18% annotation transfer rate after corrections; do not quote this as classifier accuracy. NAS-BERT training was attempted and failed checkpoint loading (expected 60 state_dict entries, found 142), matching the open ML.NET issue https://github.com/dotnet/machinelearning/issues/7350; the unsupported option was removed. Sources: https://aclanthology.org/L12-1451/ ; user-supplied Forăscu (2011) thesis PDF.");
}

// 17. Conference-fit data plan
{
  const slide = baseSlide("Adapt the corpus question to each conference", "Conference-fit data plan", 17);
  const headers = ["VENUE / LENS", "CORPUS AND TRAINING ADAPTATION", "EVIDENCE THAT MAKES THE FIT REAL"];
  const hx = [64, 292, 786];
  const hw = [210, 475, 430];
  for (let i = 0; i < headers.length; i++) addText(slide, headers[i], hx[i], 161, hw[i], 21, 14, C.mint, true);
  addRule(slide, 64, 192, 1152, C.mint, 1);
  const rows = [
    ["EACL\nHuman in Language", "Double-annotated target narratives; retain disagreement and adjudication. Generate linguistically motivated minimal pairs, then ask annotators to verify label and span changes.", "Inter-annotator agreement, guideline rationale, disagreement analysis, and model errors by discourse-time phenomenon."],
    ["COLING\nNLP for Linguistics", "Add a human-validated second-language slice with comparable genres. Use MEANTIME only as a cross-lingual news transfer diagnostic; audit projected labels. Do not treat machine translation as native-language gold.", "Language-specific temporal cues, label portability, and genuine variation across original-language texts; document language and source coverage."],
    ["NAACL\nGeneralization / resources", "Pair a rights-cleared target-domain gold corpus with a same-domain unlabeled pool. Compare random versus cluster-guided human annotation at equal budgets; hold out works, authors, and genres.", "Transfer by source and genre, fixed-budget annotation gains, per-phenomenon scores, and human-gold evaluation. The agentic theme needs a real communication task."],
    ["ACL\nHomogenization theme", "Compare human-authored texts with outputs from multiple generator families on matched prompts/scenarios. Label semantic phenomena and rare-event coverage; train with controlled synthetic mixtures.", "Measure semantic and lexical diversity, phenomenon coverage/entropy, and generator-held-out performance. Claim theme fit only if homogenization is directly tested."],
    ["EMNLP\nEmpirical study", "Cross source/domain × generation strategy × generator × synthetic-to-human ratio. Preserve document lineage and use human-only, template, and diverse-generation controls.", "Paired runs, variance, held-out source/genre/generator, slice-level errors, and negative results. Re-check the CFP once the 2027 call is published."],
  ];
  rows.forEach((r, i) => {
    const y = 203 + i * 86;
    addText(slide, r[0], hx[0], y + 3, hw[0], 61, 15, i === 3 ? C.amber : C.mint2, true);
    addText(slide, r[1], hx[1], y, hw[1], 66, 15, C.white, false);
    addText(slide, r[2], hx[2], y, hw[2], 66, 15, C.muted, false);
    addRule(slide, 64, y + 72, 1152, C.line, 1);
  });
  addText(slide, "These are different corpus questions built on one shared core; a venue change must change the data or analysis, not just the title.", 64, 638, 1140, 24, 16, C.amber, true);
  setNotes(slide, "Conference-specific adaptation is based on the official 2027 calls checked on 1 October 2026. EACL's theme is The Human in Language; a new EACL main ARR submission is no longer possible because its 3 August deadline passed, so this fit applies only to already reviewed ARR work or an eligible separate route. COLING's theme is NLP for Linguistics and explicitly highlights linguistic diversity and under-researched languages. NAACL covers low-resource NLP, generalization, human-centric NLP, resources, evaluation, and semantics; its agentic-communication theme says multiple agents alone are insufficient. ACL's theme is Homogenization and knowledge collapse in LLMs, requiring direct diversity or homogenization analysis for a thematic claim. EMNLP 2027 CFP/theme is not yet published; the empirical-methods framing is provisional. Sources: https://2027.eacl.org/calls/papers/ ; https://2027.coling-iccl.org/calls/main_conference_papers/ ; https://2027.naacl.org/calls/main_conference_papers/ ; https://2027.aclweb.org/calls/main/.");
}

// 18. Workback
{
  const slide = baseSlide("Build the corpus first; use deadlines as gates", "Execution plan", 18);
  const items = [
    ["OCT 1–12", "Feasibility gate", "COLING / NAACL share ARR. Submit only if publishable gold, rights, and pilot evidence already exist; do not rush a new corpus."],
    ["OCT–NOV", "Gold corpus + strata", "Choose target genres/languages, lock work-level splits, sample unlabeled clusters, annotate and adjudicate priority strata."],
    ["NOV", "Synthetic pilot", "Generate train-only contrast sets; audit labels/spans, capture lineage, and deduplicate against source and test texts."],
    ["DEC", "Controlled study", "Compare human-only and synthetic mixtures; run equal-size controls and held-out generator / genre analyses."],
    ["JAN 4", "ACL ARR gate", "Commit only if corpus claims match the evidence. EACL main ARR is already past; EMNLP timing/theme remains TBA."],
  ];
  items.forEach((it, i) => {
    const y = 178 + i * 92;
    addText(slide, it[0], 64, y, 185, 24, 16, C.mint, true);
    addText(slide, it[1], 263, y - 2, 214, 31, 21, C.white, true);
    addText(slide, it[2], 500, y - 2, 650, 54, 19, C.muted, false);
    addRule(slide, 64, y + 61, 1152, C.line, 1);
  });
  setNotes(slide, "Workback is a proposed plan from the current date, 1 October 2026. Dates are submission gates, not synthetic-data inputs. COLING and NAACL share the 12 October ARR cycle; only one primary venue can receive the paper after review. ACL's ARR deadline is 4 January 2027. EACL main ARR deadline was 3 August 2026 and has passed, unless a reviewed paper is already available to commit. EMNLP dates and 2027 CFP are not yet available. Verify official calls again before submission. Sources: https://2027.coling-iccl.org/calls/main_conference_papers/ ; https://2027.naacl.org/calls/main_conference_papers/ ; https://2027.aclweb.org/calls/main/ ; https://2027.eacl.org/calls/papers/.");
}

// 19. Gates and risks
{
  const slide = baseSlide("A publishable claim needs three gates", "Decision criteria", 19);
  const items = [
    ["DATA RIGHTS", "The source corpus and derived examples can be used for research and described reproducibly."],
    ["LABEL VALIDITY", "Guidelines, annotator disagreement, adjudication, and uncertainty are documented."],
    ["GENERALIZATION", "Synthetic gains survive held-out documents and generator patterns on a locked gold test."],
  ];
  items.forEach((it, i) => {
    const x = 64 + i * 390;
    addText(slide, "0" + (i + 1), x, 202, 52, 28, 17, C.mint, true);
    addText(slide, it[0], x + 60, 200, 270, 28, 17, C.white, true);
    addText(slide, it[1], x + 60, 250, 278, 118, 20, C.muted, false);
    if (i < 2) addRect(slide, x + 360, 202, 1, 235, C.line);
  });
  addRule(slide, 64, 459, 1152, C.line, 1);
  addText(slide, "Pre-register the interpretation", 64, 492, 320, 27, 18, C.amber, true);
  addText(slide, "Set a practically meaningful effect threshold before the final run. Report neutral or negative synthetic-data results with the same care as positive results.", 64, 540, 1090, 62, 21, C.white, false);
  setNotes(slide, "These are proposed decision gates. Do not commit private corpus text to the public repository unless its rights permit it. Synthetic examples can hide template leakage and pseudo-label error; generator holdout and independent human gold address these risks. Clustering introduces its own risk: lexical groups can look coherent without representing the target semantic classes.");
}

// 20. Corpus references
{
  const slide = baseSlide("Corpus references: task coverage and limits", "References", 20);
  const refs = [
    ["NarrativeTime / TimeBankNT", "Full dense timeline annotation; two expert annotators; underlying documents are TimeBank-Dense news.", "aclanthology.org/2024.lrec-main.1054/"],
    ["TIMELINE", "News corpus designed for exhaustive long-distance links and nominalized event mentions.", "aclanthology.org/2023.emnlp-main.1016/"],
    ["CaTeRS / ROCStories", "320 five-sentence stories with causal and temporal event relations; useful diagnostic, not TYM frame gold.", "aclanthology.org/W16-1007/"],
    ["MEANTIME", "480 translated news articles in English, Spanish, Italian, and Dutch; some annotations were projected from English.", "aclanthology.org/L16-1699/"],
    ["TYM source paper", "User-provided PDF. Keep paper-reported metrics separate from this repo's seed models.", "Supplied local PDF, 14 pages"],
  ];
  refs.forEach((r, i) => {
    const y = 169 + i * 96;
    addText(slide, r[0], 64, y, 310, 27, 19, C.mint, true);
    addText(slide, r[1], 393, y - 1, 758, 48, 17, C.white, false);
    addText(slide, r[2], 393, y + 51, 758, 20, 13, C.muted, false);
    addRule(slide, 64, y + 76, 1152, C.line, 1);
  });
  setNotes(slide, "Corpus sources: NarrativeTime / TimeBankNT, https://aclanthology.org/2024.lrec-main.1054/ ; TIMELINE, https://aclanthology.org/2023.emnlp-main.1016/ ; CaTeRS, https://aclanthology.org/W16-1007/ ; MEANTIME, https://aclanthology.org/L16-1699/ ; user-provided TYM paper PDF. Corpus overlap is task-specific and source text may have separate redistribution conditions. MEANTIME annotations are not all independently human-authored in each language.");
}

// 21. References and tools
{
  const slide = baseSlide("Research references: synthetic data and tools", "References", 21);
  const refs = [
    ["Synthetic diversity", "Kambhatla et al. 2025. Persona detail and lexical diversity in synthetic text.", "aclanthology.org/2025.findings-emnlp.1146/"],
    ["Synthetic distribution shift", "Geng et al. 2025. Label alignment and distribution shift in synthetic quality-estimation data.", "aclanthology.org/2025.acl-long.373/"],
    ["ML.NET metrics", "Official metrics reference for classification and clustering.", "learn.microsoft.com/dotnet/machine-learning/resources/metrics"],
    ["ML.NET clustering", "Official task reference: ML.NET clustering uses unsupervised centroid-based K-Means.", "learn.microsoft.com/dotnet/machine-learning/resources/tasks"],
    ["Conference research themes", "EACL: human in language; COLING: NLP for linguistics; ACL: homogenization. NAACL main track covers resources, transfer, and evaluation.", "Official 2027 CFPs checked 1 Oct 2026; EMNLP theme TBA"],
  ];
  refs.forEach((r, i) => {
    const y = 166 + i * 91;
    addText(slide, r[0], 64, y, 295, 26, 18, C.mint, true);
    addText(slide, r[1], 380, y - 1, 775, 37, 18, C.white, false);
    addText(slide, r[2], 380, y + 42, 775, 20, 13, C.muted, false);
    addRule(slide, 64, y + 68, 1152, C.line, 1);
  });
  setNotes(slide, "Full source URLs: https://aclanthology.org/2025.findings-emnlp.1146/ ; https://aclanthology.org/2025.acl-long.373/ ; https://learn.microsoft.com/en-us/dotnet/machine-learning/resources/metrics ; https://learn.microsoft.com/en-us/dotnet/machine-learning/resources/tasks ; EACL theme: https://2027.eacl.org/calls/papers/ ; COLING theme: https://2027.coling-iccl.org/calls/main_conference_papers/ ; NAACL topics/theme: https://2027.naacl.org/calls/main_conference_papers/ ; ACL homogenization theme: https://2027.aclweb.org/calls/main/ ; EMNLP 2027 CFP not yet published. Submission timing: https://2027.eacl.org/calls/papers/ ; https://2027.coling-iccl.org/ ; https://2027.naacl.org/calls/main_conference_papers/ ; https://2027.aclweb.org/calls/main/. React sources: https://react.dev/learn/adding-interactivity and https://react.dev/learn/sharing-state-between-components.");
}

// 22. Modern model architecture references
{
  const slide = baseSlide("Model research: multilingual encoders to temporal graphs", "Model architecture", 22);
  const refs = [
    ["XLM-R", "Multilingual encoder candidate for aligned English/Romanian transfer; evaluate native-language held-out works.", "Conneau et al. 2020 · ACL"],
    ["TIMERS", "Document-level temporal relation extraction; contextual, discourse, and time-aware graph encoders.", "Mathur et al. 2021 · ACL"],
    ["Span + graph transformer", "Strong graph-architecture precedent, but the published results use clinical data and do not transfer automatically to narrative.", "Chaturvedi et al. 2025 · ACL"],
    ["Global temporal graphs", "Predict and constrain document-level event graphs; compare with local pairwise classification.", "Eirew et al. 2025 · EMNLP"],
    ["Transformer-era survey", "Review of 32 temporal-IE datasets across 15 languages; documents source/domain and schema gaps.", "Su, Howard & Bethard 2025 · EMNLP"],
  ];
  refs.forEach((r, i) => {
    const y = 157 + i * 90;
    addText(slide, r[0], 64, y, 265, 25, 18, C.mint, true);
    addText(slide, r[1], 340, y - 1, 810, 48, 17, C.white, false);
    addText(slide, r[2], 340, y + 49, 810, 20, 13, C.muted, false);
    addRule(slide, 64, y + 69, 1152, C.line, 1);
  });
  setNotes(slide, "Model papers: Conneau et al. (2020), XLM-R, https://aclanthology.org/2020.acl-main.747/ ; Mathur et al. (2021), TIMERS, https://aclanthology.org/2021.acl-short.67/ ; Chaturvedi et al. (2025), https://aclanthology.org/2025.acl-long.1251/ ; Eirew et al. (2025), https://aclanthology.org/2025.emnlp-main.1601/ ; Su, Howard & Bethard (2025), https://aclanthology.org/2025.emnlp-main.1467/. ML.NET supports ONNX inference for some external models, but not every transformer/graph export; validate opsets, tokenizer, span outputs, and runtime support before promising deployment. The current trainable pipeline remains ML.NET SDCA until a working multilingual neural training path and suitable data are available.");
}

await fs.mkdir(TMP_DIR, { recursive: true });
await fs.mkdir(path.dirname(FINAL_PPTX), { recursive: true });
const candidatePath = path.join(TMP_DIR, "candidate.pptx");
await (await PresentationFile.exportPptx(presentation)).save(candidatePath);
for (let i = 0; i < presentation.slides.items.length; i++) {
  const slide = presentation.slides.items[i];
  const preview = await presentation.export({ slide, format: "png", scale: 1 });
  await fs.writeFile(path.join(TMP_DIR, "slide-" + String(i + 1).padStart(2, "0") + ".png"), new Uint8Array(await preview.arrayBuffer()));
}
const { finalizePresentation } = await import(pathToFileURL(path.join(SKILL_DIR, "container_tools/artifact_tool_utils.mjs")).href);
const stagingDir = path.join(workspaceDir, "tym_deck_validation");
await fs.mkdir(stagingDir, { recursive: true });
const result = await finalizePresentation({
  explicitTotalSlideCount: 22,
  requiredNativeTableOwnerSlides: [],
  requiredNativeChartOwnerSlides: [],
  workspaceDir,
  candidatePath,
  finalPath: FINAL_PPTX,
  pythonExecutable: "C:\\Users\\serba\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\python\\python.exe",
  integrityValidatorPath: path.join(SKILL_DIR, "container_tools/inspect_presentation_package_integrity.py"),
  layoutValidatorPath: path.join(SKILL_DIR, "container_tools/inspect_presentation_layout_geometry.py"),
  layoutArgs: ["--expected-slide-size-emu", "12192000,6858000", "--validate-bullet-geometry", "--validate-heading-fit"],
  fontPolicy: { basis: "design", families: [FONT] },
  verifyArtifactToolImport: true,
  receiptPath: path.join(stagingDir, "TYM_MLNET_Synthetic_Research_Roadmap_2027_RoTimeBank_Annotation_Pilot.validation.json"),
});
console.log(JSON.stringify({ candidatePath, finalPath: FINAL_PPTX, font: FONT, slideCount: presentation.slides.items.length, finalizeResult: result }, null, 2));

