import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { Presentation, PresentationFile } from '@oai/artifact-tool';

const workspaceDir = 'C:/Users/serba/Documents/ChatGPT/PhD';
const SKILL_DIR = 'C:/Users/serba/.codex/plugins/cache/openai-primary-runtime/presentations/26.909.12148/skills/presentations';
const TMP_DIR = path.join(workspaceDir, 'tym-private-research/presentation-build-20261002');
const FINAL_PPTX = path.join(workspaceDir, 'output/TYM_MLNET_Updated_Results_20261002.pptx');
const RUNTIME_PYTHON = 'C:/Users/serba/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe';
process.env.RUNTIME_NODE_MODULES='C:/Users/serba/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules';
process.env.RUNTIME_NODE='C:/Users/serba/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node.exe';
process.env.RUNTIME_PYTHON=RUNTIME_PYTHON;
process.env.RUNTIME_BIN_DIR='C:/Users/serba/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override';
const runDir = path.join(workspaceDir, 'tym-private-research/mlnet-run-20261002');
const { resolvePresentationFont, applyPresentationChartFont, makeNativeBulletParagraphs, finalizePresentation } =
  await import(pathToFileURL(path.join(SKILL_DIR, 'container_tools/artifact_tool_utils.mjs')).href);

// Arial is installed in C:/Windows/Fonts. The policy records a design choice.
const family = resolvePresentationFont({ fontFamily: 'Arial' });
const fontPolicy = { basis: 'design', families: [family] };
const navy = '#142D43', teal = '#167D8D', ink = '#19364A', gray = '#536979';
const white = '#FFFFFF', paper = '#F8FAFB', line = '#D9E2E7';
const pres = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const src = x => path.join(runDir, x);
const readJson = async p => JSON.parse((await fs.readFile(p, 'utf8')).replace(/^\uFEFF/, ''));
const names = [
  'segment_type_en', 'segment_type_ro', 'temporal_relation_en', 'temporal_relation_ro',
  'timebank_event_class_ro', 'timebank_event_tense_ro', 'timebank_timex_type_ro',
  'timebank_tlink_ro', 'timebank_slink_ro', 'timebank_alink_ro',
];
const trainings = await Promise.all(names.map(n => readJson(src(`training/${n}.json`))));
const predictions = Object.fromEntries(await Promise.all(names.map(async n => [n, await readJson(src(`predictions/${n}.json`))])));
if (trainings.some(o => o.status !== 'trained')) throw new Error('Incomplete training evidence');
const totalRows = trainings.reduce((sum, o) => sum + o.training_rows, 0);
if (totalRows !== 21666) throw new Error(`Unexpected training total ${totalRows}`);
const importReport = await readJson(src('ro-timebank-import-report.json'));
const benchmark = await readJson(src('grouped-mlnet-benchmark.json'));
const structuredBenchmark=await readJson(src('grouped-mlnet-structured-benchmark.json'));
const structuredByTask=Object.fromEntries(structuredBenchmark.tasks.map(t=>[t.task,t]));
const benchmarkByTask = Object.fromEntries(benchmark.tasks.map(t=>[t.task,t]));
const benchmarkOrder=['timebank_event_class','timebank_event_tense','timebank_timex_type','timebank_tlink','timebank_slink','timebank_alink'];
const taskTitles={'timebank_event_class':'EVENT class','timebank_event_tense':'EVENT tense','timebank_timex_type':'TIMEX3 type','timebank_tlink':'TLINK','timebank_slink':'SLINK','timebank_alink':'ALINK'};
const booksReport = await readJson(src('books-import-report.json'));
const structuredTraining=await readJson(src('structured-full-training.json'));
const structuredSnapshots=await readJson(src('structured-free-text-results.json'));
const apiVerification=await readJson(src('corpus-api-public-verification.json'));
const apiReadiness=await readJson(src('corpus-api-public-ready.json'));
if(apiVerification.free_text_snapshots_matched!==32 || !apiVerification.source_preserved || !apiVerification.spans_exact || apiReadiness.ready_models!==16 || !apiReadiness.all_integrity_verified) throw new Error('Incomplete API publication evidence');
if(structuredSnapshots.length!==12) throw new Error('Incomplete structured free-text observations');
if(structuredTraining.status!=='full_data_inference_artifacts' || structuredTraining.models.length!==6 || structuredTraining.models.reduce((n,x)=>n+x.manifest.training_rows,0)!==20151) throw new Error('Incomplete structured training evidence');
if(benchmark.total_rows!==20151 || benchmark.tasks.length!==6) throw new Error('Unexpected benchmark coverage');
if (importReport.documents !== 157 || importReport.output_rows !== 20151) throw new Error('Unexpected retained corpus counts');
const trx = await fs.readFile(src('test-results/standalone-modeling-tests.trx'), 'utf8');
if (!/<Counters[^>]*failed="0"/.test(trx)) throw new Error('Modeling suite failures');
const apiTrx=await fs.readFile(src('test-results/standalone-api-tests.trx'),'utf8');
const researchTrx=await fs.readFile(src('test-results/private-book-data-tests.trx'),'utf8');
if(!/<Counters[^>]*failed="0"/.test(apiTrx) || !/<Counters[^>]*failed="0"/.test(researchTrx)) throw new Error('API/research suite failures');
const passed=x=>Number(x.match(/<Counters[^>]*passed="(\d+)"/)[1]);
const modelPassed=passed(trx),apiPassed=passed(apiTrx),researchPassed=passed(researchTrx),allPassed=modelPassed+apiPassed+researchPassed;
let deployment = { verified: false, ui: 'https://tym-corpus-ui-serban.livelyrock-2726c024.eastus.azurecontainerapps.io', api: apiVerification.api_url, description:'Public API ready with 16 loaded models and 32 matching snapshots. UI publication awaits verification.' };
try { deployment = await readJson(path.join(TMP_DIR, 'deployment.json')); } catch (e) { if (e.code !== 'ENOENT') throw e; }

function textbox(s, text, x, y, w, h, size=28, color=ink, bold=false) {
  const sh = s.shapes.add({ geometry: 'textbox', position: {left:x, top:y, width:w, height:h}, fill:'none', line:{fill:'none', width:0} });
  sh.text = text;
  sh.text.style = {typeface:family, fontSize:size, color, bold, autoFit:'none', wrap:'square', verticalAlignment:'top', insets:{left:0,right:0,top:0,bottom:0}};
  return sh;
}
function slide(title, notes='') {
  const s = pres.slides.add(); s.background.fill = white;
  textbox(s,title,72,44,1136,100,44,navy,true);
  s.speakerNotes.textFrame.setText(notes);
  return s;
}
function caption(s,t,y=633) { textbox(s,t,72,y,1136,62,21,gray); }
function table(s, values, widths, y=169, h=410, fontSize=24) {
  const t = s.tables.add({rows:values.length, columns:values[0].length, left:72,top:y,width:1136,height:h,columnWidths:widths,values});
  t.styleOptions = {headerRow:true,bandedRows:false};
  t.borders.assign({style:'solid',fill:line,width:1});
  const all = t.cells.block({row:0,column:0,rowCount:values.length,columnCount:values[0].length});
  all.assign({fill:white,textStyle:{typeface:family,fontSize,color:ink},margins:{left:16,right:14,top:11,bottom:8},anchor:'middle'});
  t.cells.block({row:0,column:0,rowCount:1,columnCount:values[0].length}).assign({fill:navy,textStyle:{typeface:family,fontSize,bold:true,color:white}});
  for(let r=1;r<values.length;r++) if(r%2===0) t.cells.block({row:r,column:0,rowCount:1,columnCount:values[0].length}).fill=paper;
  return t;
}
function chart(s,categories,values,position,{color=teal,max,unit=1000}={}) {
  const ch = s.charts.add('bar',{
    position,categories,series:[{name:'Examples',values,fill:color,valuesFormatCode:'#,##0'}],
    barOptions:{direction:'bar',grouping:'clustered',gapWidth:55},hasLegend:false,
    xAxis:{visible:true,min:0,max,majorUnit:unit,numberFormatCode:'#,##0',textStyle:{fontSize:22,fill:gray},line:{fill:line,width:1},majorGridlines:{fill:line,width:1}},
    yAxis:{visible:true,textStyle:{fontSize:22,fill:ink},line:{fill:'none',width:0},majorGridlines:null},
    dataLabels:{showValue:true,position:'outEnd',textStyle:{fontSize:22,fill:ink,bold:true}},
    chartFill:white,plotAreaFill:white,chartLine:{fill:'none',width:0},plotAreaLine:{fill:'none',width:0},
  });
  applyPresentationChartFont(ch,{fontFamily:family});
  return ch;
}

// 1. Minimal academic cover.
{
  const s=pres.slides.add();s.background.fill=navy;
  textbox(s,'TYM and ML.NET',72,160,1100,95,64,white,true);
  textbox(s,'Grouped diagnostics and narrative annotation',72,285,1090,85,39,white);
  textbox(s,'2 October 2026',72,540,900,50,27,'#B6D9E1');
  s.speakerNotes.textFrame.setText('Sources: supplied research materials and the verified local ML.NET run dated 2026-10-02. This presentation distinguishes corpus training and component tests from an independent scientific evaluation.');
}
// 2. Conceptual model.
{
 const s=slide('The Time Yards Model',`Sources: supplied Deep Learning manuscript pp. 5-9 and Practical Results pp. 3-12. Current workbench supports draft fields and reviewer edits. Event coreference, actor identities and full perspective attribution remain unimplemented.`);
 table(s,[['Unit','Meaning in the narrative'],['Event','An event mention with actor, location, action and temporal anchor'],['Time segment (TS)','A contiguous narrative span with a shared temporal perspective'],['Time track (TT)','Segments arranged along a narrative strand'],['Temporal relation (TREL)','A directed constraint between segment endpoints']], [355,781],168,400,26);
 caption(s,'Text order and story order can differ. Explicit anchors and reviewed relations preserve partial chronology.');
}
// 3. Historical evidence clearly separated.
{
 const s=slide('Research evidence and current implementation',`Sources: project AGENTS.md and README.md summarize the supplied paper's historical CNN result. Current evidence: ${runDir}/training/*.json and test-results/mlnet-all-tests.trx. No equivalence or reproduction claim links these experiments.`);
 table(s,[['Evidence','Dataset or scope','Result'],['Historical CNN in the paper','2,000 events\nPast, present, future','58.9% accuracy\nabout 58.7% macro F1'],['Current corpus training','Ten baselines and six\nstructured variants','21,666 unique task rows\nSaved ML.NET artifacts'],['Current grouped diagnostic','Six TimeBank label tasks\nSupplied spans and endpoints','20,151 out-of-fold rows\nUnknown annotation adjudication']], [350,375,411],170,400,25);
 caption(s,'The current run uses lexical ML.NET classifiers. The historical CNN metrics describe the paper’s experiment.');
}
// 4. Audit.
{
 const s=slide('Corpus provenance and import audit',`Sources: ${path.join(workspaceDir,'tym-private-research/import-report.json')}; ${src('ro-timebank-import-report.json')}; tools/Tym.Modeling/ro-timebank-exclusions.json. Ro-TimeBank archive SHA256: 6aa0ba361274b2f13a60e9da0c883c380fe0e959297ef86862a91ff3f499d878. No private corpus text appears on these slides.`);
 table(s,[['Corpus','Retained material','Audit decision'],['TYM parallel narrative','One story in two languages\n751 segments and 764 relations','Duplicate XML variants add no rows\n1,515 task examples'],['Ro-TimeBank ISO-TimeML','157 of 183 documents\n20,151 task examples','26 mixed documents quarantined\n1 ambiguous signal link skipped'],['Annotation provenance','provided_annotation\nadjudication_unknown','Exploratory training/diagnostics\nHuman adjudication unverified']], [290,406,440],165,425,23);
 caption(s,'The paired TYM translations share one story. ISO-TimeML labels keep separate task namespaces.');
}
// 5. All ten model fits.
{
 const s=slide('Sixteen supervised classifiers',`Sources: ${src('training')}/*.json and structured full-data training reports. Ten baseline classifiers plus six structured-channel TimeBank classifiers reuse 21,666 unique task rows. Rows count annotation examples, not independent narratives. No SUP segment examples occur in the supplied annotation pair.`);
 const values=[['Task','Lang.','Rows','Classes','Variants'],...trainings.map(o=>[o.task.replace('timebank_','TimeBank '),o.language,String(o.training_rows.toLocaleString('en-US')),String(o.labels.length),o.task.startsWith('timebank_')?'whole + structured':'whole'])];
 const t=table(s,values,[430,130,170,150,256],150,466,23);
 t.cells.block({row:0,column:0,rowCount:values.length,columnCount:5}).assign({margins:{left:16,right:14,top:5,bottom:5}});
 for(let r=0;r<values.length;r++) t.rows[r].height=42;
 caption(s,'Four TYM models and two variants of each TimeBank task. Two additional K-Means artifacts remain private.',638);
}
// 6. Native charts for coverage.
{
 const s=slide('Training coverage is uneven',`Sources: TYM import-report.json segment_labels summed across the supplied English/Romanian pair: NAR=627, REM=66, FIC=50, GEN=8, total=751. Ro-TimeBank task counts from ${src('ro-timebank-import-report.json')}: event_class=5947, event_tense=5920, timex_type=1144, tlink=4808, slink=2133, alink=199. Both charts retain native editable data and embedded workbook snapshots.`);
 textbox(s,'TYM segment labels',72,157,470,42,28,navy,true);
 textbox(s,'Ro-TimeBank task examples',640,157,568,42,28,navy,true);
 chart(s,['NAR','REM','FIC','GEN'],[627,66,50,8],{left:65,top:215,width:510,height:350},{max:720,unit:200});
 chart(s,['Class','Tense','TIMEX','TLINK','SLINK','ALINK'],[5947,5920,1144,4808,2133,199],{left:628,top:212,width:590,height:370},{max:7000,unit:2000});
 caption(s,'NAR accounts for 83.5% of the supplied TYM segments. ALINK has 199 examples across five classes.');
}
// Additional raw books remain explicitly unlabeled.
{
 const s=slide('Unlabeled books and private discovery',`Sources: ${src('books-import-report.json')}, books_ro_cluster_report.json, books_en_cluster_report.json and books_*_segment_predictions.json. C# preparation: tym-corpus-app/tools/Tym.Corpus.Data/BookImporter.cs. No source-book passages appear in this deck. The separate Tash Aw raw editions have zero exact TS offset matches and need reconciliation before span evaluation.`);
 table(s,[['Pool','Files / groups','Passages','Private experiment'],['Romanian raw books','26 files / 14 groups','5,213','K-Means with 8 clusters\nDistance 0.9019, DBI 6.3917'],['English raw novel','1 file / 1 group','446','K-Means with 8 clusters\nDistance 0.7985, DBI 5.1398'],['RO classifier probes','Unlabeled chunks','5,213','NAR 4,734  REM 295\nFIC 183  GEN 1'],['EN classifier probes','Unlabeled chunks','446','NAR 404  REM 24\nFIC 16  GEN 2']], [285,270,166,415],160,430,24);
 caption(s,'DBI is the Davies-Bouldin index. Fitted geometry and chunk label counts provide no gold semantics or accuracy.');
}
// Pipeline and marking.
{
 const s=slide('ML.NET inference and annotation context',`Sources: tym-corpus-app/tools/Tym.Modeling/Program.cs, TextModelPredictor.cs, tools/Tym.Corpus.Data and src/Tym.Corpus.Api. The illustrative target-marked sentence is original free text. The independent solution targets .NET 10 LTS with ML.NET 5.0.0, while initial full-data model artifacts precede the framework migration. No transformer participates in these ten fits.`);
 table(s,[['Stage','Implemented behavior'],['Prepare in C#','Preserve label families and source provenance\nKeep directed relation endpoints and marked target context'],['Train with ML.NET','FeaturizeText n-gram features\nSdcaMaximumEntropy multiclass classifier'],['Predict from artifacts','Load a saved model and score task-specific free text\nReturn label inventory and provenance'],['Draft document analysis','Sentence/rule candidates and segment classifier\nImmutable source offsets and explicit date anchors']], [280,856],160,420,25);
 caption(s,'Example: Context: Mâine Mara [TARGET] va pleca [/TARGET] spre munte. Event: va pleca');
}
// Document grouping defines the conditional diagnostic.
{
 const s=slide('Grouped TimeBank diagnostic',`Source: ${src('grouped-mlnet-benchmark.json')}; tym-corpus-app/tools/Tym.Benchmark/Program.cs and GroupedSplit.cs. Featurizer/label mapping fit only training rows. Algorithms and 100-iteration settings were fixed before examining scores. This retrospective diagnostic has no separate model-selection holdout and exports no model weights.`);
 table(s,[['Protocol','Executed scope'],['Input','20,151 rows from 157 documents\nprovided_annotation / adjudication_unknown'],['Grouping','Exact normalized task/language feature duplicates join documents\n50 components with 107 cross-document duplicate feature rows'],['Validation','3 folds with 6,744, 6,721 and 6,686 rows\nEvery row receives one out-of-fold prediction'],['Whole-text comparison','Training-fold majority, SDCA and L-BFGS maximum entropy\nFeaturizeText, at most 100 iterations, one trainer thread']], [295,841],162,425,24);
 caption(s,'Scores classify labels given supplied spans and endpoints. TYM has one source work and remains outside this split.');
}
// Editable metric table retains the whole-text comparisons.
{
 const s=slide('Baseline classification results',`Source: ${src('grouped-mlnet-benchmark.json')}. Pooled out-of-fold accuracy and fixed-inventory macro F1 percentages. Algorithms sdca/lbfgs/majority are trained independently for each task and fold. Task label inventories differ, so no combined NLP accuracy is claimed.`);
 const vals=[['Task','Rows','Majority\nAcc. / F1 %','SDCA\nAcc. / F1 %','L-BFGS\nAcc. / F1 %'],...benchmarkOrder.map(k=>{const o=benchmarkByTask[k];return [taskTitles[k],o.rows.toLocaleString('en-US'),...['majority','sdca','lbfgs'].map(a=>`${(o.algorithms[a].accuracy*100).toFixed(2)} / ${(o.algorithms[a].macro_f1*100).toFixed(2)}`)];})];
 const t=table(s,vals,[230,115,260,260,271],160,433,23);
 t.cells.block({row:0,column:0,rowCount:vals.length,columnCount:5}).assign({margins:{left:14,right:10,top:5,bottom:5}});
 for(let r=0;r<vals.length;r++) t.rows[r].height=r===0?67:61;
 caption(s,'Each cell gives accuracy / macro F1 in percent. TLINK remains weak and DATE dominance masks minority errors.');
}
// Exploratory structured channels preserve mention/endpoint roles.
{
 const s=slide('Structured target and endpoint features',`Source: ${src('grouped-mlnet-structured-benchmark.json')}, algorithms.structured_sdca. Same input and document/duplicate-feature folds as the whole-text comparison. Separate target/context and ordered From/To/Signal channels improve all six pooled scores. A channel empty in training is omitted using training data only. This variant was added after examining the initial comparison. It is exploratory and has no independent model-selection holdout.`);
 const vals=[['Task','Accuracy %','Macro F1 %','F1 gain in points'],...benchmarkOrder.map(k=>{const m=structuredByTask[k].algorithms.structured_sdca;return [taskTitles[k],(m.accuracy*100).toFixed(2),(m.macro_f1*100).toFixed(2),'+'+((m.macro_f1-benchmarkByTask[k].algorithms.sdca.macro_f1)*100).toFixed(2)];})];
 const t=table(s,vals,[325,240,240,331],158,403,25);
 t.cells.block({row:0,column:0,rowCount:vals.length,columnCount:4}).assign({margins:{left:14,right:10,top:6,bottom:6}});
 for(let r=0;r<vals.length;r++) t.rows[r].height=57;
 caption(s,'Exploratory follow-up with no independent selection holdout. Gains use unrounded scores.\nSET, DURING_INV, IAFTER and IBEFORE still have zero F1.',605);
}
// Tests.
{
 const s=slide(`${allPassed} .NET test cases pass`,`Sources: ${src('test-results/standalone-modeling-tests.trx')} (${modelPassed}), standalone-api-tests.trx (${apiPassed}), private-book-data-tests.trx (${researchPassed}). API cases cover typed spans, explicit DCT and limited partial-graph constraints. Research cases cover C# preparation, grouping/metrics and actual private Books.zip checks. Snapshots record model behavior including qualitative errors. Counts measure software assertions, not linguistic accuracy. TypeScript/browser checks are recorded separately.`);
 chart(s,['Modeling and artifact checks','API and workbench','Corpus and benchmark'],[modelPassed,apiPassed,researchPassed],{left:72,top:182,width:1120,height:355},{max:Math.ceil(Math.max(modelPassed,apiPassed,researchPassed)/10)*10+10,unit:10});
 textbox(s,'All three suites report zero failures',72,566,1100,47,30,teal,true);
 caption(s,'Passing snapshots verifies recorded outputs. Conditional grouped NLP scores come from separate validation rows.',636);
}
// 9. Positive qualitative examples.
{
 const s=slide('Selected free-text predictions',`Sources: ${src('predictions/segment_type_en.json')}, timebank_event_tense_ro.json, timebank_timex_type_ro.json and timebank_alink_ro.json. All example text is original free text supplied as test input. The descriptions of agreement are qualitative readings, not adjudicated gold labels.`);
 table(s,[['Input cue or mention','Task','Prediction'],['Mara remembered the blue room','TYM segment, English','REM'],['Mâine Mara va pleca spre munte','TimeBank tense','FUTURE'],['primăvara trecută','TimeBank TIMEX type','DATE'],['trei zile','TimeBank TIMEX type','DURATION'],['Ploaia a început să cadă','TimeBank ALINK','INITIATES']], [570,350,216],168,410,26);
 caption(s,'These authored examples show outputs of the ten saved baseline classifiers. They carry no adjudicated gold labels.');
}
// 10. Honest errors.
{
 const s=slide('Baseline errors and structured probes',`Sources: ${src('predictions/timebank_tlink_ro.json')} case free-ro-before-link, ${src('predictions/segment_type_ro.json')} case free-ro-memory, ${src('predictions/temporal_relation_en.json')} case free-en-concurrent and ${src('structured-free-text-results.json')}. The baseline before cue yields AFTER, while the structured variant yields BEFORE. The structured SLINK fixture containing știa yields MODAL versus baseline EVIDENTIAL. All snapshot expected_label fields represent observed regression labels, not human gold. Cue-based interpretations here are qualitative diagnoses.`);
 table(s,[['Cue and task','Observed label','Qualitative reading'],['Romanian TLINK with “înainte”\nDeparture before the later arrival','AFTER','The stated ordering suggests BEFORE'],['Romanian segment with “și-a amintit”\nMara recalls her childhood room','NAR','The memory cue suggests REM'],['English TREL with “at the same time”\nRain stops as the ferry arrives','AFTER','The concurrent cue suggests SIMULTANEOUS']], [560,210,366],172,388,25);
 caption(s,'Structured TLINK changes the before fixture to BEFORE. Its “știa” SLINK fixture changes to MODAL.\nThese changes remain qualitative observations without adjudicated gold labels.');
}
// Workbench derives its scope from the research papers.
{
 const s=slide('Narrative annotation workbench',`Sources: supplied updated paper pp. 5-9, State of the Art pp. 10-13, tym-corpus-app/src/Tym.Corpus.Api and TypeScript frontend source. Current analysis uses lexical candidates, sentence rules and an existing segment-mode model. It has no event coreference, actor identity resolution, full interval closure or automatic causal relation extraction. The validator checks schema, typed endpoints, source hash/span validity and limited strict-cycle/simultaneous contradictions. Human edits never automatically become human_gold.`);
 table(s,[['Layer','Implemented behavior'],['Immutable source','Half-open UTF-16 spans and source hash\nSeparate normalized corpus input from reviewer text'],['Typed review','TS, TIMEX and relation editors\nActor, location and perspective fields with evidence'],['Temporal anchors','Explicit document creation time and local cues\nNo server-date substitution'],['Graph views and checks','Review supplied directed edges and simultaneous groups\nPreserve partial order without forced adjacency or causality']], [285,851],162,430,25);
 caption(s,'Draft analysis proposes annotations. Human adjudication and end-to-end extraction evaluation remain future work.');
}
// Research plan.
{
 const s=slide('Independent evaluation and next experiments',`Sources: AGENTS.md; data/README.md; pocs/model-evaluation/README.md; docs/research/2027-corpus-and-synthetic-data-plan.md. These are proposed research steps. The current CLI evaluates classification only and does not yet evaluate event spans, TS boundaries or complete graph consistency. Chapter/document splits should group translations and synthetic descendants.`);
 const sh=textbox(s,'',72,175,1136,385,28,ink);
 sh.text=makeNativeBulletParagraphs([
  'Collect adjudicated annotations from independent narratives in both languages.',
  'Keep parallel translations, chapters from one work and synthetic descendants in the same split group.',
  'Report accuracy, macro F1 and per-class precision/recall. Evaluate segment boundaries and temporal graphs.',
  'Compare the lexical baseline with contextual and sequence models after implementing and evaluating them.',
 ],{marginLeftPoints:18,hangingPoints:9,spaceAfterPoints:18});
 sh.text.style={typeface:family,fontSize:28,color:ink,autoFit:'none',wrap:'square',insets:{left:0,right:0,top:0,bottom:0}};
 textbox(s,'Human gold is the prerequisite for a generalization claim',72,598,1110,65,28,teal,true);
}
// Final slide awaits verified deployment details from root.
{
 const s=slide(deployment.verified ? 'Azure publication' : 'Azure publication in progress',`Deployment source: root agent verification and Azure resource inventory. Publication state is ${deployment.verified?'verified':'pending verification'}. ${deployment.notes??''}`);
 textbox(s,deployment.verified ? 'Live research prototype' : 'Planned publication targets',72,168,1136,54,34,teal,true);
 textbox(s,'Web interface',72,258,1080,40,27,navy,true);
 const ui=textbox(s,deployment.ui,72,311,1130,62,25,ink);ui.text.get(deployment.ui).link={uri:deployment.ui,isExternal:true};
 textbox(s,'API',72,425,1080,40,27,navy,true);
 const api=textbox(s,deployment.api,72,478,1130,63,25,ink);api.text.get(deployment.api).link={uri:deployment.api,isExternal:true};
 caption(s,deployment.description??'Deployment verification will establish the live API behavior and available model endpoints.',620);
}

await fs.mkdir(TMP_DIR,{recursive:true});
for(let i=0;i<pres.slides.items.length;i++) {
 const s=pres.slides.items[i];
 const png=await pres.export({slide:s,format:'png',scale:1});
 await fs.writeFile(path.join(TMP_DIR,`slide-${String(i+1).padStart(2,'0')}.png`),new Uint8Array(await png.arrayBuffer()));
 const layout=await s.export({format:'layout'});
 await fs.writeFile(path.join(TMP_DIR,`slide-${String(i+1).padStart(2,'0')}.layout.json`),await layout.text());
}
await fs.writeFile(path.join(TMP_DIR,'deck.proto.json'),JSON.stringify(pres.toProto()));
if(process.argv.includes('--finalize') || process.argv.includes('--validate-draft')) {
 const draftOnly=process.argv.includes('--validate-draft');
 if(!draftOnly && !deployment.verified) throw new Error('Final deployment verification is required before export');
 const candidate=path.join(TMP_DIR,'candidate.pptx');
 await (await PresentationFile.exportPptx(pres)).save(candidate);
 const draftStamp=Date.now();
 const exportFinal=draftOnly?path.join(TMP_DIR,`draft-output/validated-draft-${draftStamp}.pptx`):FINAL_PPTX;
 await fs.mkdir(path.dirname(exportFinal),{recursive:true});
 const result=await finalizePresentation({workspaceDir,candidatePath:candidate,finalPath:exportFinal,pythonExecutable:RUNTIME_PYTHON,
  integrityValidatorPath:path.join(SKILL_DIR,'container_tools/inspect_presentation_package_integrity.py'),
  layoutValidatorPath:path.join(SKILL_DIR,'container_tools/inspect_presentation_layout_geometry.py'),
  layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-bullet-geometry','--validate-heading-fit',...[2,3,4,5,7,8,9,10,11,13,14,15].flatMap(n=>['--require-native-table-slide',String(n)])],
  explicitTotalSlideCount:17,requiredNativeTableOwnerSlides:[2,3,4,5,7,8,9,10,11,13,14,15],requiredNativeChartOwnerSlides:[6,12],
  materializeLiteralChartWorkbooks:true,fontPolicy,verifyArtifactToolImport:true,
  receiptPath:path.join(TMP_DIR,draftOnly?`draft-${draftStamp}.validation.json`:'final.validation.json')});
 console.log(JSON.stringify(result));
} else console.log(JSON.stringify({slides:pres.slides.items.length,previewDirectory:TMP_DIR,awaitingDeployment:true}));
