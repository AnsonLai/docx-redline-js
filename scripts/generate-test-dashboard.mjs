import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'fs';
import { dirname, join, resolve } from 'path';
import { fileURLToPath } from 'url';

import {
    COVERAGE_ORACLES,
    COVERAGE_STRUCTURES,
    COVERAGE_TASKS
} from './lib/word-coverage-metadata.mjs';
import { loadCoverageCatalogue } from './lib/word-coverage-catalogue.mjs';
import { selectVisualReviewCases } from './prepare-word-visual-review.mjs';

function parseArgs(argv) {
    const outputIndex = argv.indexOf('--output');
    if (outputIndex >= 0 && !argv[outputIndex + 1]) throw new Error('--output requires a path');
    const fixturesIndex = argv.indexOf('--fixtures-dir');
    if (fixturesIndex >= 0 && !argv[fixturesIndex + 1]) throw new Error('--fixtures-dir requires a path');
    const corpusIndex = argv.indexOf('--corpus-fixtures-dir');
    if (corpusIndex >= 0 && !argv[corpusIndex + 1]) throw new Error('--corpus-fixtures-dir requires a path');
    return {
        outputPath: outputIndex >= 0
            ? resolve(process.cwd(), argv[outputIndex + 1])
            : join(process.cwd(), 'docs', 'test-comparison-dashboard.html'),
        fixturesDir: fixturesIndex >= 0
            ? resolve(process.cwd(), argv[fixturesIndex + 1])
            : join(process.cwd(), 'tmp', 'dashboard-docx'),
        corpusFixturesDir: corpusIndex >= 0
            ? resolve(process.cwd(), argv[corpusIndex + 1])
            : null
    };
}

export function buildDashboardData(fixturesDir = null, corpusFixturesDir = null) {
    const { cases, priorities } = loadCoverageCatalogue();
    const corpusSuitePath = corpusFixturesDir ? join(corpusFixturesDir, 'suite.json') : null;
    const corpusSuite = corpusSuitePath && existsSync(corpusSuitePath)
        ? JSON.parse(readFileSync(corpusSuitePath, 'utf8'))
        : { cases: [] };
    const corpusNames = new Map(corpusSuite.cases.map(item => [item.scenarioKey || item.sourceId, item.name]));
    const visualEligible = new Set(
        selectVisualReviewCases().map(testCase => `synthetic:${testCase.name}`)
    );
    return {
        generatedAt: new Date().toISOString(),
        tasks: COVERAGE_TASKS,
        structures: COVERAGE_STRUCTURES,
        oracles: COVERAGE_ORACLES,
        priorities,
        cases: cases.map(item => {
            const syntheticName = item.identity.startsWith('synthetic:')
                ? item.identity.slice('synthetic:'.length)
                : null;
            const sourceId = item.identity.startsWith('superdoc:')
                ? item.identity.slice('superdoc:'.length)
                : null;
            const name = syntheticName || corpusNames.get(sourceId) || null;
            const variantDir = syntheticName ? fixturesDir : corpusFixturesDir;
            const readBase64 = suffix => {
                const path = name && variantDir ? join(variantDir, `${name}${suffix}.docx`) : null;
                return path && existsSync(path) ? readFileSync(path).toString('base64') : null;
            };
            const expectedPath = name && variantDir ? join(variantDir, `${name}.expected.json`) : null;
            const expected = expectedPath && existsSync(expectedPath)
                ? JSON.parse(readFileSync(expectedPath, 'utf8'))
                : null;
            const documentXmlPath = name && variantDir ? join(variantDir, `${name}.document.xml`) : null;
            const documentXml = documentXmlPath && existsSync(documentXmlPath)
                ? readFileSync(documentXmlPath, 'utf8')
                : '';
            return {
                identity: item.identity,
                displayName: sourceId
                    ? `${name} — ${expected?.originalTarget || item.detail}`
                    : name,
                lane: item.lane,
                category: item.category,
                task: item.metadata.task,
                structures: item.metadata.structures,
                oracles: item.metadata.oracles,
                manualReview: item.metadata.manualReview,
                visualEligible: visualEligible.has(item.identity) || Boolean(sourceId),
                docxVariants: name ? {
                    source: readBase64('.source'),
                    tracked: readBase64(''),
                    accepted: readBase64('.accepted'),
                    rejected: readBase64('.rejected')
                } : null,
                expectations: expected ? {
                    source: expected.sourceText || expected.originalTarget,
                    accepted: expected.expectedAcceptedText || expected.modifiedTarget,
                    rejected: expected.expectedRejectedText || expected.originalTarget
                } : null,
                revisions: {
                    insertions: (documentXml.match(/<w:ins\b/g) || []).length,
                    deletions: (documentXml.match(/<w:del\b/g) || []).length,
                    formatting: (documentXml.match(/<w:(?:rPrChange|pPrChange)\b/g) || []).length
                }
            };
        })
    };
}

function escapeJsonForHtml(value) {
    return JSON.stringify(value).replace(/</g, '\\u003c').replace(/>/g, '\\u003e').replace(/&/g, '\\u0026');
}

export function renderDashboardHtml(data, libraries = {}) {
    const encoded = escapeJsonForHtml(data);
    const jszipSource = String(libraries.jszipSource || '').replace(/<\/script/gi, '<\\/script');
    const docxPreviewSource = String(libraries.docxPreviewSource || '').replace(/<\/script/gi, '<\\/script');
    return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>DOCX Redline Test Comparison Dashboard</title>
<style>
:root{color-scheme:light dark;--bg:#f4f7fb;--surface:#fff;--surface2:#eef3f9;--text:#172033;--muted:#68758a;--line:#d9e1ec;--blue:#2563eb;--blue2:#dbeafe;--green:#138a5b;--green2:#d9f5e9;--amber:#b66a00;--amber2:#fff1cf;--red:#be3b45;--red2:#ffe2e5;--purple:#7656c8;--purple2:#eee8ff;--shadow:0 12px 30px rgba(38,55,80,.09)}
@media(prefers-color-scheme:dark){:root{--bg:#10141d;--surface:#181e29;--surface2:#222a37;--text:#edf2fb;--muted:#a8b3c5;--line:#323c4d;--blue:#78a7ff;--blue2:#223d68;--green:#58d2a0;--green2:#183f33;--amber:#ffc66d;--amber2:#533d1e;--red:#ff929a;--red2:#55282d;--purple:#b9a1ff;--purple2:#382f59;--shadow:none}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text);font:14px/1.45 Inter,ui-sans-serif,system-ui,-apple-system,"Segoe UI",sans-serif}button,select,input{font:inherit;color:inherit}.shell{max-width:1480px;margin:auto;padding:28px;transition:max-width .18s ease}.top{display:flex;justify-content:space-between;gap:24px;align-items:end;margin-bottom:22px}.eyebrow{text-transform:uppercase;letter-spacing:.12em;color:var(--blue);font-weight:700;font-size:11px}h1{margin:4px 0 2px;font-size:clamp(25px,4vw,40px);line-height:1.12}h2{font-size:18px;margin:0 0 14px}.stamp{color:var(--muted);text-align:right}.controls{display:flex;flex-wrap:wrap;gap:10px;margin:0 0 18px}.controls label{display:flex;align-items:center;gap:7px;background:var(--surface);border:1px solid var(--line);border-radius:10px;padding:8px 11px}.controls select,.controls input{border:0;background:transparent;outline:none}.stats{display:grid;grid-template-columns:repeat(6,minmax(120px,1fr));gap:12px;margin-bottom:18px}.stat,.panel{background:var(--surface);border:1px solid var(--line);border-radius:14px;box-shadow:var(--shadow)}.stat{padding:16px}.stat strong{display:block;font-size:25px;line-height:1.1}.stat span{color:var(--muted);font-size:12px}.layout{display:grid;grid-template-columns:minmax(0,1.65fr) minmax(300px,.7fr);gap:18px}.sidebar-hidden .shell{max-width:none}.sidebar-hidden .layout{grid-template-columns:minmax(0,1fr)}.sidebar-hidden #dashboard-sidebar{display:none}.panel{padding:18px;margin-bottom:18px}.panel-head{display:flex;align-items:start;justify-content:space-between;gap:12px}.sub{color:var(--muted);font-size:12px;margin-top:-8px;margin-bottom:14px}.matrix-wrap{overflow:auto}.matrix{display:grid;min-width:1040px;gap:4px;align-items:stretch}.matrix .label{font-size:11px;color:var(--muted);padding:8px 5px;display:flex;align-items:end}.matrix .row-label{justify-content:flex-end;text-align:right;align-items:center}.cell{border:0;border-radius:7px;min-height:42px;padding:4px;cursor:pointer;background:var(--surface2);display:flex;align-items:center;justify-content:center;font-weight:700}.cell:hover,.cell:focus{outline:2px solid var(--blue);outline-offset:1px}.cell.tested{background:var(--green2);color:var(--green)}.cell.planned{background:var(--amber2);color:var(--amber)}.cell.missing{background:var(--red2);color:var(--red)}.cell.empty{color:var(--muted);font-weight:400}.cell.selected{box-shadow:inset 0 0 0 3px currentColor}.legend{display:flex;flex-wrap:wrap;gap:15px;margin-top:12px;color:var(--muted);font-size:12px}.legend i{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:5px}.bars{display:grid;gap:9px}.bar-row{display:grid;grid-template-columns:130px minmax(0,1fr) 34px;gap:8px;align-items:center}.track{height:12px;background:var(--surface2);border-radius:99px;overflow:hidden;display:flex}.seg-syn{background:var(--blue)}.seg-real{background:var(--purple)}.bar-value{text-align:right;font-variant-numeric:tabular-nums}.detail{min-height:180px}.detail h3{font-size:16px;margin:0 0 8px}.detail ul{margin:8px 0 0;padding-left:18px;max-height:310px;overflow:auto}.detail code{font-size:11px;overflow-wrap:anywhere}.badge{display:inline-block;border-radius:99px;padding:3px 8px;margin:3px 4px 3px 0;background:var(--surface2);font-size:11px}.badge.syn{background:var(--blue2);color:var(--blue)}.badge.real{background:var(--purple2);color:var(--purple)}.gap{padding:11px 0;border-top:1px solid var(--line)}.gap:first-of-type{border-top:0}.gap strong{display:block}.gap p{margin:4px 0;color:var(--muted);font-size:12px}.oracle{display:grid;grid-template-columns:1fr 44px;gap:8px;align-items:center;margin:9px 0}.oracle .track{height:8px}.search-results{margin-top:8px}.case-row{display:grid;grid-template-columns:minmax(220px,1.4fr) 100px 130px;gap:12px;padding:9px 0;border-top:1px solid var(--line);align-items:center}.case-row:first-child{border-top:0}.case-row code{overflow-wrap:anywhere;font-size:11px}.empty-state{color:var(--muted);padding:18px 0}.docx-toolbar{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:14px}.docx-toolbar select{min-width:300px;max-width:100%;border:1px solid var(--line);background:var(--surface2);border-radius:8px;padding:7px 9px}.docx-status{color:var(--muted);font-size:12px}.docx-compare{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px}.docx-pane{min-width:0}.docx-pane h3{font-size:13px;margin:0 0 7px}.docx-view{background:var(--surface2);border:1px solid var(--line);border-radius:10px;min-height:330px;max-height:620px;overflow:auto}.docx-view .docx-wrapper{background:var(--surface2)!important;padding:12px!important}.docx-view .docx-wrapper>section.docx{background:#fff!important;color:#111!important;width:100%!important;min-height:380px!important;padding:44px!important;margin:0!important;box-shadow:none!important}.docx-view ins{background:#dcfce7;color:#166534;text-decoration:none}.docx-view del{background:#fee2e2;color:#991b1b}.screen-reader{position:absolute!important;width:1px;height:1px;padding:0;margin:-1px;overflow:hidden;clip:rect(0,0,0,0);white-space:nowrap;border:0}@media(max-width:980px){.stats{grid-template-columns:repeat(3,1fr)}.layout{grid-template-columns:1fr}.stamp{text-align:left}.top{align-items:start;flex-direction:column}}@media(max-width:760px){.docx-compare{grid-template-columns:1fr}}@media(max-width:620px){.shell{padding:16px}.stats{grid-template-columns:repeat(2,1fr)}.bar-row{grid-template-columns:100px minmax(0,1fr) 30px}.case-row{grid-template-columns:1fr}.controls label{width:100%;justify-content:space-between}.docx-toolbar select{min-width:0;width:100%}.docx-view .docx-wrapper>section.docx{padding:24px!important}}
.action{border:1px solid var(--line);background:var(--surface2);border-radius:8px;padding:7px 10px;cursor:pointer}.action:hover,.action:focus{border-color:var(--blue)}.action.primary{background:var(--blue);border-color:var(--blue);color:#fff}.preset-row{display:flex;gap:7px;flex-wrap:wrap;margin:0 0 14px}.docx-meta{display:flex;gap:8px;flex-wrap:wrap;margin:0 0 14px}.pane-head{display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:7px}.pane-head h3{margin:0}.pane-controls{display:flex;gap:7px}.pane-controls select{border:1px solid var(--line);background:var(--surface2);border-radius:7px;padding:5px}.expectations{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;margin-top:14px}.expectation{background:var(--surface2);border-radius:9px;padding:11px;white-space:pre-wrap;overflow-wrap:anywhere}.expectation strong{display:block;margin-bottom:5px}.sync-control{display:inline-flex;align-items:center;gap:5px}.case-row{grid-template-columns:minmax(220px,1.4fr) 100px 120px auto}@media(max-width:760px){.expectations{grid-template-columns:1fr}}@media(max-width:620px){.case-row{grid-template-columns:1fr}.pane-head{align-items:flex-start;flex-direction:column}.pane-controls{width:100%}.pane-controls select{flex:1}}
</style>
</head>
<body>
<main class="shell">
  <header class="top"><div><div class="eyebrow">Reliability coverage</div><h1>What is actually being tested?</h1><div class="sub">Task × structure × oracle comparison</div></div><div class="stamp" id="stamp"></div></header>
  <div class="controls" aria-label="Dashboard filters">
    <label>Lane <select id="lane"><option value="all">All cases</option><option value="synthetic">Synthetic Word</option><option value="superdoc">Real documents</option></select></label>
    <label>Category <select id="category"><option value="all">All categories</option><option value="legal">Legal</option><option value="administrative">Administrative</option></select></label>
    <label>Find case <input id="search" type="search" placeholder="Name or structure"></label>
    <button class="action" id="sidebar-toggle" type="button" aria-controls="dashboard-sidebar" aria-expanded="true">Hide sidebar</button>
  </div>
  <section class="stats" id="stats" aria-label="Coverage summary"></section>
  <div class="layout">
    <div>
      <section class="panel"><div class="panel-head"><div><h2>Task × structure matrix</h2><div class="sub">Select a cell to inspect its cases. Planned cells are high-priority gaps with recorded dependencies.</div></div></div><div class="matrix-wrap"><div class="matrix" id="matrix"></div></div><div class="legend"><span><i style="background:var(--green2)"></i>Tested</span><span><i style="background:var(--amber2)"></i>Planned high priority</span><span><i style="background:var(--red2)"></i>Unplanned high priority</span><span><i style="background:var(--surface2)"></i>Not prioritized</span></div></section>
      <section class="panel" id="docx-comparison"><h2>DOCX comparison workbench</h2><div class="sub">Compare source, tracked, accepted, and rejected states from synthetic fixtures and reviewed real legal/administrative documents. Tracked markup uses docxjs experimental revision rendering.</div><div class="docx-toolbar"><label for="docx-case">Document</label><select id="docx-case"></select><label class="sync-control"><input id="sync-scroll" type="checkbox" checked> Sync scroll</label><span class="docx-status" id="docx-status"></span></div><div class="preset-row"><button class="action primary" type="button" data-preset="source,tracked">Source ↔ tracked</button><button class="action" type="button" data-preset="source,accepted">Source ↔ accepted</button><button class="action" type="button" data-preset="accepted,rejected">Accepted ↔ rejected</button></div><div class="docx-meta" id="docx-meta"></div><div class="docx-compare"><div class="docx-pane"><div class="pane-head"><h3>Left document</h3><div class="pane-controls"><select id="left-view" aria-label="Left document state"></select><button class="action" id="left-download" type="button">Download</button></div></div><div class="docx-view" id="docx-left"></div></div><div class="docx-pane"><div class="pane-head"><h3>Right document</h3><div class="pane-controls"><select id="right-view" aria-label="Right document state"></select><button class="action" id="right-download" type="button">Download</button></div></div><div class="docx-view" id="docx-right"></div></div></div><div class="expectations"><div class="expectation"><strong>Expected rejected/source text</strong><span id="expected-before"></span></div><div class="expectation"><strong>Expected accepted text</strong><span id="expected-after"></span></div></div></section>
      <section class="panel"><h2>Coverage by structure</h2><div class="sub">Unique case count, split between synthetic packages and reviewed real documents.</div><div class="bars" id="structure-bars"></div></section>
      <section class="panel"><h2>Matching cases</h2><div class="sub" id="case-caption"></div><div class="search-results" id="case-list"></div></section>
    </div>
    <aside id="dashboard-sidebar">
      <section class="panel detail" id="detail"><h2>Cell detail</h2><div class="empty-state">Select a matrix cell to see exact scenarios and gap rationale.</div></section>
      <section class="panel"><h2>Oracle comparison</h2><div class="sub">How many filtered cases are checked by each independent oracle.</div><div id="oracles"></div></section>
      <section class="panel"><h2>Planned high-priority gaps</h2><div id="gaps"></div></section>
    </aside>
  </div>
</main>
<script id="dashboard-data" type="application/json">${encoded}</script>
<script>${jszipSource}</script>
<script>${docxPreviewSource}</script>
<script>
const DATA=JSON.parse(document.getElementById('dashboard-data').textContent);
const $=id=>document.getElementById(id);let selected=null;
const label=s=>s.replaceAll('-',' ').replace(/\\b\\w/g,c=>c.toUpperCase());
const dispositionMap=new Map(DATA.priorities.emptyCellDispositions.map(d=>[d.task+'|'+d.structure,d]));
const prioritySet=new Set(DATA.priorities.highPriorityCells.map(d=>d.task+'|'+d.structure));
function filtered(){const lane=$('lane').value,cat=$('category').value,q=$('search').value.trim().toLowerCase();return DATA.cases.filter(c=>(lane==='all'||c.lane===lane)&&(cat==='all'||c.category===cat)&&(!q||c.identity.toLowerCase().includes(q)||c.task.includes(q)||c.structures.some(s=>s.includes(q))||c.oracles.some(o=>o.includes(q))))}
function countUnique(cases,key){return new Set(cases.flatMap(c=>c[key])).size}
function renderStats(cases){const high=DATA.priorities.highPriorityCells.length;const coveredHigh=DATA.priorities.highPriorityCells.filter(p=>cases.some(c=>c.task===p.task&&c.structures.includes(p.structure))).length;const values=[['Cases',cases.length],['Synthetic',cases.filter(c=>c.lane==='synthetic').length],['Real documents',cases.filter(c=>c.lane==='superdoc').length],['Structures',countUnique(cases,'structures')],['High-priority cells',coveredHigh+' / '+high],['Visual-render eligible',cases.filter(c=>c.visualEligible).length]];$('stats').innerHTML=values.map(([k,v])=>'<div class="stat"><strong>'+v+'</strong><span>'+k+'</span></div>').join('')}
function cellCases(cases,t,s){return cases.filter(c=>c.task===t&&c.structures.includes(s))}
function renderMatrix(cases){const grid=$('matrix');grid.style.gridTemplateColumns='150px repeat('+DATA.structures.length+',minmax(50px,1fr))';let html='<div></div>'+DATA.structures.map(s=>'<div class="label">'+label(s)+'</div>').join('');for(const t of DATA.tasks){html+='<div class="label row-label">'+label(t)+'</div>';for(const s of DATA.structures){const matches=cellCases(cases,t,s),key=t+'|'+s,priority=prioritySet.has(key),plan=dispositionMap.get(key);let cls=matches.length?'tested':priority?(plan?'planned':'missing'):'empty';html+='<button class="cell '+cls+(selected===key?' selected':'')+'" data-task="'+t+'" data-structure="'+s+'" aria-label="'+label(t)+' with '+label(s)+': '+(matches.length?matches.length+' cases':plan?'planned':'not covered')+'">'+(matches.length|| (plan?'P':priority?'!':'·'))+'</button>'}}grid.innerHTML=html;grid.querySelectorAll('button').forEach(b=>b.addEventListener('click',()=>{selected=b.dataset.task+'|'+b.dataset.structure;renderMatrix(cases);renderDetail(cases,b.dataset.task,b.dataset.structure);const preview=cellCases(cases,b.dataset.task,b.dataset.structure).find(c=>c.docxVariants?.tracked);if(preview)selectDocx(preview.identity)}))}
function renderDetail(cases,t,s){const matches=cellCases(cases,t,s),plan=dispositionMap.get(t+'|'+s);let html='<h2>Cell detail</h2><h3>'+label(t)+' × '+label(s)+'</h3>';if(matches.length){html+='<span class="badge">'+matches.length+' case'+(matches.length===1?'':'s')+'</span><ul>'+matches.map(c=>'<li><code>'+c.identity+'</code> <span class="badge '+(c.lane==='synthetic'?'syn':'real')+'">'+(c.lane==='synthetic'?'synthetic':'real doc')+'</span></li>').join('')+'</ul>'}else if(plan){html+='<span class="badge">Planned</span><p>'+plan.reason+'</p><p><strong>Dependency:</strong> '+plan.dependency+'</p>'}else{html+='<div class="empty-state">No case is declared for this combination'+(prioritySet.has(t+'|'+s)?', and it is a high-priority gap.':'.')+'</div>'}$('detail').innerHTML=html}
function renderBars(cases){const counts=DATA.structures.map(s=>{const set=cases.filter(c=>c.structures.includes(s));return{s,syn:set.filter(c=>c.lane==='synthetic').length,real:set.filter(c=>c.lane==='superdoc').length,total:set.length}});const max=Math.max(1,...counts.map(x=>x.total));$('structure-bars').innerHTML=counts.map(x=>'<div class="bar-row"><span>'+label(x.s)+'</span><div class="track" aria-label="'+x.total+' cases"><span class="seg-syn" style="width:'+(x.syn/max*100)+'%"></span><span class="seg-real" style="width:'+(x.real/max*100)+'%"></span></div><span class="bar-value">'+x.total+'</span></div>').join('')}
function renderOracles(cases){const max=Math.max(1,cases.length);$('oracles').innerHTML=DATA.oracles.map(o=>{const n=cases.filter(c=>c.oracles.includes(o)).length;return'<div class="oracle"><div><div>'+label(o)+'</div><div class="track"><span class="seg-syn" style="width:'+(n/max*100)+'%"></span></div></div><strong>'+n+'</strong></div>'}).join('')}
function renderGaps(){const gaps=DATA.priorities.emptyCellDispositions;$('gaps').innerHTML=gaps.map(g=>'<div class="gap"><strong>'+label(g.task)+' × '+label(g.structure)+'</strong><p>'+g.reason+'</p><span class="badge">'+g.dependency+'</span></div>').join('')}
function renderCases(cases){$('case-caption').textContent=cases.length+' cases match the current filters.';$('case-list').innerHTML=cases.length?cases.map(c=>'<div class="case-row"><code>'+c.identity+'</code><span class="badge '+(c.lane==='synthetic'?'syn':'real')+'">'+(c.lane==='synthetic'?'synthetic':'real doc')+'</span><span>'+label(c.task)+'</span>'+(c.docxVariants?.tracked?'<button class="action view-case" type="button" data-identity="'+c.identity+'">Compare</button>':'')+'</div>').join(''):'<div class="empty-state">No cases match.</div>';$('case-list').querySelectorAll('.view-case').forEach(button=>button.addEventListener('click',()=>{selectDocx(button.dataset.identity);$('docx-comparison').scrollIntoView({behavior:'smooth',block:'start'})}))}
const VIEW_LABELS={source:'Source',tracked:'Tracked changes',accepted:'Accepted',rejected:'Rejected'};
const escapeHtml=value=>String(value??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
function decodeDocx(value){const raw=atob(value),bytes=new Uint8Array(raw.length);for(let i=0;i<raw.length;i++)bytes[i]=raw.charCodeAt(i);return bytes}
function activeDocx(){return DATA.cases.find(c=>c.identity===$('docx-case').value)}
function viewOptions(selected){return Object.entries(VIEW_LABELS).map(([value,text])=>'<option value="'+value+'"'+(value===selected?' selected':'')+'>'+text+'</option>').join('')}
function renderDocxMeta(item){const r=item.revisions;$('docx-meta').innerHTML='<span class="badge '+(item.lane==='superdoc'?'real':'syn')+'">'+(item.lane==='superdoc'?'Reviewed real document':'Synthetic fixture')+'</span><span class="badge">'+label(item.category)+'</span><span class="badge">'+label(item.task)+'</span>'+item.structures.map(s=>'<span class="badge">'+label(s)+'</span>').join('')+'<span class="badge">'+r.insertions+' insertion revision'+(r.insertions===1?'':'s')+'</span><span class="badge">'+r.deletions+' deletion revision'+(r.deletions===1?'':'s')+'</span>'+(r.formatting?'<span class="badge">'+r.formatting+' formatting revision'+(r.formatting===1?'':'s')+'</span>':'');$('expected-before').textContent=item.expectations?.rejected||item.expectations?.source||'';$('expected-after').textContent=item.expectations?.accepted||''}
async function renderPane(side,item){const state=$(side+'-view').value,target=$('docx-'+side),payload=item.docxVariants?.[state];target.innerHTML='';if(!payload){target.innerHTML='<div class="empty-state">This document state is unavailable.</div>';return}const base={experimental:true,renderHeaders:true,renderFooters:true,renderFootnotes:true,renderEndnotes:true,renderComments:true,ignoreWidth:true,ignoreHeight:true,breakPages:false,useBase64URL:true,renderChanges:state==='tracked'};await window.docx.renderAsync(decodeDocx(payload),target,null,base)}
async function renderDocxComparison(){const item=activeDocx();if(!item?.docxVariants?.tracked){$('docx-status').textContent='No embedded DOCX is available.';return}if(!window.docx?.renderAsync){$('docx-status').textContent='docxjs failed to load.';return}$('docx-status').textContent='Rendering both documents…';renderDocxMeta(item);try{await Promise.all([renderPane('left',item),renderPane('right',item)]);$('docx-status').textContent='Ready · docx-preview 0.4.0';}catch(error){$('docx-status').textContent='Render error: '+error.message}}
function setComparison(left,right){$('left-view').value=left;$('right-view').value=right;renderDocxComparison()}
function selectDocx(identity){if(!$('docx-case').querySelector('option[value="'+identity+'"]'))return;$('docx-case').value=identity;renderDocxComparison()}
function downloadView(side){const item=activeDocx(),state=$(side+'-view').value,payload=item?.docxVariants?.[state];if(!payload)return;const blob=new Blob([decodeDocx(payload)],{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=item.identity.replace(/^(synthetic|superdoc):/,'')+'.'+state+'.docx';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000)}
function initDocx(){const available=DATA.cases.filter(c=>c.docxVariants?.tracked),options=lane=>available.filter(c=>c.lane===lane).map(c=>'<option value="'+c.identity+'">'+escapeHtml(c.displayName||c.identity)+'</option>').join('');$('docx-case').innerHTML='<optgroup label="Reviewed real documents">'+options('superdoc')+'</optgroup><optgroup label="Synthetic fixtures">'+options('synthetic')+'</optgroup>';$('left-view').innerHTML=viewOptions('source');$('right-view').innerHTML=viewOptions('tracked');const preferred=available.find(c=>c.lane==='superdoc')||available.find(c=>c.identity==='synthetic:administrative-tab-aligned-status')||available[0];if(preferred)$('docx-case').value=preferred.identity;$('docx-case').addEventListener('change',renderDocxComparison);$('left-view').addEventListener('change',renderDocxComparison);$('right-view').addEventListener('change',renderDocxComparison);document.querySelectorAll('[data-preset]').forEach(button=>button.addEventListener('click',()=>setComparison(...button.dataset.preset.split(','))));$('left-download').addEventListener('click',()=>downloadView('left'));$('right-download').addEventListener('click',()=>downloadView('right'));let syncing=false;for(const [from,to] of [[$('docx-left'),$('docx-right')],[$('docx-right'),$('docx-left')]])from.addEventListener('scroll',()=>{if(!$('sync-scroll').checked||syncing)return;syncing=true;const maxFrom=from.scrollHeight-from.clientHeight,maxTo=to.scrollHeight-to.clientHeight;to.scrollTop=maxFrom>0?from.scrollTop/maxFrom*maxTo:0;to.scrollLeft=from.scrollLeft;requestAnimationFrame(()=>{syncing=false})});renderDocxComparison()}
function render(){const cases=filtered();renderStats(cases);renderMatrix(cases);renderBars(cases);renderOracles(cases);renderCases(cases)}
function setSidebarHidden(hidden){document.body.classList.toggle('sidebar-hidden',hidden);const toggle=$('sidebar-toggle');toggle.textContent=hidden?'Show sidebar':'Hide sidebar';toggle.setAttribute('aria-expanded',String(!hidden));try{localStorage.setItem('docx-dashboard-sidebar-hidden',hidden?'1':'0')}catch{}}
function initSidebar(){let hidden=false;try{hidden=localStorage.getItem('docx-dashboard-sidebar-hidden')==='1'}catch{}setSidebarHidden(hidden);$('sidebar-toggle').addEventListener('click',()=>setSidebarHidden(!document.body.classList.contains('sidebar-hidden')))}
$('lane').addEventListener('change',render);$('category').addEventListener('change',render);$('search').addEventListener('input',render);$('stamp').textContent='Generated '+new Date(DATA.generatedAt).toLocaleString();renderGaps();render();initSidebar();initDocx();
</script>
</body>
</html>\n`;
}

const isCli = process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (isCli) {
    const { outputPath, fixturesDir, corpusFixturesDir } = parseArgs(process.argv.slice(2));
    const libraries = {
        jszipSource: readFileSync(join(process.cwd(), 'node_modules', 'jszip', 'dist', 'jszip.min.js'), 'utf8'),
        docxPreviewSource: readFileSync(join(process.cwd(), 'node_modules', 'docx-preview', 'dist', 'docx-preview.min.js'), 'utf8')
    };
    const html = renderDashboardHtml(buildDashboardData(fixturesDir, corpusFixturesDir), libraries);
    mkdirSync(dirname(outputPath), { recursive: true });
    writeFileSync(outputPath, html, 'utf8');
    console.log(`Wrote test comparison dashboard: ${outputPath}`);
}
