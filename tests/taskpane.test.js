const { test } = require("node:test");
const assert = require("node:assert/strict");
const vm = require("node:vm");
const fs = require("node:fs");
const path = require("node:path");
const tasks = require("../public/ai-tasks");

function setup(existingStorage = new Map()) {
  const nodes = new Map(), storage = existingStorage;
  const node = id => {
    if (!nodes.has(id)) nodes.set(id, { value: "", textContent: "", dataset: {}, classList: { toggle() {} }, setAttribute() {}, addEventListener() {}, replaceChildren() {}, add() {} });
    return nodes.get(id);
  };
  const context = vm.createContext({ document: { getElementById: node }, Office: { onReady() {} }, window: { localStorage: { getItem: k => storage.get(k), setItem: (k, v) => storage.set(k, v), removeItem: k => storage.delete(k) } }, console, AbortController, Map, Set, Intl, setTimeout, clearTimeout, crypto: { randomUUID: () => "test" }, PictosAiTasks: tasks });
  const run = source => vm.runInContext(source, context);
  run(fs.readFileSync(path.join(__dirname, "../public/taskpane.js"), "utf8"));
  run('requestRender = () => {}; syncAiToggle = () => {}; syncAiSettingsStatus = () => {};');
  return { run, context, storage, node };
}
test("provider readiness and cache isolate API, Codex and models", () => {
  const { run } = setup();
  assert.equal(run('aiApiKey="saved-api-key"; isAiProviderReady()'), true);
  const key = run('buildAiSearchCacheKey("ambition")');
  assert.equal(run('aiProvider="codex"; isAiProviderReady()'), false);
  run('codexReady=true; codexClient={}; codexModel="codex-model"');
  assert.equal(run('isAiProviderReady()'), true);
  assert.notEqual(run('buildAiSearchCacheKey("ambition")'), key);
  assert.equal(run('aiProvider="api"; aiApiKey'), "saved-api-key");
});
test("provider switch aborts in-flight work and prevents stale results", async () => {
  const { run, context } = setup();
  let resolve;
  context.deferred = new Promise(r => { resolve = r; });
  run('aiProvider="codex"; aiEnabled=true; codexReady=true; codexModel="test"; codexClient={search:()=>deferred}; allLogos=[{id:1,name:"A.svg",keywords:[]},{id:2,name:"B.svg",keywords:[]}]; logoById=new Map(allLogos.map(x=>[x.id,x]));');
  const pending = run('requestAiSearch("ambition")');
  const signal = run('aiRequestController.signal');
  run('clearAiSearchState(); aiProvider="api";');
  assert.equal(signal.aborted, true);
  resolve({ parsed: { ordered_ids: [2, 1], note: "" } });
  await pending;
  assert.equal(run('aiSearchState.resultIds.length'), 0);
  assert.equal(run('aiSearchCache.size'), 0);
});
test("unknown candidate IDs never leak into displayed ranking", () => {
  const { run } = setup();
  const ids = run('logoById=new Map([[1,{id:1}],[2,{id:2}],[3,{id:3}]]);completeAiResultIds([3,999,1,1],[{id:1},{id:2}])');
  assert.deepEqual(Array.from(ids), [1, 2]);
});

test("identical concurrent AI searches share one provider call and completed result", async () => {
  const { run, context } = setup();
  let finish, calls = 0;
  context.response = new Promise(resolve => { finish = resolve; });
  context.call = () => { calls++; return context.response; };
  run('aiEnabled=true;aiProvider="codex";codexReady=true;codexClient={search:call};codexModel="test";allLogos=[{id:1,name:"Rocket",keywords:[]},{id:2,name:"Mountain",keywords:[]}];logoById=new Map(allLogos.map(l=>[l.id,l]));');
  const first = run('requestAiSearch("ambition")');
  const second = run('requestAiSearch("ambition")');
  assert.equal(first, second);
  finish({ parsed: { ordered_ids: [1], note: "" } });
  await Promise.all([first, second]);
  await run('requestAiSearch("ambition")');
  assert.equal(calls, 1);
});

test("AI debounce resets to two seconds and Enter flush cancels the timer", () => {
  const { run, context, node } = setup();
  let current = 0, searches = 0;
  const timers = new Map();
  context.setTimeout = (fn, delay) => { timers.set(++current, { fn, delay }); return current; };
  context.clearTimeout = id => timers.delete(id);
  context.search = () => { searches++; };
  run('runSearchCycle=search;aiEnabled=true;aiApiKey="test";');
  node('search-input').value = 'Supplier';
  run('scheduleSearch()');
  assert.equal([...timers.values()][0].delay, 2000);
  node('search-input').value = 'Supplier Questionnaire';
  run('scheduleSearch()');
  assert.equal(timers.size, 1);
  assert.equal(searches, 0);
  run('scheduleSearch({immediate:true})');
  assert.equal(timers.size, 0);
  assert.equal(searches, 1);
});

test("search scores cached for one query are independent of subsequent queries", () => {
  const { run, node } = setup();
  run('allLogos=attachKeywords([{id:1,name:"rocket growth.svg"},{id:2,name:"growth.svg"}],new Map());buildSearchIndex(allLogos)');
  node('search-input').value = 'rocket';
  const firstScore = run('filterLogos()[0].relevanceScore');
  node('search-input').value = '';
  run('filterLogos()');
  node('search-input').value = 'rocket';
  assert.equal(run('filterLogos()[0].relevanceScore'), firstScore);
  assert.ok(firstScore > 0);
});

test("late SVG loads cannot repopulate a replaced library", async () => {
  const { run, context } = setup();
  let finish;
  context.pendingSvg = new Promise(resolve => { finish = resolve; });
  run('fetchSvgTextFromZip=()=>pendingSvg');
  const pending = run('getSvgText({id:1,name:"one.svg"})');
  run('libraryGeneration++;previewCache.clear();previewCacheBytes=0');
  finish('<svg/>');
  await assert.rejects(pending, /Bibliothèque remplacée/);
  assert.equal(run('previewCache.size'), 0);
  assert.equal(run('previewCacheBytes'), 0);
});

function shortcutSetup() {
  const result = setup();
  const { run, context } = result;
  context.Office.addin = { showAsTaskpane: async () => {} };
  context.Office.context = { requirements: { isSetSupported: () => true } };
  context.slide = 'one'; context.insertions = 0; context.deleted = 0;
  run('getSelectedText=async()=>"Supplier Questionnaire";getSelectedSlideId=async()=>slide;setSettingsPanelOpen=()=>{};getPreparedSvg=async()=>"<svg/>";goToSlide=async()=>{};getNextInsertPosition=async()=>({});insertSvg=async()=>{insertions++};recordRecent=()=>{};runSearchCycle=async()=>{};getRenderableLogos=()=>[{id:1,name:"Checklist"}];allLogos=[{id:1}];replaceSelectionEnabled=true;');
  return result;
}

test("shortcut shows selected text and adds one logo without replacement, duplicate shortcut ignored", async () => {
  const { run, context, node } = shortcutSetup();
  let release;
  context.prepared = new Promise(resolve => { release = resolve; });
  run('getPreparedSvg=()=>prepared');
  const action = run('searchSelectedPictogram()');
  const second = run('searchSelectedPictogram()');
  release('<svg/>');
  await Promise.all([action, second]);
  assert.equal(context.insertions, 1);
  assert.equal(node('search-input').value, 'Supplier Questionnaire');
});

test("shortcut refuses to insert if slide or search changed during AI/preparation", async () => {
  const { run, context, node } = shortcutSetup();
  run('getPreparedSvg=async()=>{slide="two";return "<svg/>"}');
  await run('searchSelectedPictogram()');
  assert.equal(context.insertions, 0);
  assert.match(node('status').textContent, /changé/);
});

test("shared runtime alone never claims that PowerPoint supports keyboard shortcuts", async () => {
  const { run, context, node } = setup();
  let queried = false;
  context.Office.context = { requirements: { isSetSupported: name => name === 'SharedRuntime' }, diagnostics: { version: '16.0.17425.20146' } };
  context.Office.actions = { associate() {}, getShortcuts: async () => { queried = true; return { SearchSelectedPictogram: 'Ctrl+Alt+P' }; } };
  await run('refreshShortcutStatus()');
  assert.equal(node('shortcut-status').dataset.state, 'unsupported');
  assert.match(node('shortcut-status').textContent, /16\.0\.17425\.20146/);
  assert.equal(queried, false);
  assert.equal(node('shortcut-refresh').disabled, false);
});

test("shortcut diagnosis distinguishes missing manifest, conflict, actual binding and Office failure", async () => {
  const { run, context, node } = setup();
  context.Office.context = { requirements: { isSetSupported: () => true } };
  context.Office.actions = { associate() {}, getShortcuts: async () => ({}) };
  await run('refreshShortcutStatus()');
  assert.equal(node('shortcut-status').dataset.state, 'missing');
  context.Office.actions.getShortcuts = async () => ({ SearchSelectedPictogram: null });
  await run('refreshShortcutStatus()');
  assert.equal(node('shortcut-status').dataset.state, 'conflict');
  context.Office.actions.getShortcuts = async () => ({ SearchSelectedPictogram: 'Ctrl+Shift+P' });
  await run('refreshShortcutStatus()');
  assert.equal(node('shortcut-status').dataset.state, 'registered');
  assert.equal(node('shortcut-key').textContent, 'Ctrl + Shift + P');
  context.Office.actions.getShortcuts = async () => { throw new Error('Runtime unavailable'); };
  await run('refreshShortcutStatus()');
  assert.equal(node('shortcut-status').dataset.state, 'error');
  assert.match(node('shortcut-status').textContent, /Runtime unavailable/);
});

test("action registration retries when Office was unavailable and restoration only changes this action", async () => {
  const { run, context, node } = setup();
  context.Office.context = { platform: 'Mac', requirements: { isSetSupported: () => true } };
  let attempts = 0, mapping;
  context.Office.actions = {
    associate: () => { if (++attempts === 1) throw new Error('Not ready'); },
    getShortcuts: async () => ({ SearchSelectedPictogram: mapping?.SearchSelectedPictogram || 'Cmd+Shift+P' }),
    replaceShortcuts: async value => { mapping = JSON.parse(JSON.stringify(value)); }
  };
  run('registerSelectionAction()');
  assert.equal(run('selectionActionRegistered'), false);
  await run('refreshShortcutStatus()');
  assert.equal(attempts, 2);
  assert.equal(node('shortcut-key').textContent, 'Cmd + Shift + P');
  assert.equal(mapping, undefined);
  await run('restoreSelectionShortcut()');
  assert.deepEqual(mapping, { SearchSelectedPictogram: 'Cmd+Alt+P' });
  assert.equal(node('shortcut-key').textContent, 'Cmd + Alt + P');
  assert.equal(attempts, 2);
});

test("selection button works without shared runtime and always releases its busy state", async () => {
  const { run, context, node } = shortcutSetup();
  context.Office.addin = undefined;
  context.Office.context.requirements.isSetSupported = name => name !== 'SharedRuntime';
  await run('searchSelectedPictogram()');
  assert.equal(context.insertions, 1);
  assert.equal(node('selection-insert').disabled, false);
  context.Office.context.requirements.isSetSupported = () => true;
  context.Office.addin = { showAsTaskpane: async () => { throw new Error('Not running in a shared runtime'); } };
  await run('searchSelectedPictogram(undefined, { revealPane: false })');
  assert.equal(context.insertions, 2);
  run('getSelectedText=async()=>{throw new Error("Selection unavailable")}');
  let completions = 0;
  context.actionEvent = { completed: () => completions++ };
  await run('searchSelectedPictogram(actionEvent)');
  assert.equal(completions, 1);
  assert.equal(context.insertions, 2);
  assert.equal(node('selection-insert').disabled, false);
  assert.match(node('status').textContent, /Selection unavailable/);
});

test("captured text remains visible when insertion is blocked by a missing library", async () => {
  const { run, context, node } = shortcutSetup();
  run('allLogos=[]');
  await run('searchSelectedPictogram()');
  assert.equal(node('search-input').value, 'Supplier Questionnaire');
  assert.equal(context.insertions, 0);
  assert.match(node('status').textContent, /bibliothèque ZIP/);
});

test("ZIP import parses nested SVGs and distinguishes same basenames", async () => {
  const { run, context } = setup();
  const JSZip = require('../public/vendor/jszip-3.10.1.min.js');
  const zip = new JSZip();
  zip.file('One/icon.svg','<svg/>'); zip.file('Two/icon.svg','<svg/>'); zip.file('ignore.txt','not an icon');
  context.JSZip = JSZip;
  context.archive = await zip.generateAsync({type:'nodebuffer'});
  const parsed = await run('testSession=createMainZipSession();testSession.load(archive)');
  assert.equal(parsed.items.length, 2);
  assert.notEqual(parsed.items[0].displayName, parsed.items[1].displayName);
  assert.equal((await run('testSession.getSvg("Two/icon.svg")')).svgText, '<svg/>');
});

test("failed candidate ZIP preserves the active library", async () => {
  const { run, context, node } = setup();
  context.file = {name:'broken.zip',arrayBuffer:async()=>new ArrayBuffer(8)};
  run('allLogos=[{id:1,name:"Existing.svg"}];zipSession={active:true};createZipSession=()=>({load:async()=>{throw Error("invalid ZIP")},terminate:()=>{}})');
  await run('handleZipFile(file)');
  assert.equal(run('allLogos[0].name'), 'Existing.svg');
  assert.equal(run('zipSession.active'), true);
  assert.equal(run('libraryBusy'), false);
  assert.equal(node('refresh-btn').disabled, false);
});

test("preview cache evicts old off-screen assets and revokes their URLs", () => {
  const { run, context, node } = setup();
  const revoked = [];
  context.URL = {revokeObjectURL:url=>revoked.push(url)};
  node('logo-grid').children = [{dataset:{logoId:'0'}}];
  run('for(let i=0;i<400;i++){previewCache.set(i,{bytes:100,url:`blob:${i}`});localObjectUrls.add(`blob:${i}`);previewCacheBytes+=100}trimPreviewCache()');
  assert.equal(run('previewCache.size'),256);
  assert.equal(run('previewCacheBytes'),25600);
  assert.equal(run('previewCache.has(0)'),true);
  assert.equal(revoked.length,144);
});

test("large-library AI shortlist favors matching icons and prepares phrases once", () => {
  const { run } = setup();
  run('allLogos=attachKeywords(Array.from({length:1740},(_,id)=>({id,name:id%2?`Rocket ${id}.svg`:`Dog ${id}.svg`})),new Map());let normalizations=0;const originalNormalize=normalizeSearchText;normalizeSearchText=value=>{normalizations++;return originalNormalize(value)};');
  const result = run('collectAiCandidates("rocket",{coreConcepts:["rocket"],concreteObjects:["rocket"],visualMetaphors:[],relatedKeywords:[]})');
  assert.equal(result.length,72);
  assert.ok(result.every(logo => logo.name.startsWith('Rocket')));
  assert.equal(run('normalizations'),3);
});

test("virtual grid follows its own scroll viewport and skips unchanged row layouts", () => {
  const { run, context, node } = setup();
  const grid = node('logo-grid'), scroller = node('library-scroll');
  let writes = 0;
  const style = () => new Proxy({}, {set(target, key, value) { writes++; target[key] = value; return true; }});
  grid.getBoundingClientRect = () => ({top:-7536, width:284});
  grid.style = style(); grid.children = []; grid.appendChild = card => grid.children.push(card);
  scroller.getBoundingClientRect = () => ({top:64}); scroller.clientHeight = 586;
  context.window.innerHeight = 10000;
  context.PictosGrid = require('../public/grid-window');
  context.makeCard = (logo, index) => ({dataset:{index:String(index)}, style:style()});
  run('displayedLogos=Array.from({length:1740},(_,id)=>({id}));createLogoCard=makeCard;trimPreviewCache=()=>{};renderGridWindow()');
  const indexes = grid.children.map(card => Number(card.dataset.index));
  assert.ok(indexes[0] > 200);
  assert.ok(indexes.includes(230));
  assert.ok(indexes.length < 60);
  const previousWrites = writes;
  run('renderGridWindow()');
  assert.equal(writes, previousWrites);
});

test("preview loading observes the library viewport", () => {
  const { run, context, node } = setup();
  let options;
  context.IntersectionObserver = class { constructor(callback, config) { options = config; } };
  run('getLazyObserver()');
  assert.equal(options.root, node('library-scroll'));
});

function connectionSetup(storage) {
  const result=setup(storage), {run,context,node}=result;
  context.setTimeout=()=>1;context.clearTimeout=()=>{};context.Option=class {constructor(name,id){this.text=name;this.value=id;}};
  context.reply={connected:true,models:[{id:'luna'}],limits:null}; context.failure=null; context.reads=0; context.searches=0;
  run('PictosAiProviders={DEFAULT_URL:"http://127.0.0.1:43129",formatLimits:()=>"",CodexClient:class {constructor({url}){this.url=url;} async status(){reads++;if(failure)throw failure;return reply;} async connection(){reads++;if(failure)throw failure;return {connected:reply.connected};} search(){searches++;}}};aiProvider="codex";aiEnabled=true;');
  node('ai-codex-url').value='http://127.0.0.1:43129';node('ai-codex-token').value='a'.repeat(64);
  return result;
}

test('a verified pairing survives a new PowerPoint session', async () => {
  const first=connectionSetup();await first.run('checkCodexConnection({full:true})');
  assert.equal(first.storage.get('logosPptCodexToken'),'a'.repeat(64));
  const reopened=connectionSetup(first.storage);reopened.node('ai-codex-token').value='';
  reopened.run('restorePreferences()');
  assert.equal(reopened.node('ai-codex-token').value,'a'.repeat(64));
});

test('status checks preserve pending searches, deduplicate reads and recover after a restart', async () => {
  const {run,context,storage}=connectionSetup();
  run('aiRequestController=new AbortController()'); const signal=run('aiRequestController.signal');
  const first=run('checkCodexConnection({full:true})'),second=run('checkCodexConnection({full:true})');
  assert.equal(first,second);await first;
  assert.equal(context.reads,1);assert.equal(signal.aborted,false);
  context.failure=Object.assign(new Error('Compagnon arrêté'),{code:'COMPANION_OFFLINE'});
  await run('checkCodexConnection({silent:true})');assert.equal(run('codexConnectionState'),'offline');
  assert.equal(storage.get('logosPptCodexToken'),'a'.repeat(64));
  context.failure=null;await run('checkCodexConnection({silent:true})');
  assert.equal(run('codexConnectionState'),'ready');assert.equal(run('codexReady'),true);
  assert.equal(context.searches,0);assert.equal(signal.aborted,false);
});

test('a wrong pairing code is distinguished from an account awaiting login', async () => {
  const {run,context}=connectionSetup();
  context.failure=Object.assign(new Error('Code incorrect'),{code:'PAIRING_REQUIRED'});
  await run('checkCodexConnection({full:true})');assert.equal(run('codexConnectionState'),'unpaired');
  context.failure=null;context.reply={connected:false,models:[]};
  await run('checkCodexConnection({full:true})');assert.equal(run('codexConnectionState'),'login');assert.equal(run('codexReady'),false);
});

test('blocked Office storage is reported without disabling the current connection', async () => {
  const {run,context}=connectionSetup();
  context.window.localStorage.setItem=()=>{throw new Error('Storage blocked');};
  await run('checkCodexConnection({full:true})');
  assert.equal(run('codexReady'),true);
  assert.match(run('codexConnectionMessage'),/pas pu mémoriser/);
});
