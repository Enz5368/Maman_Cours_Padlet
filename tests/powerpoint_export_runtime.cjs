const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const {test} = require('node:test');
const source = fs.readFileSync(require('node:path').join(__dirname, '../assets/app.js'), 'utf8');
function extract(name) {
  const start = source.search(new RegExp(`^      (?:async )?function ${name}\\(`, 'm'));
  assert.ok(start >= 0);
  const rest = source.slice(start);
  const next = rest.slice(1).search(/^      (?:async )?function /m);
  return next < 0 ? rest : rest.slice(0, next + 1);
}
test('PowerPoint image fills the slide without an added title', () => {
  const ctx = vm.createContext({});
  for (const name of ['pptxSlide', 'pptxMediaShape', 'pptxPresentation']) vm.runInContext(extract(name), ctx);
  const xml = ctx.pptxSlide({title:'Cours'}, {fullBleed:true,elements:[{kind:'image',x:0,y:0,w:960,h:540}]}, 0, [{elementIndex:0,kind:'image',mediaRelId:'rId2'}]);
  assert.ok(xml.includes('<a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/>'));
  assert.ok(ctx.pptxPresentation(1).includes('<p:sldSz cx="12192000" cy="6858000"'));
  assert.equal((xml.match(/<p:pic>/g)||[]).length,1);
  assert.ok(!xml.includes('<p:sp>'));
});
test('Export preserves selected slide order and releases its lock', async () => {
  let captured, clicked = false, unlocked = false;
  const pages = [{number:2},{number:4}];
  const ctx = vm.createContext({
    findLessonContext: () => ({lesson:{title:'Cours'}}),
    beginSaveLock: () => () => {unlocked=true;},
    document:{fonts:{ready:Promise.resolve()},getElementById:()=>({querySelectorAll:selector=>{
      assert.ok(selector.includes(':not([data-word-export="false"])')); return pages;
    }}),createElement:()=>({click(){clicked=true;}})},
    rasterizePreviewPage:async page=>new Uint8Array([page.number]),
    slideSize:{width:960,height:540},
    makePptx:async (activity,media)=>{captured={activity,media};return new Uint8Array();},
    Blob, URL:{createObjectURL:()=> 'blob:test',revokeObjectURL(){}},
    slugify:()=> 'cours',setTimeout:fn=>fn(),toast:message=>{assert.ok(!message.includes('impossible'));}
  });
  vm.runInContext(extract('exportPreviewPowerPoint'),ctx);
  await ctx.exportPreviewPowerPoint('lessonPrintPreview','lesson','lesson');
  assert.deepEqual(Array.from(captured.media, items=>items[0].bytes[0]),[2,4]);
  assert.equal(captured.activity.slides.length,2);
  assert.ok(clicked && unlocked);
});

test('Word preview groups half slides on portrait A4 pages and full slides on landscape pages', () => {
  const element = tag => ({tag, children: [], dataset: {}, appendChild(child) { this.children.push(child); }, setAttribute() {}});
  const slides = ['half', 'half', 'landscape', 'half'].map(layout => ({dataset: {wordLayout: layout, wordExport: 'true'}}));
  const preview = element('article');
  preview.querySelectorAll = () => slides;
  preview.replaceChildren = () => { preview.children = []; };
  const ctx = vm.createContext({document: {getElementById: () => preview, createElement: element}});
  vm.runInContext(extract('arrangeWordPreviewPages'), ctx);
  ctx.arrangeWordPreviewPages('lessonPrintPreview');
  assert.deepEqual(Array.from(preview.children, page => page.className), [
    'word-preview-sheet portrait',
    'word-preview-sheet landscape',
    'word-preview-sheet portrait',
  ]);
  assert.equal(preview.children[0].children.length, 3);
  assert.equal(preview.children[1].children.length, 2);
  assert.equal(preview.children[2].children.length, 2);
  assert.deepEqual(slides.map(slide => slide.dataset.wordOrder), ['0', '1', '2', '3']);
});

test('Landscape Word export keeps the original 16:9 ratio', () => {
  const ctx = vm.createContext({slideSize: {width: 960, height: 540}});
  vm.runInContext(extract('docxMixedPageParagraph'), ctx);
  const xml = ctx.docxMixedPageParagraph('rId1', 'landscape', 1, false);
  const extent = /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(xml);
  assert.ok(extent);
  assert.ok(Math.abs(Number(extent[1]) / Number(extent[2]) - 16 / 9) < 0.0001);
});

test('Generated PowerPoint uses a valid slide layout identifier', () => {
  const ctx = vm.createContext({});
  vm.runInContext(extract('pptxSlideMaster'), ctx);
  assert.ok(ctx.pptxSlideMaster().includes('<p:sldLayoutId id="2147483649" r:id="rId1"/>'));
});

