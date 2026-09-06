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

