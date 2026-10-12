const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const {JSDOM} = require('jsdom');
const assets = path.join(__dirname, '../src/gui/assets');
const html = fs.readFileSync(path.join(assets, 'index.html'), 'utf8');
const script = fs.readFileSync(path.join(assets, 'app.js'), 'utf8');

async function setup(t, overrides = {}) {
  const dom = new JSDOM(html, {runScripts:'outside-only'});
  t.after(() => dom.window.close());
  const window = dom.window;
  const defaults = {template:'',format_spec:'',ai_base_url:'',ai_model:'',ai_key_env:'M2W_AI_API_KEY',ai_json_mode:false,python:'python.exe',pandoc:'pandoc.exe',convert_formulas:true,
    add_numbers:true,number_format:'Simple',create_visio:false,auto_start_word:true,overwrite:false,
    preserve_table_style:false,auto_captions:true,reference_superscript:false,stage_timeout:1800,dark_mode:false};
  const state = {state:'idle',status:'准备好',progress:0,error:'',warnings:[],output:'',busy:false,
    close_requested:false,generation:0,last_seq:0,logs:[],truncated:false};
  const calls = [];
  window.pywebview = {api:{bootstrap:async()=>({settings:defaults}),
    poll:async after=>({...state,logs:state.logs.filter(line=>line.seq>after)}),save_settings:async()=>({ok:true}),
    start_conversion:async data=>{calls.push(data);state.state='running';state.busy=true;return {ok:true};},
    ...overrides}};
  window.eval(script);
  const settle = () => new Promise(resolve => setTimeout(resolve, 20));
  await settle();
  return {window, state, calls, settle, $:id=>window.document.getElementById(id)};
}

test('automatic output follows source while preserving a manually selected output', async t => {
  const {$,window} = await setup(t);
  $('markdown').value='C:\\reports\\first.md';
  $('markdown').dispatchEvent(new window.Event('input'));
  assert.equal($('output').value,'C:\\reports\\first.docx');
  $('markdown').value='C:\\reports\\next.markdown';
  $('markdown').dispatchEvent(new window.Event('input'));
  assert.equal($('output').value,'C:\\reports\\next.docx');
  $('output').value='D:\\custom.docx';
  $('markdown').value='third.md';
  $('markdown').dispatchEvent(new window.Event('input'));
  assert.equal($('output').value,'D:\\custom.docx');
});

test('protocol ownership disables and restores legacy formatting controls', async t => {
  const {$,window} = await setup(t);
  $('format_spec').value='template/spec.json';
  $('format_spec').dispatchEvent(new window.Event('change'));
  for (const id of ['auto_captions','preserve_table_style','add_numbers','number_format']) assert.equal($(id).disabled,true);
  $('format_spec').value='';
  $('format_spec').dispatchEvent(new window.Event('change'));
  assert.equal($('auto_captions').disabled,false);
  assert.equal($('add_numbers').disabled,false);
});

test('AI chat forwards draft and ephemeral key without putting the key in settings', async t => {
  let sent, saved;
  const {$,window,settle} = await setup(t, {
    template_chat:async data=>{sent=data;return {ok:true};},
    save_settings:async data=>{saved=data;return {ok:true};}
  });
  $('ai-key').value='secret-test';
  $('ai_model').value='test-model';
  $('ai_model').dispatchEvent(new window.Event('change'));
  $('ai-send').click();await settle();
  assert.equal(sent.api_key,'secret-test');
  assert.equal(sent.config.model,'test-model');
  assert.ok(JSON.parse(sent.draft).spec);
  await new Promise(resolve=>setTimeout(resolve,450));
  assert.ok(saved);
  assert.equal(JSON.stringify(saved).includes('secret-test'),false);
});

test('failed template job retains manual edits and displays server text safely', async t => {
  let result={revision:0,draft:'',messages:[],error:'',project:''};
  const {$,window} = await setup(t,{template_snapshot:async()=>result});
  $('spec-editor').value='manual edits';
  result={revision:1,draft:'old draft',messages:['<script>bad()</script>'],error:'failed',project:''};
  await new Promise(resolve=>setTimeout(resolve,500));
  assert.equal($('spec-editor').value,'manual edits');
  assert.equal($('ai-conversation').querySelector('script'),null);
  assert.equal(window.bad,undefined);
});

test('native-mode submission forwards options and freezes controls', async t => {
  const {$,window,calls,settle} = await setup(t);
  window.document.querySelector('[data-mode=native]').click();
  $('reference_superscript').checked=true;
  $('markdown').value='test.md';$('output').value='test.docx';
  $('start').click(); await settle();
  assert.equal(calls.length,1);
  assert.equal(calls[0].convert_formulas,false);
  assert.equal(calls[0].reference_superscript,true);
  assert.equal($('number-options').hidden,true);
  assert.equal($('files-fields').disabled,true);
  assert.equal($('start').disabled,true);
  assert.equal($('cancel').hidden,false);
});

test('backend errors unfreeze controls and display untrusted strings as text', async t => {
  const attack='<img src=x onerror="window.compromised=true">';
  const {$,window,settle} = await setup(t, {start_conversion:async()=>({ok:false,error:attack})});
  $('start').click(); await settle();
  assert.equal($('files-fields').disabled,false);
  assert.equal($('toast').textContent,attack);
  assert.equal($('toast').querySelector('img'),null);
  assert.equal(window.compromised,undefined);
});

test('completion exposes result actions, warning details, and safe logs', async t => {
  const {$,state,settle} = await setup(t);
  state.state='warning';state.output='report.docx';state.warnings=['missing.png'];
  state.logs=[{seq:1,text:'<script>bad()</script>\n'}];state.last_seq=1;
  await new Promise(resolve=>setTimeout(resolve,500));await settle();
  assert.equal($('result-buttons').hidden,false);
  assert.equal($('result-warnings').textContent,'missing.png');
  assert.equal($('log-details').open,true);
  assert.equal($('log').querySelector('script'),null);
  assert.match($('log').textContent, /<script>/);
});

test('pending start request cannot be submitted twice', async t => {
  let release, starts=0;
  const {$,settle} = await setup(t, {start_conversion:async()=>{
    starts++; await new Promise(resolve=>release=resolve); return {ok:true};
  }});
  $('start').click(); $('start').click();
  assert.equal(starts,1);
  assert.equal($('files-fields').disabled,true);
  release(); await settle();
  assert.equal($('files-fields').disabled,false);
});

test('copy uses the native clipboard bridge without browser clipboard permissions', async t => {
  let copies = 0;
  const {$,settle} = await setup(t, {copy_log:async()=>{copies++;return {ok:true,source:'file'};}});
  $('copy-log').click();await settle();
  assert.equal(copies,1);
  assert.equal($('toast').textContent,'已复制完整日志。');
});

test('polling preserves selected log nodes and unchanged error details', async t => {
  const {$,state,window} = await setup(t);
  state.state='error';state.error='可复制的错误详情';
  state.logs=[{seq:1,text:'第一行日志\n'}];state.last_seq=1;
  await new Promise(resolve=>setTimeout(resolve,500));
  const original=$('log').firstChild, errorNode=$('result-error').firstChild;
  const range=window.document.createRange();range.selectNodeContents(original);
  window.getSelection().addRange(range);
  state.logs.push({seq:2,text:'第二行日志\n'});state.last_seq=2;
  await new Promise(resolve=>setTimeout(resolve,500));
  assert.equal($('log').firstChild,original);
  assert.equal($('result-error').firstChild,errorNode);
  assert.equal(window.getSelection().toString(),'第一行日志\n');
  assert.equal($('log').textContent,'第一行日志\n第二行日志\n');
});
