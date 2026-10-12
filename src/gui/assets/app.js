'use strict';
(() => {
  const $ = id => document.getElementById(id);
  function setText(id, value) {
    const node = $(id);
    if (node.textContent !== value) node.textContent = value;
  }
  const settingKeys = ['template','format_spec','ai_base_url','ai_model','ai_key_env','ai_json_mode','python','pandoc','convert_formulas','add_numbers','number_format','create_visio','auto_start_word','overwrite','preserve_table_style','auto_captions','reference_superscript','stage_timeout','dark_mode'];
  let api, ready = false, submitting = false, busy = false, mode = true, dark = false;
  let after = 0, generation = -1, autoOutput = '', toastTimer, saveTimer, logText = '', previousState = '', suppressClosePrompt = false;
  function toast(message) {
    $('toast').textContent = String(message); $('toast').hidden = false;
    clearTimeout(toastTimer); toastTimer = setTimeout(() => $('toast').hidden = true, 5500);
  }
  async function call(name, ...args) {
    if (!api) throw new Error('桌面服务尚未就绪，请稍候。');
    const result = await api[name](...args);
    if (result && result.ok === false) throw new Error(result.error || '操作未能完成。');
    return result;
  }
  function settings() {
    const data = {convert_formulas:mode, dark_mode:dark};
    for (const key of settingKeys) {
      if (key in data) continue;
      const element = $(key);
      data[key] = element.type === 'checkbox' ? element.checked : element.value;
    }
    data.markdown = $('markdown').value.trim(); data.output = $('output').value.trim();
    return data;
  }
  function updateMode() {
    document.querySelectorAll('[data-mode]').forEach(button => button.setAttribute('aria-pressed', String((button.dataset.mode === 'mathtype') === mode)));
    $('number-options').hidden = !mode;
    $('mode-note').textContent = mode ? '需要 Microsoft Word 和 MathType，转换期间请勿操作 Office。' : '可编辑的 Word 原生公式，无需启动 Office。此模式暂不提供右编号。';
    const protocol = !!$('format_spec').value.trim();
    $('number_format').disabled = busy || protocol || !$('add_numbers').checked || !mode;
    for (const id of ['auto_captions','preserve_table_style','add_numbers']) $(id).disabled = busy || protocol;
  }
  function updateSource() {
    const path = $('markdown').value.trim(), output = $('output');
    if (path && (!output.value.trim() || output.value === autoOutput)) {
      autoOutput = /\.(md|markdown)$/i.test(path) ? path.replace(/\.(md|markdown)$/i, '.docx') : path + '.docx';
      output.value = autoOutput;
    }
    $('source-name').textContent = path ? path.split(/[\\/]/).pop() : '添加 Markdown 文档';
    $('source-name').title = path;
    $('source-note').textContent = path ? '已选择 · 图片路径相对源文件解析' : '选择 .md 或 .markdown 文件';
  }
  function freeze(value) {
    busy = value;
    for (const id of ['files-fields','template-fields','option-fields','advanced-fields']) $(id).disabled = value;
    $('start').disabled = !ready || value || submitting;
    $('start').firstChild.textContent = value ? '正在处理' : '开始转换';
    updateMode();
  }
  function scheduleSave() {
    if (!ready) return;
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => call('save_settings', settings()).catch(error => toast(error.message)), 400);
  }
  function showLog() { $('log-details').open = true; }
  function render(snapshot) {
    if (snapshot.generation !== generation) {
      generation = snapshot.generation; logText = ''; $('log').textContent = '';
    }
    if (snapshot.logs.length) {
      const box = $('log'), atBottom = box.scrollTop + box.clientHeight >= box.scrollHeight - 24;
      const added = snapshot.logs.map(line => line.text).join('');
      logText += added;
      if (logText.length > 180000) {
        logText = '…较早日志已省略，请打开完整日志。\n' + logText.slice(-150000);
        box.textContent = logText;
      } else box.appendChild(document.createTextNode(added));
      if (atBottom) box.scrollTop = box.scrollHeight;
    }
    if (snapshot.truncated) $('log-caption').textContent = '部分早期日志已省略 · 完整日志另存';
    else $('log-caption').textContent = '详细过程与诊断信息';
    after = snapshot.last_seq;
    freeze(snapshot.busy || submitting);
    $('cancel').hidden = !['running','cancelling'].includes(snapshot.state);
    $('cancel-template').hidden = snapshot.state !== 'templating';
    $('cancel').disabled = snapshot.state === 'cancelling';
    $('status-text').textContent = snapshot.status;
    $('status-dot').dataset.state = snapshot.state;
    $('status-note').textContent = snapshot.state === 'templating' ? '正在处理格式要求' : (snapshot.busy ? (mode ? '转换期间请勿操作 Office' : '正在本机处理文档') : '文档转换在本机处理');
    $('progress-fill').style.width = `${snapshot.progress}%`;
    $('progress-track').setAttribute('aria-valuenow', String(Math.round(snapshot.progress)));
    $('result-card').dataset.state = snapshot.state;
    const titles = {templating:'正在处理模板',idle:'准备好，即可开始',running:'正在整理你的文档',cancelling:'正在清理当前任务',diagnosing:'检查运行环境',success:'文档已准备就绪',warning:'已完成，请留意提示',error:'还差一步',cancelled:'转换已取消'};
    $('result-title').textContent = titles[snapshot.state] || snapshot.status;
    const descriptions = {idle:'选择文档与转换偏好，剩下的交给我们。',running:snapshot.status,cancelling:'等待后台退出，已有成品保持不变。',diagnosing:'检查 Python、Pandoc 和 Office 注册信息。',success:'图片、公式与排版，已汇入一份 Word 文档。',warning:'已生成 Word 文档，下方列出了需要检查的内容。',error:'查看下方原因，调整后可以重新开始。',cancelled:'临时文件已清理，可以开始新的转换。'};
    $('result-description').textContent = snapshot.state === 'idle' && logText ? snapshot.status : descriptions[snapshot.state] || '';
    $('result-path').textContent = snapshot.output; $('result-path').hidden = !snapshot.output;
    $('result-buttons').hidden = !snapshot.output;
    setText('result-error', snapshot.error); $('result-error').hidden = !snapshot.error;
    setText('result-warnings', snapshot.warnings.join('\n\n')); $('result-warnings').hidden = !snapshot.warnings.length;
    $('feature-list').hidden = snapshot.busy || !!snapshot.output || !!snapshot.error;
    $('result-icon').querySelector('use').setAttribute('href', snapshot.state === 'success' ? '#i-check' : '#i-doc');
    if (snapshot.state !== previousState && ['error','warning','diagnosing'].includes(snapshot.state)) showLog();
    previousState = snapshot.state;
    if (!snapshot.close_requested) suppressClosePrompt = false;
    if (snapshot.close_requested && !suppressClosePrompt && !$('close-dialog').open) $('close-dialog').showModal();
  }
  async function poll() {
    try { render(await call('poll', after)); await pollTemplate(); }
    catch (error) { $('status-text').textContent = '连接暂时中断，正在重试…'; }
    setTimeout(poll, busy ? 180 : 450);
  }
  async function startTask(name) {
    if (busy || submitting) return;
    if (!$('stage_timeout').checkValidity()) { $('advanced').open = true; $('stage_timeout').reportValidity(); return; }
    submitting = true; freeze(true); clearTimeout(saveTimer);
    try { await call(name, settings()); if (name === 'diagnose') showLog(); }
    catch (error) { toast(error.message); }
    finally { submitting = false; try { render(await call('poll', after)); } catch { freeze(false); } }
  }
  document.querySelectorAll('[data-pick]').forEach(button => button.addEventListener('click', async () => {
    if (busy) return;
    const kind = button.dataset.pick;
    button.disabled = true;
    try {
      const result = await call('pick_file', kind, $(kind).value || (kind === 'output' ? autoOutput : ''));
      if (result.path) { $(kind).value = result.path; if (kind === 'markdown') updateSource(); scheduleSave(); }
    } catch (error) { toast(error.message); }
    finally { button.disabled = false; }
  }));
  document.querySelectorAll('[data-mode]').forEach(button => button.addEventListener('click', () => {
    if (busy) return; mode = button.dataset.mode === 'mathtype'; updateMode(); scheduleSave();
  }));
  $('markdown').addEventListener('input', updateSource);
  let templateRevision = 0;
  async function pollTemplate() {
    // The method guard also supports older desktop hosts during UI tests.
    if (!api.template_snapshot) return;
    const result = await call('template_snapshot');
    if (result.revision === templateRevision) return;
    templateRevision = result.revision;
    if (result.draft && !result.error) { $('spec-editor').value = result.draft; $('template-reviewed').checked = false; }
    $('ai-conversation').textContent = result.messages.join('\n\n');
    $('template-result').textContent = result.error || (result.project ? `已保存：${result.project}。请打开 preview.docx 检查样张。` : '草稿已更新，请检查并处理 issues 中的问题。');
    if (result.project) {
      $('format_spec').value = result.project + '/spec.json';
      $('template').value = '';
      scheduleSave();
    }
  }
  async function templateAction(name, ...args) {
    if (busy || submitting) return;
    submitting = true; freeze(true);
    try { await call(name, ...args); }
    catch (error) { toast(error.message); }
    finally { submitting = false; render(await call('poll', after)); }
  }
  $('import-requirements').addEventListener('click', () => templateAction('import_requirements', $('requirements').value.trim()));
  $('ai-send').addEventListener('click', () => templateAction('template_chat', {
    draft:$('spec-editor').value, message:$('ai-message').value, api_key:$('ai-key').value,
    config:{base_url:$('ai_base_url').value, model:$('ai_model').value, api_key_env:$('ai_key_env').value, json_mode:$('ai_json_mode').checked, timeout:60}
  }));
  $('validate-template').addEventListener('click', async () => {
    try {
      const result = await call('validate_template', $('spec-editor').value);
      $('spec-report').textContent = JSON.stringify(result.issues,null,2) + '\n\n' + result.resolved;
      $('spec-report').parentElement.open = true;
    } catch (error) { toast(error.message); }
  });
  $('generate-template').addEventListener('click', () => templateAction('generate_template', $('spec-editor').value, $('template-output').value.trim(), $('template-reviewed').checked));
  $('save-draft').addEventListener('click', async () => {
    try {
      const selection = await call('pick_file','draft_output','');
      if (selection.path) { await call('save_template_draft',$('spec-editor').value,selection.path); toast('协议草稿已保存。'); }
    } catch (error) { toast(error.message); }
  });
  $('spec-editor').addEventListener('input', () => { $('template-reviewed').checked = false; });
  $('cancel-template').addEventListener('click', () => call('cancel_template').catch(error => toast(error.message)));
  document.querySelectorAll('input, select').forEach(element => element.addEventListener('change', () => { updateMode(); scheduleSave(); }));
  $('theme').addEventListener('click', () => { dark = !dark; document.documentElement.dataset.theme = dark ? 'dark' : 'light'; scheduleSave(); });
  $('start').addEventListener('click', () => startTask('start_conversion'));
  $('diagnose').addEventListener('click', () => startTask('diagnose'));
  $('cancel').addEventListener('click', () => call('cancel').catch(error => toast(error.message)));
  $('open-document').addEventListener('click', () => call('open_output', false).catch(error => toast(error.message)));
  $('open-folder').addEventListener('click', () => call('open_output', true).catch(error => toast(error.message)));
  $('open-log').addEventListener('click', () => call('open_log').catch(error => toast(error.message)));
  $('copy-log').addEventListener('click', async () => {
    try { const result = await call('copy_log'); toast(result.source === 'file' ? '已复制完整日志。' : '已复制当前日志。'); }
    catch (error) { toast(error.message); }
  });
  $('keep-working').addEventListener('click', () => { suppressClosePrompt = true; $('close-dialog').close(); call('request_close', false).catch(error => toast(error.message)); });
  $('close-app').addEventListener('click', () => { suppressClosePrompt = true; $('close-dialog').close(); call('request_close', true).catch(error => toast(error.message)); });
  $('close-dialog').addEventListener('cancel', event => { event.preventDefault(); $('keep-working').click(); });
  async function boot() {
    if (ready || !window.pywebview?.api) return;
    api = window.pywebview.api;
    try {
      const initial = await call('bootstrap');
      mode = initial.settings.convert_formulas; dark = initial.settings.dark_mode;
      for (const key of settingKeys) {
        const element = $(key); if (!element) continue;
        if (element.type === 'checkbox') element.checked = initial.settings[key];
        else element.value = initial.settings[key] ?? '';
      }
      document.documentElement.dataset.theme = dark ? 'dark' : 'light';
      ready = true; updateMode(); freeze(false); poll();
    } catch (error) { toast(error.message); $('status-text').textContent = '初始化失败，请重新启动应用。'; }
  }
  window.addEventListener('pywebviewready', boot);
  boot();
})();
