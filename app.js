/*
 * app.js — 업로드, PPT 파싱, 화면 갱신, API 호출
 * 프롬프트 내용은 prompts.js에 있다.
 */
'use strict';

// ===== 상태 =====
const state = {
  fileName: '',
  slides: [],   // [{ index, text, notes, memo }]
  results: [],  // 생성 결과 기록 (설정을 바꿔 다시 생성하면 계속 쌓인다)
  busy: false,
};
// 주의: API 키는 state에 넣지 않는다. 호출할 때마다 입력칸에서 읽어 쓰고 버린다.

// ===== DOM =====
const $ = (id) => document.getElementById(id);
const el = {
  dropzone: $('dropzone'),
  fileInput: $('file-input'),
  uploadStatus: $('upload-status'),
  previewSection: $('preview-section'),
  slideList: $('slide-list'),
  settingsSection: $('settings-section'),
  minutes: $('minutes'),
  professorStyle: $('professor-style'),
  provider: $('provider'),
  model: $('model'),
  apiKey: $('api-key'),
  generateBtn: $('generate-btn'),
  generateStatus: $('generate-status'),
  errorBox: $('error-box'),
  errorMessage: $('error-message'),
  rawDetails: $('raw-details'),
  rawOutput: $('raw-output'),
  retryBtn: $('retry-btn'),
  resultSection: $('result-section'),
  resultGrid: $('result-grid'),
  compareWrap: $('compare-wrap'),
  compareSelect: $('compare-select'),
};

// 작은 DOM 생성 도우미 (텍스트는 항상 textContent로 넣어 HTML 주입을 막는다)
function h(tag, attrs = {}, ...children) {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'class') node.className = v;
    else if (k.startsWith('on')) node.addEventListener(k.slice(2), v);
    else node.setAttribute(k, v);
  }
  for (const c of children.flat()) {
    if (c == null || c === false) continue;
    node.append(c instanceof Node ? c : document.createTextNode(String(c)));
  }
  return node;
}

// ======================================================================
// 1. PPT 파싱 (JSZip)
// ======================================================================
const NS = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
};

function parseXml(text) {
  return new DOMParser().parseFromString(text, 'application/xml');
}

// "ppt/slides" + "../notesSlides/notesSlide1.xml" → "ppt/notesSlides/notesSlide1.xml"
function resolvePath(baseDir, target) {
  if (target.startsWith('/')) return target.slice(1);
  const out = [];
  for (const part of `${baseDir}/${target}`.split('/')) {
    if (part === '..') out.pop();
    else if (part && part !== '.') out.push(part);
  }
  return out.join('/');
}

// .rels 파일을 읽어 [{ id, type, target }] 목록으로 만든다
async function readRels(zip, relsPath, baseDir) {
  const file = zip.file(relsPath);
  if (!file) return [];
  const doc = parseXml(await file.async('string'));
  return [...doc.getElementsByTagNameNS(NS.rel, 'Relationship')].map((r) => ({
    id: r.getAttribute('Id'),
    type: r.getAttribute('Type') || '',
    target: resolvePath(baseDir, r.getAttribute('Target') || ''),
  }));
}

// 슬라이드 순서: presentation.xml의 sldIdLst 순서를 따른다.
// (파일 이름 번호와 실제 순서가 다를 수 있음. 실패하면 파일 번호 순으로 정렬)
async function getSlidePaths(zip) {
  try {
    const presFile = zip.file('ppt/presentation.xml');
    const rels = await readRels(zip, 'ppt/_rels/presentation.xml.rels', 'ppt');
    if (presFile && rels.length) {
      const pres = parseXml(await presFile.async('string'));
      const relMap = new Map(rels.map((r) => [r.id, r.target]));
      const paths = [...pres.getElementsByTagNameNS(NS.p, 'sldId')]
        .map((s) => relMap.get(s.getAttributeNS(NS.r, 'id') || s.getAttribute('r:id')))
        .filter((p) => p && zip.file(p));
      if (paths.length) return paths;
    }
  } catch (e) {
    console.warn('presentation.xml 순서 읽기 실패, 파일 번호 순으로 정렬합니다.', e);
  }
  return Object.keys(zip.files)
    .map((name) => ({ name, m: name.match(/^ppt\/slides\/slide(\d+)\.xml$/) }))
    .filter((x) => x.m)
    .sort((a, b) => Number(a.m[1]) - Number(b.m[1]))
    .map((x) => x.name);
}

// <a:p> 문단 단위로 <a:t> 텍스트를 모아 줄바꿈으로 잇는다
function extractParagraphs(root) {
  const lines = [];
  for (const p of root.getElementsByTagNameNS(NS.a, 'p')) {
    const line = [...p.getElementsByTagNameNS(NS.a, 't')].map((t) => t.textContent).join('');
    if (line.trim()) lines.push(line.trim());
  }
  return lines.join('\n');
}

// 발표자 노트: 슬라이드 번호·머리글 같은 자리표시자는 빼고 본문만 가져온다
const NOTES_SKIP_PLACEHOLDERS = new Set(['sldNum', 'sldImg', 'hdr', 'ftr', 'dt']);

async function extractNotes(zip, slidePath) {
  const dir = slidePath.slice(0, slidePath.lastIndexOf('/'));
  const fileName = slidePath.slice(slidePath.lastIndexOf('/') + 1);
  const rels = await readRels(zip, `${dir}/_rels/${fileName}.rels`, dir);
  const notesRel = rels.find((r) => r.type.endsWith('/notesSlide'));
  const notesFile = notesRel && zip.file(notesRel.target);
  if (!notesFile) return '';

  const doc = parseXml(await notesFile.async('string'));
  const parts = [];
  for (const sp of doc.getElementsByTagNameNS(NS.p, 'sp')) {
    const ph = sp.getElementsByTagNameNS(NS.p, 'ph')[0];
    if (ph && NOTES_SKIP_PLACEHOLDERS.has(ph.getAttribute('type'))) continue;
    const text = extractParagraphs(sp);
    if (text) parts.push(text);
  }
  return parts.join('\n');
}

async function parsePptx(file) {
  const zip = await JSZip.loadAsync(file);
  const paths = await getSlidePaths(zip);
  if (!paths.length) throw new Error('슬라이드를 찾지 못했습니다. 올바른 .pptx 파일인지 확인해 주세요.');

  const slides = [];
  for (let i = 0; i < paths.length; i++) {
    const doc = parseXml(await zip.file(paths[i]).async('string'));
    slides.push({
      index: i + 1,
      text: extractParagraphs(doc),
      notes: await extractNotes(zip, paths[i]),
      memo: '',
    });
  }
  return slides;
}

// ======================================================================
// 2. 업로드 & 슬라이드 미리보기
// ======================================================================
function setStatus(node, message, isError = false) {
  node.textContent = message;
  node.classList.toggle('error', isError);
}

async function handleFile(file) {
  if (!file) return;
  if (!/\.pptx$/i.test(file.name)) {
    setStatus(el.uploadStatus, '.pptx 파일만 올릴 수 있습니다. (.ppt는 PowerPoint에서 .pptx로 다시 저장해 주세요)', true);
    return;
  }
  setStatus(el.uploadStatus, `"${file.name}" 읽는 중...`);
  try {
    state.slides = await parsePptx(file);
    state.fileName = file.name;
    const withNotes = state.slides.filter((s) => s.notes).length;
    setStatus(el.uploadStatus,
      `"${file.name}" — 슬라이드 ${state.slides.length}장을 읽었습니다.` +
      (withNotes ? ` (발표자 노트 ${withNotes}장 포함)` : ''));
    renderSlides();
    el.previewSection.classList.remove('hidden');
    el.settingsSection.classList.remove('hidden');
  } catch (e) {
    console.error(e);
    setStatus(el.uploadStatus, `파일을 읽지 못했습니다: ${e.message}`, true);
  }
}

function renderSlides() {
  el.slideList.replaceChildren(...state.slides.map((s) => {
    const isEmpty = !s.text.trim();
    const memo = h('textarea', {
      rows: 2,
      placeholder: isEmpty ? '이미지 위주 슬라이드예요. 어떤 내용인지 적어 주세요.' : '보충 메모 (선택)',
      oninput: (e) => { s.memo = e.target.value; },
    });
    memo.value = s.memo;
    return h('article', { class: 'slide-card' },
      h('header', {},
        `슬라이드 ${s.index}`,
        isEmpty ? h('span', { class: 'badge' }, '텍스트 없음') : null),
      h('pre', { class: `slide-text${isEmpty ? ' empty' : ''}` }, isEmpty ? '추출된 텍스트가 없습니다.' : s.text),
      s.notes ? h('div', { class: 'slide-notes' }, h('b', {}, '노트: '), s.notes) : null,
      memo);
  }));
}

// 드래그 앤 드롭
el.fileInput.addEventListener('change', (e) => handleFile(e.target.files[0]));
['dragenter', 'dragover'].forEach((type) => el.dropzone.addEventListener(type, (e) => {
  e.preventDefault();
  el.dropzone.classList.add('dragover');
}));
['dragleave', 'drop'].forEach((type) => el.dropzone.addEventListener(type, (e) => {
  e.preventDefault();
  el.dropzone.classList.remove('dragover');
}));
el.dropzone.addEventListener('drop', (e) => handleFile(e.dataTransfer.files[0]));

// ======================================================================
// 3. AI 호출 (Claude / OpenAI)
// ======================================================================
const DEFAULT_MODELS = {
  claude: 'claude-opus-5-5',
  openai: 'gpt-4o',
};

// 서버 측 자동 대체(fallbacks)를 지원하는 Claude 모델.
// 안전 분류기가 요청을 거절(refusal)하면 서버가 다른 모델로 다시 시도해 준다.
const CLAUDE_FALLBACK_MODELS = new Set([
  'claude-opus-5-5', 'claude-opus-5', 'claude-fable-5-2', 'claude-fable-5-1', 'claude-sonnet-5-5',
]);

async function readErrorMessage(res) {
  try {
    const data = await res.json();
    return data?.error?.message || JSON.stringify(data);
  } catch {
    return res.statusText;
  }
}

function httpErrorText(status, message) {
  const tips = {
    401: 'API 키가 올바른지 확인해 주세요.',
    403: '이 키로 해당 모델을 쓸 권한이 없습니다.',
    404: '모델 이름이 올바른지 확인해 주세요.',
    429: '요청이 너무 많거나 크레딧이 부족합니다. 잠시 후 다시 시도해 주세요.',
  };
  return `API 오류 (${status}): ${message}${tips[status] ? `\n→ ${tips[status]}` : ''}`;
}

// Claude Messages API를 브라우저에서 직접 호출
async function callClaude({ model, apiKey, system, user }) {
  const headers = {
    'content-type': 'application/json',
    'x-api-key': apiKey,
    'anthropic-version': '2026-10-21',
    // 브라우저에서 직접 호출할 때 필요한 헤더 (키가 브라우저에 노출됨을 알고 쓴다는 표시)
    'anthropic-dangerous-direct-browser-access': 'true',
  };
  const body = {
    model,
    max_tokens: 16000,
    system,
    messages: [{ role: 'user', content: user }],
  };
  if (CLAUDE_FALLBACK_MODELS.has(model)) {
    headers['anthropic-beta'] = 'server-side-fallback-2026-07-01';
    body.fallbacks = 'default';
  }

  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST', headers, body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(httpErrorText(res.status, await readErrorMessage(res)));

  const data = await res.json();
  if (data.stop_reason === 'refusal') {
    throw new Error('AI가 이 요청에 대한 응답을 거절했습니다. 슬라이드 내용이나 설정을 바꿔 다시 시도해 주세요.');
  }
  // 응답에는 thinking 블록 등이 섞일 수 있으므로 text 블록만 모은다
  const text = (data.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('');
  return { text, truncated: data.stop_reason === 'max_tokens' };
}

// OpenAI Chat Completions API 호출
async function callOpenAI({ model, apiKey, system, user }) {
  const res = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model,
      messages: [
        { role: 'system', content: system },
        { role: 'user', content: user },
      ],
      response_format: { type: 'json_object' },
      max_completion_tokens: 16000,
    }),
  });
  if (!res.ok) throw new Error(httpErrorText(res.status, await readErrorMessage(res)));

  const data = await res.json();
  const choice = data.choices?.[0];
  return { text: choice?.message?.content || '', truncated: choice?.finish_reason === 'length' };
}

function callAI(provider, args) {
  return provider === 'openai' ? callOpenAI(args) : callClaude(args);
}

// AI 응답 문자열 → { slides, questions }. 형식이 틀리면 예외를 던진다.
function parseAIJson(raw) {
  let text = raw.trim();
  // 코드블록(```json ... ```)으로 감싸서 오는 경우 벗겨 낸다
  text = text.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '');
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start === -1 || end <= start) throw new Error('응답에서 JSON을 찾지 못했습니다.');

  const data = JSON.parse(text.slice(start, end + 1));
  if (!Array.isArray(data.slides)) throw new Error('JSON에 "slides" 배열이 없습니다.');

  const slides = data.slides
    .map((s) => ({
      index: Number(s.index),
      script: String(s.script ?? '').trim(),
      seconds: Math.max(0, Math.round(Number(s.seconds) || 0)),
    }))
    .filter((s) => Number.isFinite(s.index));
  const questions = (Array.isArray(data.questions) ? data.questions : [])
    .map((q) => ({ question: String(q.question ?? '').trim(), answerHint: String(q.answerHint ?? '').trim() }))
    .filter((q) => q.question);
  return { slides, questions };
}

// ======================================================================
// 4. 시간 계산
// ======================================================================
// 한국어 발화 속도(분당 약 280자, 공백 제외)로 계산한 예상 시간(초)
function estimateSeconds(script) {
  const chars = script.replace(/\s/g, '').length;
  return Math.round((chars / SPEECH_CHARS_PER_MINUTE) * 60);
}

function formatTime(totalSeconds) {
  const m = Math.floor(totalSeconds / 60);
  const s = totalSeconds % 60;
  return m ? `${m}분 ${s}초` : `${s}초`;
}

// AI 응답을 슬라이드 순서에 맞춰 정리하고 시간을 계산해 결과 객체를 만든다
function buildResult(parsed, settings, slideCount) {
  const byIndex = new Map(parsed.slides.map((s) => [s.index, s]));
  const slides = [];
  for (let i = 1; i <= slideCount; i++) {
    const s = byIndex.get(i);
    slides.push(s
      ? { index: i, script: s.script, aiSeconds: s.seconds, calcSeconds: estimateSeconds(s.script) }
      : { index: i, script: '', aiSeconds: 0, calcSeconds: 0, missing: true });
  }
  return {
    id: state.results.length + 1,
    createdAt: new Date(),
    settings,
    slides,
    questions: parsed.questions,
    targetSeconds: Math.round(settings.minutes * 60),
    aiTotal: slides.reduce((sum, s) => sum + s.aiSeconds, 0),
    calcTotal: slides.reduce((sum, s) => sum + s.calcSeconds, 0),
  };
}

// ======================================================================
// 5. 생성 흐름
// ======================================================================
function readSettings() {
  return {
    minutes: Number(el.minutes.value),
    professorStyle: el.professorStyle.value,
    tone: document.querySelector('input[name="tone"]:checked').value,
    provider: el.provider.value,
    model: el.model.value.trim(),
  };
}

function showError(message, raw) {
  el.errorMessage.textContent = message;
  el.rawDetails.classList.toggle('hidden', !raw);
  el.rawOutput.textContent = raw || '';
  el.errorBox.classList.remove('hidden');
}

function hideError() {
  el.errorBox.classList.add('hidden');
}

function setBusy(busy) {
  state.busy = busy;
  el.generateBtn.disabled = busy;
  el.retryBtn.disabled = busy;
  el.generateBtn.textContent = busy ? '생성 중...' : '대본 생성하기';
}

async function generate() {
  if (state.busy) return;
  const settings = readSettings();
  const apiKey = el.apiKey.value.trim();

  // 입력 확인
  const problems = [];
  if (!state.slides.length) problems.push('먼저 PPT 파일을 올려 주세요.');
  if (!(settings.minutes > 0)) problems.push('발표 시간을 1분 이상으로 입력해 주세요.');
  if (!settings.model) problems.push('모델 이름을 입력해 주세요.');
  if (!apiKey) problems.push('API 키를 입력해 주세요.');
  if (problems.length) {
    setStatus(el.generateStatus, problems.join(' '), true);
    return;
  }

  hideError();
  setBusy(true);
  setStatus(el.generateStatus, `${settings.model}에게 대본을 요청하는 중... (슬라이드 수에 따라 1~2분 걸릴 수 있어요)`);

  let raw = '';
  try {
    const reply = await callAI(settings.provider, {
      model: settings.model,
      apiKey,
      system: buildSystemPrompt(),
      user: buildUserPrompt({ slides: state.slides, ...settings }),
    });
    raw = reply.text;

    let parsed;
    try {
      parsed = parseAIJson(raw);
    } catch (e) {
      const why = reply.truncated
        ? '응답이 길이 제한에 걸려 중간에 잘렸습니다. 슬라이드 수를 줄이거나 발표 시간을 줄여 다시 시도해 주세요.'
        : `AI 응답을 JSON으로 읽지 못했습니다. (${e.message})`;
      showError(why, raw);
      setStatus(el.generateStatus, '');
      return;
    }

    const result = buildResult(parsed, settings, state.slides.length);
    state.results.push(result);
    renderResults();
    setStatus(el.generateStatus, `완료! 결과 #${result.id}가 아래에 추가되었습니다.`);
    el.resultSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (e) {
    console.error(e);
    // fetch 자체가 실패하면(네트워크/CORS) TypeError가 난다
    const message = e instanceof TypeError
      ? `네트워크 오류: API 서버에 연결하지 못했습니다. 인터넷 연결을 확인해 주세요. (${e.message})`
      : e.message;
    showError(message, raw);
    setStatus(el.generateStatus, '');
  } finally {
    setBusy(false);
  }
}

el.generateBtn.addEventListener('click', generate);
el.retryBtn.addEventListener('click', generate);
el.provider.addEventListener('change', () => {
  el.model.value = DEFAULT_MODELS[el.provider.value];
  el.apiKey.placeholder = el.provider.value === 'openai' ? 'sk-...' : 'sk-ant-...';
});

// ======================================================================
// 6. 결과 화면
// ======================================================================
function timeSummary(result) {
  const { targetSeconds, aiTotal, calcTotal } = result;
  const worst = Math.max(aiTotal, calcTotal);
  let level = 'ok';
  let verdict = '목표 시간에 잘 맞습니다.';
  if (worst > targetSeconds) {
    level = 'over';
    verdict = `⚠ 목표 시간보다 최대 ${formatTime(worst - targetSeconds)} 깁니다. 대본을 줄이거나 발표 시간을 늘려 다시 생성해 보세요.`;
  } else if (worst < targetSeconds * 0.8) {
    level = 'under';
    verdict = `목표 시간보다 ${formatTime(targetSeconds - worst)} 이상 짧습니다. 내용을 보충해도 좋아요.`;
  }
  return h('div', { class: `time-summary ${level}` },
    h('div', { class: 'rows' },
      h('span', {}, '목표 시간'), h('b', {}, formatTime(targetSeconds)),
      h('span', {}, 'AI 예상 합계'), h('span', {}, formatTime(aiTotal)),
      h('span', {}, '글자 수 기준'), h('span', {}, `${formatTime(calcTotal)} (분당 ${SPEECH_CHARS_PER_MINUTE}자)`)),
    h('div', { class: 'verdict' }, verdict));
}

function renderResultColumn(result, { previous = false } = {}) {
  const s = result.settings;
  const time = result.createdAt.toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit' });

  const scripts = result.slides.map((sl) => h('div', { class: `script-item${sl.missing ? ' missing' : ''}` },
    h('header', {},
      `슬라이드 ${sl.index}`,
      sl.missing ? null : h('span', { class: 'secs' }, `AI ${sl.aiSeconds}초 · 글자 수 기준 ${sl.calcSeconds}초`)),
    h('p', {}, sl.missing ? 'AI가 이 슬라이드의 대본을 만들지 않았습니다. 다시 생성해 보세요.' : sl.script)));

  const questions = result.questions.length
    ? h('ol', { class: 'questions' }, result.questions.map((q) => h('li', {},
      h('div', { class: 'q' }, q.question),
      q.answerHint ? h('div', { class: 'a' }, q.answerHint) : null)))
    : h('p', { class: 'hint' }, '예상 질문이 생성되지 않았습니다.');

  return h('article', { class: `result-col${previous ? ' previous' : ''}` },
    h('div', { class: 'result-head' },
      h('div', {},
        h('h3', {}, `결과 #${result.id}${previous ? ' (비교 대상)' : ''}`),
        h('div', { class: 'chips' },
          h('span', { class: 'chip' }, time),
          h('span', { class: 'chip' }, `${s.minutes}분`),
          h('span', { class: 'chip' }, s.tone),
          h('span', { class: 'chip' }, s.model),
          s.professorStyle.trim() ? h('span', { class: 'chip' }, `교수님: ${s.professorStyle.trim()}`) : null)),
      h('div', { class: 'result-actions' },
        h('button', { type: 'button', onclick: (e) => copyResult(result, e.currentTarget) }, '전체 복사'),
        h('button', { type: 'button', onclick: () => downloadResult(result) }, '.txt 다운로드'))),
    timeSummary(result),
    h('h4', {}, '발표 대본'),
    scripts,
    h('h4', {}, '예상 질문'),
    questions);
}

// 비교 선택: 새 결과가 생기면 바로 직전 결과를 기본 비교 대상으로 고른다
function renderResults() {
  const latest = state.results[state.results.length - 1];
  if (!latest) return;
  el.resultSection.classList.remove('hidden');

  const previousOnes = state.results.slice(0, -1);
  el.compareWrap.classList.toggle('hidden', previousOnes.length === 0);
  el.compareSelect.replaceChildren(
    h('option', { value: '' }, '비교 안 함'),
    ...previousOnes.slice().reverse().map((r) => h('option', { value: r.id },
      `결과 #${r.id} (${r.settings.minutes}분 · ${r.settings.tone} · ${r.settings.model})`)));
  el.compareSelect.value = previousOnes.length ? String(previousOnes[previousOnes.length - 1].id) : '';
  renderCompare();
}

function renderCompare() {
  const latest = state.results[state.results.length - 1];
  const other = state.results.find((r) => String(r.id) === el.compareSelect.value);
  el.resultGrid.classList.toggle('compare', Boolean(other));
  const columns = [renderResultColumn(latest)];
  if (other) columns.push(renderResultColumn(other, { previous: true }));
  el.resultGrid.replaceChildren(...columns);
}

el.compareSelect.addEventListener('change', renderCompare);

// ======================================================================
// 7. 복사 / 다운로드
// ======================================================================
function resultToText(result) {
  const s = result.settings;
  const lines = [
    `[대본.pptx] ${state.fileName || '발표'} — 발표 대본`,
    `목표 시간: ${formatTime(result.targetSeconds)} / AI 예상: ${formatTime(result.aiTotal)} / 글자 수 기준: ${formatTime(result.calcTotal)}`,
    `말투: ${s.tone} · 모델: ${s.model}${s.professorStyle.trim() ? ` · 교수님 스타일: ${s.professorStyle.trim()}` : ''}`,
    '',
  ];
  for (const sl of result.slides) {
    lines.push(`■ 슬라이드 ${sl.index} (약 ${sl.missing ? '-' : formatTime(sl.calcSeconds)})`);
    lines.push(sl.missing ? '(대본 없음)' : sl.script);
    lines.push('');
  }
  if (result.questions.length) {
    lines.push('■ 예상 질문');
    result.questions.forEach((q, i) => {
      lines.push(`${i + 1}. ${q.question}`);
      if (q.answerHint) lines.push(`   → 답변 방향: ${q.answerHint}`);
    });
  }
  return lines.join('\n');
}

async function copyResult(result, button) {
  const text = resultToText(result);
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    // file:// 등 clipboard API를 못 쓰는 환경을 위한 대체 방법
    const ta = h('textarea', {});
    ta.value = text;
    document.body.append(ta);
    ta.select();
    document.execCommand('copy');
    ta.remove();
  }
  const original = button.textContent;
  button.textContent = '복사됨!';
  setTimeout(() => { button.textContent = original; }, 1500);
}

function downloadResult(result) {
  const pad = (n) => String(n).padStart(2, '0');
  const d = result.createdAt;
  const stamp = `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}`;
  const base = (state.fileName || '발표').replace(/\.pptx$/i, '');
  const blob = new Blob(['﻿' + resultToText(result)], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = h('a', { href: url, download: `${base}_대본_${stamp}_#${result.id}.txt` });
  document.body.append(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}
