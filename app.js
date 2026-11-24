/*
 * app.js — 업로드, PPT 파싱, 화면 갱신, API 호출
 * 프롬프트 내용은 prompts.js에 있다.
 */
'use strict';

// ===== 상태 =====
const state = {
  fileName: '',
  slides: [],   // [{ index, text, notes, memo }]
};

// ===== DOM =====
const $ = (id) => document.getElementById(id);
const el = {
  dropzone: $('dropzone'),
  fileInput: $('file-input'),
  uploadStatus: $('upload-status'),
  previewSection: $('preview-section'),
  slideList: $('slide-list'),
  settingsSection: $('settings-section'),
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
