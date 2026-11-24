/*
 * prompts.js — AI에게 보내는 프롬프트 템플릿
 *
 * 발표에서 "프롬프트 설계"를 보여줄 수 있도록 화면/통신 코드(app.js)와 분리해 두었다.
 * 프롬프트를 바꿀 때는 PROMPT_VERSION을 올리고 docs/prompt-log.md에 기록한다.
 */

const PROMPT_VERSION = 'v1';

// 한국어 발화 속도 (분당 글자 수, 공백 제외). app.js의 시간 계산에서도 같은 값을 쓴다.
const SPEECH_CHARS_PER_MINUTE = 280;

// 말투 선택값 → AI에게 줄 구체적인 지시
const TONE_GUIDES = {
  '발표체': '"~습니다/~입니다"로 끝나는 정중한 발표 말투. 수업 발표에 어울리게 또박또박.',
  '친근하게': '"~요"로 끝나는 부드러운 말투. 청중에게 말을 거는 듯 자연스럽게, 그래도 반말은 쓰지 않는다.',
};

/**
 * 시스템 프롬프트: AI의 역할과 출력 규칙
 */
function buildSystemPrompt() {
  return [
    '너는 한국 대학생의 수업 발표를 돕는 발표 코치다.',
    '학생이 준 슬라이드 내용으로 실제로 소리 내어 읽을 발표 대본과, 담당 교수님이 할 법한 예상 질문을 만든다.',
    '',
    '출력 규칙:',
    '- 반드시 JSON 객체 하나만 출력한다.',
    '- JSON 앞뒤에 설명, 인사말, 마크다운 코드블록(```)을 붙이지 않는다.',
    '- 문자열 안의 줄바꿈은 \\n으로 이스케이프한다.',
  ].join('\n');
}

/**
 * 직접 붙여넣기 모드용: 웹 채팅(ChatGPT, claude.ai 등)에는 시스템 프롬프트 칸이 없으므로
 * 시스템 프롬프트와 사용자 프롬프트를 한 덩어리로 합친다.
 */
function buildChatPrompt(params) {
  return `${buildSystemPrompt()}\n\n아래 내용으로 작업해 주세요. 답변은 JSON만 보내 주세요.\n\n${buildUserPrompt(params)}`;
}

/**
 * 사용자 프롬프트: 발표 정보 + 작성 규칙 + 출력 형식 + 슬라이드 내용
 * @param {object} p
 * @param {Array<{index:number,text:string,notes:string,memo:string}>} p.slides
 * @param {number} p.minutes 목표 발표 시간(분)
 * @param {string} p.professorStyle 교수님 스타일 (자유 입력)
 * @param {string} p.tone '발표체' | '친근하게'
 */
function buildUserPrompt({ slides, minutes, professorStyle, tone }) {
  const totalSeconds = Math.round(minutes * 60);
  const targetChars = Math.round(minutes * SPEECH_CHARS_PER_MINUTE);
  const style = professorStyle.trim() || '(특별한 정보 없음 — 일반적인 대학 교수님 기준)';

  const slideBlocks = slides.map((s) => {
    const lines = [`### 슬라이드 ${s.index}`];
    lines.push('[슬라이드 텍스트]');
    lines.push(s.text.trim() || '(텍스트 없음 — 이미지/도표 위주 슬라이드일 수 있음)');
    if (s.notes.trim()) {
      lines.push('[발표자 노트]');
      lines.push(s.notes.trim());
    }
    if (s.memo.trim()) {
      lines.push('[학생 보충 메모]');
      lines.push(s.memo.trim());
    }
    return lines.join('\n');
  }).join('\n\n');

  return `## 발표 정보
- 슬라이드 수: ${slides.length}장
- 목표 발표 시간: ${minutes}분 (총 ${totalSeconds}초)
- 말투: ${tone} — ${TONE_GUIDES[tone] || ''}
- 교수님 스타일: ${style}

## 대본 작성 규칙
1. 모든 슬라이드(1~${slides.length})에 대해 빠짐없이 대본을 하나씩 쓴다. index는 아래 슬라이드 번호와 같아야 한다.
2. 한국어 발화 속도는 분당 약 ${SPEECH_CHARS_PER_MINUTE}자(공백 제외)다. 전체 대본 분량은 공백 제외 약 ${targetChars}자가 되도록 맞추고, 내용이 많은 슬라이드에 시간을 더 배분한다.
3. seconds는 그 슬라이드 대본을 읽는 데 걸리는 예상 시간(초, 정수)이다. 모든 seconds의 합이 ${totalSeconds}초에 가깝게(±10%) 한다.
4. 슬라이드 문구를 그대로 읽지 말고, 듣는 사람이 이해하기 쉽게 풀어서 설명한다.
5. 슬라이드 사이에는 자연스럽게 이어지는 연결 문장을 넣는다. 첫 슬라이드는 인사와 주제 소개로, 마지막 슬라이드는 마무리 인사로 끝낸다.
6. 슬라이드·노트·메모에 없는 수치나 사실은 지어내지 않는다. 꼭 필요하면 "[실제 수치 확인 필요]"처럼 대괄호로 표시한다.
7. 텍스트가 없는 슬라이드는 보충 메모를 바탕으로 쓰고, 메모도 없으면 "[이 슬라이드 내용 설명]" 같은 자리표시를 넣는다.

## 예상 질문 작성 규칙
1. 교수님 스타일을 기준으로, 이 발표를 들은 교수님이 실제로 할 법한 질문을 3~5개 만든다.
2. 근거가 약한 주장, 설명이 생략된 부분, 한계점, 실제 적용 가능성처럼 파고들 만한 지점을 고른다.
3. answerHint는 학생이 어떤 방향으로 답하면 좋은지 1~2문장으로 쓴다.

## 출력 형식 (이 JSON 형식만 출력)
{
  "slides": [{ "index": 1, "script": "발표 대본", "seconds": 40 }],
  "questions": [{ "question": "예상 질문", "answerHint": "답변 방향" }]
}

## 슬라이드 내용
${slideBlocks}`;
}
