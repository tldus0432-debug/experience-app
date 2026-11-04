const SHEET_NAME = '경험';
const KINDS = ['동아리', '프로젝트', '대외활동', '인턴', '기타'];

const HEADERS = [
  '날짜', '경험 이름', '종류', '역할', '상황(S)', '과제(T)',
  '행동(A)', '결과(R)', '결과 미리보기', '저장 상태', '기타 설명', '경험 ID'
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('나의 경험 정리')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getExperienceSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  } else {
    // A~K 기존 데이터는 보존하고 L열에 ID 열을 추가한다.
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const idRange = sheet.getRange(2, 12, lastRow - 1, 1);
    const ids = idRange.getValues();
    let changed = false;
    ids.forEach(row => {
      if (!row[0]) {
        row[0] = Utilities.getUuid();
        changed = true;
      }
    });
    if (changed) idRange.setValues(ids);
  }
  return sheet;
}

function normalizeForm_(form) {
  form = form || {};
  return {
    id: String(form.id || '').trim(),
    date: String(form.date || '').trim(),
    name: String(form.name || '').trim(),
    kind: String(form.kind || '').trim(),
    role: String(form.role || '').trim(),
    s: String(form.s || '').trim(),
    t: String(form.t || '').trim(),
    a: String(form.a || '').trim(),
    r: String(form.r || '').trim(),
    otherDescription: String(form.otherDescription || '').trim()
  };
}

function saveExperience(form) {
  const f = normalizeForm_(form);
  if (!f.date || !f.name || !f.kind || !f.role) {
    throw new Error('날짜, 경험 이름, 종류, 역할을 입력해 주세요.');
  }
  if (!KINDS.includes(f.kind)) throw new Error('올바른 경험 종류를 선택해 주세요.');
  if (f.kind === '기타' && !f.otherDescription) {
    throw new Error('기타 경험에 대한 설명을 입력해 주세요.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getExperienceSheet_();
    const complete = [f.s, f.t, f.a, f.r].every(value => value.trim() !== '');
    const status = complete ? '완료' : '미완성';
    const preview = f.r ? f.r.substring(0, 100) : '결과 내용 없음';
    const id = f.id || Utilities.getUuid();

    const row = [
      f.date, f.name, f.kind, f.role, f.s, f.t, f.a, f.r,
      preview, status, f.otherDescription, id
    ];

    if (f.id) {
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) throw new Error('수정할 경험을 찾을 수 없습니다.');
      const ids = sheet.getRange(2, 12, lastRow - 1, 1).getValues().flat();
      const index = ids.findIndex(value => String(value) === f.id);
      if (index < 0) throw new Error('수정할 경험을 찾을 수 없습니다.');
      sheet.getRange(index + 2, 1, 1, HEADERS.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    return { ok: true, id, message: f.id ? f.name + ' 수정했습니다.' : f.name + ' 저장했습니다.', status };
  } finally {
    lock.releaseLock();
  }
}

function deleteExperience(id) {
  if (!id) throw new Error('삭제할 경험 ID가 없습니다.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getExperienceSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('삭제할 경험이 없습니다.');
    const ids = sheet.getRange(2, 12, lastRow - 1, 1).getValues().flat();
    const index = ids.findIndex(value => String(value) === String(id));
    if (index < 0) throw new Error('삭제할 경험을 찾을 수 없습니다.');
    sheet.deleteRow(index + 2);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function getDashboardData() {
  const sheet = getExperienceSheet_();
  const lastRow = sheet.getLastRow();
  let experiences = [];

  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    experiences = values.map((row, index) => ({
      row: index + 2,
      id: String(row[11] || ''),
      date: formatDate_(row[0]),
      name: String(row[1] || ''),
      kind: String(row[2] || ''),
      role: String(row[3] || ''),
      s: String(row[4] || ''),
      t: String(row[5] || ''),
      a: String(row[6] || ''),
      r: String(row[7] || ''),
      preview: String(row[8] || ''),
      status: String(row[9] || '미완성'),
      otherDescription: String(row[10] || '')
    })).filter(item => item.name && item.id);
  }

  experiences.sort((a, b) => b.row - a.row);
  const total = experiences.length;
  const completed = experiences.filter(item => item.status === '완료').length;
  const categories = KINDS.map(kind => ({
    name: kind,
    count: experiences.filter(item => item.kind === kind).length
  }));

  return {
    total,
    completed,
    incomplete: total - completed,
    completionRate: total ? Math.round(completed / total * 100) : 0,
    categories,
    experiences
  };
}

function formatDate_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value || '');
}

// Vercel 서버에서만 호출한다. API_SECRET은 Script Properties에 저장한다.
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const expectedSecret = PropertiesService.getScriptProperties().getProperty('API_SECRET');
    if (!expectedSecret || body.secret !== expectedSecret) {
      throw new Error('인증에 실패했습니다.');
    }

    let data;
    if (body.action === 'dashboard') {
      data = getDashboardData();
    } else if (body.action === 'save') {
      data = saveExperience(body.data || {});
    } else if (body.action === 'delete') {
      data = deleteExperience(body.id);
    } else {
      throw new Error('지원하지 않는 요청입니다.');
    }

    return ContentService.createTextOutput(JSON.stringify({ success: true, data }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error && error.message ? error.message : '서버 오류가 발생했습니다.'
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
