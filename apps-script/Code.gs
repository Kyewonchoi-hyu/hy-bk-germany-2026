/**
 * 한양대학교 대학원전략처 · 2026년 독일 BK 연구교류 출장 사전 설문 백엔드 (Google Apps Script)
 *
 * - 응답은 이 스크립트가 연결된 구글 시트의 '응답' 탭에 한 줄씩 쌓입니다.
 * - 같은 이메일로 다시 제출하면 기존 줄을 덮어씁니다(수정 횟수 +1, 접수번호는 그대로).
 * - 제출하면 접수번호(BK2026-001 형식)가 자동으로 부여되고, 교수님 이메일로 안내 메일이 발송됩니다.
 *   메일은 이 스크립트를 배포한 구글 계정의 Gmail에서 나갑니다.
 * - 첨부한 PPTX/PDF는 내 구글 드라이브의 '2026 독일 BK 사전설문 첨부파일' 폴더에 저장됩니다.
 * - 관리자 비밀번호는 코드에 적지 않고 '스크립트 속성'(ADMIN_PASSWORD)에만 둡니다.
 *   → 아래 setPassword() 함수의 비밀번호를 바꾼 뒤 한 번 실행하고, 다시 지워 두세요.
 */

const SHEET_NAME  = '응답';
const FOLDER_NAME = '2026 독일 BK 사전설문 첨부파일';
const MAX_FILE    = 10 * 1024 * 1024;
const TZ          = 'Asia/Seoul';
const RECEIPT_PREFIX = 'BK2026-';
const MAIL_FROM_NAME = '한양대학교 대학원전략처 대학원전략팀';
const SURVEY_TITLE   = '2026년 독일 BK 연구교류 출장 사전 설문';

const HEADERS = [
  '접수번호', '최초 제출', '최종 수정', '수정 횟수', '안내메일',
  '성명', '직위', '단과대학', '학과', '이메일', '연락처',
  '1.주요 연구 분야', '1.관심 키워드', '1.협력 희망 분야', '1.자료 파일명', '1.자료 파일 링크', '1.자료 링크',
  '2.협력 희망 기관', '2.협력 방안', '2.기타', '2.구체적 방안',
  '3.수용 여부', '3.인원', '3.시기', '3.대상', '3.조건·비고',
  '4.파견 여부', '4.파견 주체', '4.인원', '4.시기', '4.대상', '4.조건·비고',
  '파일ID'
];

/* ---------- 최초 1회 실행 ---------- */
function setPassword() {
  const PASSWORD     = '여기에-관리자-비밀번호';   // ← 바꾸고 실행 → 실행 후 이 줄의 값은 지워 두세요
  const REPLY_TO     = '';   // (선택) 교수님이 안내 메일에 '답장'하면 받을 주소. 예) 'team@hanyang.ac.kr'
  const NOTIFY_EMAIL = '';   // (선택) 제출될 때마다 숨은참조로 사본을 받을 담당자 주소
  if (PASSWORD.indexOf('여기에') === 0) throw new Error('PASSWORD 값을 먼저 바꿔 주세요.');
  const props = PropertiesService.getScriptProperties();
  props.setProperty('ADMIN_PASSWORD', PASSWORD);
  props.setProperty('REPLY_TO', REPLY_TO);
  props.setProperty('NOTIFY_EMAIL', NOTIFY_EMAIL);
  sheet_(); folder_();
  Logger.log('관리자 비밀번호 저장, 시트·첨부파일 폴더 준비 완료. 오늘 보낼 수 있는 메일: ' + MailApp.getRemainingDailyQuota() + '통');
}

/** (선택) 안내 메일이 어떻게 보이는지 내 메일로 미리 받아 보기 */
function testMail() {
  const me = Session.getActiveUser().getEmail();
  const sample = { name: '홍길동', position: '교수', college: '공과대학', dept: '기계공학부', email: me,
    q1_main: '예시 연구 분야', q1_keywords: '예시, 키워드', q1_wish: '예시 협력 분야',
    q2_insts: ['아헨공대'], q2_types: ['공동연구 프로젝트'], q3_avail: '가능', q4_avail: '조건부 가능' };
  sendReceiptMail_(sample, RECEIPT_PREFIX + '000', false, '예시_연구소개.pptx', Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'));
  Logger.log(me + ' 로 예시 메일을 보냈습니다.');
}

/* ---------- 웹 앱 진입점 ---------- */
function doGet() {
  return json_({ ok: true, config: getConfig_() });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'bad_request' }); }

  try {
    switch (body.action) {
      case 'submit':
        return json_(submit_(body.data || {}, body.file));
      case 'admin_list':
        checkAdmin_(body.password);
        return json_({ ok: true, headers: HEADERS, rows: listRows_(), config: getConfig_() });
      case 'config_set':
        checkAdmin_(body.password);
        setConfig_(body.config || {});
        return json_({ ok: true, config: getConfig_() });
      default:
        return json_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    return json_({ ok: false, error: err.code || 'server_error', message: String(err.message || err) });
  }
}

/* ---------- 제출 ---------- */
function submit_(d, file) {
  const email = clean_(d.email).toLowerCase();
  if (!clean_(d.name) || !clean_(d.position) || !clean_(d.college) || !clean_(d.dept) || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ||
      !clean_(d.q1_main) || !clean_(d.q1_keywords) || !clean_(d.q1_wish) || !clean_(d.q3_avail) || !clean_(d.q4_avail)) {
    throw fail_('missing');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = sheet_();
    const col = name => HEADERS.indexOf(name);
    const values = sh.getDataRange().getValues();
    let rowIndex = -1, prev = null;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][col('이메일')]).toLowerCase() === email) { rowIndex = i + 1; prev = values[i]; break; }
    }

    // 연구 소개 발표자료는 필수: 새 파일, 공유 링크, 기존 제출 파일 중 하나
    if (!(file && file.data) && !clean_(d.q1_link) && !(prev && prev[col('1.자료 파일명')])) throw fail_('missing_file');

    // 첨부파일
    let fileName = prev ? prev[col('1.자료 파일명')] : '';
    let fileUrl  = prev ? prev[col('1.자료 파일 링크')] : '';
    let fileId   = prev ? prev[col('파일ID')] : '';
    if (file && file.data) {
      const name = String(file.name || '');
      const m = name.match(/\.(pptx|pdf)$/i);
      if (!m) throw fail_('bad_file');
      const bytes = Utilities.base64Decode(file.data);
      if (bytes.length > MAX_FILE) throw fail_('too_large');
      const ext = m[1].toLowerCase();
      const mime = ext === 'pdf' ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
      const safe = (clean_(d.name) + '_' + clean_(d.college) + '_' + clean_(d.dept) + '_연구소개').replace(/[\\/:*?"<>|]/g, '');
      const f = folder_().createFile(Utilities.newBlob(bytes, mime, safe + '.' + ext));
      if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (ignore) {} }
      fileName = name; fileUrl = f.getUrl(); fileId = f.getId();
    }

    const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
    const j = v => Array.isArray(v) ? v.map(clean_).join('; ') : clean_(v);
    const receipt = (prev && prev[col('접수번호')]) ? String(prev[col('접수번호')]) : nextReceipt_();
    const row = [
      receipt,
      prev ? prev[col('최초 제출')] : now,
      now,
      prev ? Number(prev[col('수정 횟수')] || 0) + 1 : 0,
      '발송 중',
      j(d.name), j(d.position), j(d.college), j(d.dept), email, j(d.phone),
      j(d.q1_main), j(d.q1_keywords), j(d.q1_wish), fileName, fileUrl, j(d.q1_link),
      j(d.q2_insts), j(d.q2_types), j(d.q2_other), j(d.q2_detail),
      j(d.q3_avail), j(d.q3_count), j(d.q3_terms), j(d.q3_targets), j(d.q3_note),
      j(d.q4_avail), j(d.q4_units), j(d.q4_count), j(d.q4_terms), j(d.q4_targets), j(d.q4_note),
      fileId
    ];
    let target = rowIndex;
    if (rowIndex > 0) sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    else { sh.appendRow(row); target = sh.getLastRow(); }

    // 안내 메일 (실패해도 접수는 유지)
    const updated = rowIndex > 0;
    let mailSent = false, mailNote;
    try {
      if (MailApp.getRemainingDailyQuota() < 1) throw new Error('일일 발송 한도 초과');
      sendReceiptMail_(d, receipt, updated, fileName, now);
      mailSent = true; mailNote = '발송 ' + now;
    } catch (err) {
      mailNote = '실패: ' + String(err.message || err).slice(0, 80);
    }
    sh.getRange(target, col('안내메일') + 1).setValue(mailNote);
    return { ok: true, updated: updated, receipt: receipt, mailSent: mailSent };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 접수번호 · 안내 메일 ---------- */
function nextReceipt_() {
  const props = PropertiesService.getScriptProperties();
  const n = Number(props.getProperty('RECEIPT_SEQ') || 0) + 1;   // 호출부가 스크립트 잠금을 잡고 있음
  props.setProperty('RECEIPT_SEQ', String(n));
  return RECEIPT_PREFIX + ('000' + n).slice(-Math.max(3, String(n).length));
}

function sendReceiptMail_(d, receipt, updated, fileName, when) {
  const props = PropertiesService.getScriptProperties();
  const cfg = getConfig_();
  const h = v => String(v === undefined || v === null ? '' : (Array.isArray(v) ? v.join(', ') : v))
    .replace(/^'/, '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const pv = v => String(v === undefined || v === null ? '' : (Array.isArray(v) ? v.join(', ') : v)).replace(/^'/, '');
  const rows = [
    ['접수 일시', when + (updated ? ' (수정 접수)' : '')],
    ['성명', d.name], ['직위', d.position], ['소속', [d.college, d.dept].filter(Boolean).join(' ')],
    ['주요 연구 분야', d.q1_main], ['관심 분야 키워드', d.q1_keywords], ['향후 협력 희망 분야', d.q1_wish],
    ['연구 소개 발표자료', fileName || d.q1_link || ''],
    ['협력 희망 기관', d.q2_insts], ['협력 방안', d.q2_types],
    ['3. 학생·연구자 수용', d.q3_avail], ['4. 학생·연구자 파견', d.q4_avail]
  ].filter(r => h(r[1]) !== '');
  const contact = cfg.contact || '대학원전략처 대학원전략팀';
  const due = cfg.due || '2026.10.22.(목) 오후 5시';
  const subject = '[한양대 대학원전략처] ' + SURVEY_TITLE + ' ' + (updated ? '수정 접수' : '접수 완료') + ' (' + receipt + ')';

  const html =
    '<div style="font-family:\'Malgun Gothic\',\'Apple SD Gothic Neo\',sans-serif;max-width:640px;margin:0 auto;color:#18212F;line-height:1.7">' +
      '<div style="background:#0E4A84;color:#fff;padding:22px 24px;border-radius:10px 10px 0 0">' +
        '<div style="font-size:12px;letter-spacing:.08em;opacity:.85">HANYANG UNIVERSITY · 대학원전략처</div>' +
        '<div style="font-size:19px;font-weight:700;margin-top:4px">' + h(SURVEY_TITLE) + '</div>' +
      '</div>' +
      '<div style="border:1px solid #DCE2EA;border-top:0;border-radius:0 0 10px 10px;padding:24px">' +
        '<p style="margin:0 0 14px">' + h(d.name) + ' 교수님, 안녕하십니까.<br>설문이 정상적으로 ' + (updated ? '수정 접수' : '접수') + '되었습니다. 바쁘신 중에 응답해 주셔서 감사합니다.</p>' +
        '<div style="background:#E9F0F8;border-radius:8px;padding:14px 16px;margin:0 0 18px">' +
          '<div style="font-size:12px;color:#5B6678">접수번호</div>' +
          '<div style="font-size:24px;font-weight:800;color:#0E4A84;letter-spacing:.02em">' + h(receipt) + '</div>' +
        '</div>' +
        '<table style="width:100%;border-collapse:collapse;font-size:14px">' +
          rows.map(r => '<tr><th style="text-align:left;vertical-align:top;padding:8px 10px;border-bottom:1px solid #EEF1F5;color:#5B6678;font-weight:600;width:34%">' + h(r[0]) + '</th>' +
                        '<td style="padding:8px 10px;border-bottom:1px solid #EEF1F5;white-space:pre-wrap">' + h(r[1]) + '</td></tr>').join('') +
        '</table>' +
        '<p style="margin:18px 0 0;font-size:13.5px;color:#5B6678">내용을 고치시려면 ' + h(due) + '까지 설문 페이지에서 <b>같은 이메일</b>로 다시 제출해 주십시오. 접수번호는 그대로 유지됩니다.<br>문의: ' + h(contact) + '</p>' +
      '</div>' +
      '<p style="font-size:12px;color:#8A94A6;text-align:center;margin:14px 0 0">본 메일은 설문 제출 시 자동으로 발송되었습니다.</p>' +
    '</div>';

  const text = d.name + ' 교수님, 설문이 정상적으로 ' + (updated ? '수정 접수' : '접수') + '되었습니다.\n\n' +
    rows.map(r => r[0] + ': ' + pv(r[1])).join('\n') +
    '\n\n수정은 ' + due + '까지 같은 이메일로 다시 제출해 주십시오.\n문의: ' + contact;

  const opt = { to: String(d.email).trim(), subject: subject, htmlBody: html, body: text, name: MAIL_FROM_NAME };
  const replyTo = props.getProperty('REPLY_TO'); if (replyTo) opt.replyTo = replyTo;
  const bcc = props.getProperty('NOTIFY_EMAIL'); if (bcc) opt.bcc = bcc;
  MailApp.sendEmail(opt);
}

/* ---------- 관리자 ---------- */
function checkAdmin_(pw) {
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!real) throw fail_('no_password_set');
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('admin_fails') || 0);
  if (fails >= 10) throw fail_('locked');
  if (String(pw || '') !== real) {
    cache.put('admin_fails', String(fails + 1), 600);   // 10분간 10회 틀리면 잠금
    Utilities.sleep(700);
    throw fail_('wrong_password');
  }
}

function listRows_() {
  const values = sheet_().getDataRange().getDisplayValues();
  const head = values[0];
  return values.slice(1).filter(r => r.join('') !== '').map(r => {
    const o = {}; head.forEach((h, i) => { o[h] = r[i]; }); return o;
  }).reverse();   // 최근 제출이 위로
}

function getConfig_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('CONFIG') || '{}'); }
  catch (e) { return {}; }
}
function setConfig_(c) {
  const v = { institution: clean_(c.institution).slice(0, 200), due: clean_(c.due).slice(0, 100), contact: clean_(c.contact).slice(0, 200) };
  PropertiesService.getScriptProperties().setProperty('CONFIG', JSON.stringify(v));
}

/* ---------- 도우미 ---------- */
function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#E9F0F8');
    sh.hideColumns(HEADERS.length);   // 파일ID 열 숨김
  }
  return sh;
}
function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}
function clean_(v) {
  let s = (v === null || v === undefined) ? '' : String(v).trim().slice(0, 3000);
  if (/^[=+\-@]/.test(s)) s = "'" + s;   // 시트 수식 실행 방지
  return s;
}
function fail_(code) { const e = new Error(code); e.code = code; return e; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
