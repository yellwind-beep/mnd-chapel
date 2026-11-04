/**
 * 부대 성당 홈페이지 - 구글 Apps Script
 *
 * 이 코드는 구글 스프레드시트의 [확장 프로그램 → Apps Script]에 붙여넣어 사용합니다.
 * - 홈페이지가 '기본정보', '미사시간', '공지사항' 탭의 내용을 읽어갑니다.
 * - 상담 신청은 '상담신청' 탭에 저장됩니다. 이 탭은 홈페이지에 절대 공개되지 않습니다.
 *
 * 처음 한 번: 위쪽 함수 선택 목록에서 'setup'을 고르고 [실행]을 누르세요.
 */

// 새 상담 신청이 들어오면 알림을 받을 이메일 (비워두면 알림 없음)
// 알림 메일에는 신청 내용이 담기지 않고 "새 신청이 있다"는 사실만 갑니다.
const 알림이메일 = '';

/* ---------- 카카오톡 '나에게 보내기' 알림 ----------
 * 연결 방법은 설정안내.md의 "카카오톡 알림 연결"을 보세요.
 * 아래 값은 Apps Script 화면에서만 채우고, GitHub 등 다른 곳에는 올리지 마세요.
 * kakaoConnect를 한 번 실행하면 값이 안전한 저장소(스크립트 속성)로 옮겨지므로,
 * 그 뒤에는 다시 비워두거나 새 코드로 바꿔 붙여넣어도 알림이 계속 옵니다. */
const KAKAO = {
  REST_KEY: '',        // 카카오 개발자 > 내 애플리케이션 > 앱 키 > REST API 키
  CLIENT_SECRET: '',   // 보안 > Client Secret을 '사용함'으로 해둔 경우에만
  REDIRECT_URI: 'https://mnd-chapel.vercel.app/',
  CODE: '',            // 동의 후 주소창의 ?code= 뒤 값 (10분 안에 한 번만 사용 가능)
};
// true로 바꾸면 카톡 알림에 신청자 이름과 연락 방법도 함께 보냅니다.
// 기본값(false)은 접수번호·분야·방식만 보내고, 상담 내용은 어떤 경우에도 보내지 않습니다.
const 카카오_이름연락포함 = false;

const 탭 = {
  정보: '기본정보',
  미사: '미사시간',
  변경: '미사변경',
  복음: '복음',
  공지: '공지사항',
  상담: '상담신청',
};

const 상담_머리글 = ['접수일시', '접수번호', '이름(세례명)', '소속', '연락방법', '상담분야', '상담방식', '편한시간', '내용', '처리상태', '메모'];

/* ---------- 홈페이지가 내용을 읽어갈 때 ---------- */
function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return json_({
    info: 키값읽기_(ss.getSheetByName(탭.정보)),
    masses: 표읽기_(ss.getSheetByName(탭.미사)),
    changes: 표읽기_(ss.getSheetByName(탭.변경)),
    gospel: 표읽기_(ss.getSheetByName(탭.복음)),
    notices: 표읽기_(ss.getSheetByName(탭.공지)),
    updated: new Date().toISOString(),
  });
}

/* ---------- 상담 신청이 들어올 때 ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    // 스팸 로봇이 숨은 칸을 채운 경우: 성공한 척하고 저장하지 않음
    if (d.website) return json_({ ok: true });

    const 이름 = 자르기_(d.name, 40);
    const 연락 = 자르기_(d.contact, 80);
    if (!이름 || !연락) return json_({ ok: false, error: '이름과 연락 방법은 꼭 필요합니다.' });
    if (d.agree !== true) return json_({ ok: false, error: '개인정보 동의가 필요합니다.' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = 상담탭_(ss);
    const 지금 = new Date();
    const 번호 = Utilities.formatDate(지금, 'Asia/Seoul', 'MMdd') + '-' + Math.floor(100 + Math.random() * 900);

    const 행 = [
      Utilities.formatDate(지금, 'Asia/Seoul', 'yyyy-MM-dd HH:mm'),
      번호,
      이름,
      자르기_(d.unit, 60),
      연락,
      자르기_(d.topic, 20),
      자르기_(d.method, 20),
      자르기_(d.when, 80),
      자르기_(d.message, 1500),
      '접수',
      '',
    ].map(안전하게_);

    const r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, 행.length).setNumberFormat('@').setValues([행]);

    // 알림은 실패해도 신청 접수에는 영향이 없도록 따로 처리
    if (알림이메일) {
      try {
        MailApp.sendEmail(
          알림이메일,
          '[성당 홈페이지] 새 상담 신청이 있습니다',
          '새 상담 신청이 1건 접수되었습니다.\n아래 스프레드시트의 "상담신청" 탭에서 확인해 주세요.\n\n' + ss.getUrl()
        );
      } catch (mailErr) { console.error('메일 알림 실패: ' + mailErr.message); }
    }
    if (카카오연결됨_()) {
      try {
        const 줄 = [
          '🕊️ 새 상담 신청이 들어왔습니다',
          '접수번호: ' + 번호,
          '분야: ' + (d.topic || '-') + ' · 방식: ' + (d.method || '-'),
        ];
        if (카카오_이름연락포함) {
          줄.push('이름: ' + 이름);
          줄.push('연락: ' + 연락);
        }
        줄.push('시트의 "상담신청" 탭에서 확인해 주세요.');
        카카오보내기_(줄.join('\n'), ss.getUrl());
      } catch (kakaoErr) { console.error('카카오 알림 실패: ' + kakaoErr.message); }
    }
    return json_({ ok: true, id: 번호 });
  } catch (err) {
    return json_({ ok: false, error: '서버 오류: ' + err.message });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* ---------- 카카오톡 알림: 직접 실행하는 함수 ---------- */

// 1) 처음 한 번: KAKAO.REST_KEY와 KAKAO.CODE를 채우고 이 함수를 실행
function kakaoConnect() {
  const p = PropertiesService.getScriptProperties();
  const key = KAKAO.REST_KEY.trim() || p.getProperty('KAKAO_REST_KEY');
  const secret = KAKAO.CLIENT_SECRET.trim() || p.getProperty('KAKAO_CLIENT_SECRET') || '';
  if (!key) throw new Error('KAKAO.REST_KEY에 REST API 키를 먼저 넣고 저장하세요.');
  if (!KAKAO.CODE.trim()) throw new Error('KAKAO.CODE에 주소창의 code 값을 넣고 저장하세요.');

  const payload = {
    grant_type: 'authorization_code',
    client_id: key,
    redirect_uri: KAKAO.REDIRECT_URI,
    code: KAKAO.CODE.trim(),
  };
  if (secret) payload.client_secret = secret;
  const j = 카카오요청_('https://kauth.kakao.com/oauth/token', payload);
  if (!j.access_token) throw new Error('토큰 발급 실패: ' + JSON.stringify(j));

  p.setProperties({
    KAKAO_REST_KEY: key,
    KAKAO_CLIENT_SECRET: secret,
    KAKAO_ACCESS: j.access_token,
    KAKAO_REFRESH: j.refresh_token,
    KAKAO_ACCESS_EXP: String(Date.now() + (Number(j.expires_in) - 300) * 1000),
  });

  // 오랫동안 신청이 없어도 연결이 끊기지 않도록 매주 한 번 자동 갱신
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'kakaoRefresh')
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('kakaoRefresh').timeBased().everyWeeks(1).create();

  카카오보내기_('✅ 성당 홈페이지 상담 알림이 카카오톡에 연결되었습니다.', SpreadsheetApp.getActiveSpreadsheet().getUrl());
  console.log('카카오톡 연결 완료. 카카오톡 "나와의 채팅"을 확인하세요.');
}

// 2) 시험 알림 보내기
function kakaoTest() {
  카카오보내기_('🔔 시험 알림입니다. 상담 신청이 들어오면 이렇게 알려드립니다.', SpreadsheetApp.getActiveSpreadsheet().getUrl());
  console.log('보냈습니다. 카카오톡 "나와의 채팅"을 확인하세요.');
}

// 3) 매주 자동 실행 (직접 실행할 필요 없음)
function kakaoRefresh() {
  if (카카오연결됨_()) 카카오토큰_(true);
}

// 카카오톡 알림을 끄고 싶을 때 실행
function kakaoDisconnect() {
  const p = PropertiesService.getScriptProperties();
  ['KAKAO_ACCESS', 'KAKAO_REFRESH', 'KAKAO_ACCESS_EXP'].forEach((k) => p.deleteProperty(k));
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'kakaoRefresh')
    .forEach((t) => ScriptApp.deleteTrigger(t));
  console.log('카카오톡 알림 연결을 해제했습니다.');
}

function 카카오연결됨_() {
  return !!PropertiesService.getScriptProperties().getProperty('KAKAO_REFRESH');
}

function 카카오토큰_(강제갱신) {
  const p = PropertiesService.getScriptProperties();
  const access = p.getProperty('KAKAO_ACCESS');
  const exp = Number(p.getProperty('KAKAO_ACCESS_EXP') || 0);
  if (!강제갱신 && access && Date.now() < exp) return access;

  const payload = {
    grant_type: 'refresh_token',
    client_id: p.getProperty('KAKAO_REST_KEY'),
    refresh_token: p.getProperty('KAKAO_REFRESH'),
  };
  const secret = p.getProperty('KAKAO_CLIENT_SECRET');
  if (secret) payload.client_secret = secret;
  const j = 카카오요청_('https://kauth.kakao.com/oauth/token', payload);
  if (!j.access_token) throw new Error('카카오 토큰 갱신 실패: ' + JSON.stringify(j));

  const 새값 = {
    KAKAO_ACCESS: j.access_token,
    KAKAO_ACCESS_EXP: String(Date.now() + (Number(j.expires_in) - 300) * 1000),
  };
  if (j.refresh_token) 새값.KAKAO_REFRESH = j.refresh_token; // 만료가 가까우면 새 값이 옴
  p.setProperties(새값);
  return j.access_token;
}

function 카카오보내기_(글, 링크) {
  const token = 카카오토큰_(false);
  const template = {
    object_type: 'text',
    text: String(글).slice(0, 200),
    link: { web_url: 링크 || '', mobile_web_url: 링크 || '' },
    button_title: '시트 열기',
  };
  const j = 카카오요청_('https://kapi.kakao.com/v2/api/talk/memo/default/send',
    { template_object: JSON.stringify(template) }, token);
  if (j.result_code !== 0) throw new Error('카카오톡 전송 실패: ' + JSON.stringify(j));
}

function 카카오요청_(url, payload, token) {
  const opt = { method: 'post', payload: payload, muteHttpExceptions: true };
  if (token) opt.headers = { Authorization: 'Bearer ' + token };
  const res = UrlFetchApp.fetch(url, opt);
  try { return JSON.parse(res.getContentText()); }
  catch (e) { return { error: res.getResponseCode(), body: res.getContentText().slice(0, 300) }; }
}

/* ---------- 처음 한 번 실행: 탭과 예시 내용 만들기 ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Seoul');

  탭만들기_(ss, 탭.정보, [
    ['항목', '내용'],
    ['성당이름', '국방부 직원 성당'],
    ['소개', '환영합니다'],
    ['주소', '국방부 후문 앞 군종교구청 1층 경당'],
    ['오시는길', '군종교구청으로 오셔서 전화 주세요.'],
    ['연락처', '010-9443-1239'],
    ['사무실시간', '화–금 12:00–13:00'],
    ['지도링크', 'https://map.naver.com/p/search/%EC%B2%9C%EC%A3%BC%EA%B5%90%20%EA%B5%B0%EC%A2%85%EA%B5%90%EA%B5%AC%EC%B2%AD'],
  ]);
  탭만들기_(ss, 탭.미사, [
    ['이름', '요일', '시간', '장소', '비고'],
    ['직원미사', '목', '11:45', '군종교구청 경당', ''],
  ]);
  탭만들기_(ss, 탭.공지, [
    ['날짜', '제목', '내용', '고정', '이미지'],
    [Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd'), '직원미사 안내', '직원미사는 매주 목요일 11:45에 있습니다.\n수요일이나 금요일로 바뀌는 주에는 홈페이지 첫 화면과 미사 시간에 따로 안내합니다.', 'Y', ''],
    ['2026-11-04', '미사 후 컵라면·김밥 나눔', '이번 주 미사(11월 5일 목요일)부터 미사 후에 컵라면과 김밥을 드립니다.\n점심시간에 드리는 미사인 만큼, 미사 마치고 간단히 요기하고 가세요.', '', 'images/ramen-gimbap.svg'],
  ]);

  // 미사 요일이 바뀌는 주에만 한 줄씩 적는 탭
  if (!ss.getSheetByName(탭.변경)) {
    탭만들기_(ss, 탭.변경, [['날짜', '시간', '장소', '안내', '이름']]);
    const ch = ss.getSheetByName(탭.변경);
    ch.getRange('A2:A').setNumberFormat('yyyy-mm-dd')
      .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false)
        .setHelpText('미사가 있을 날짜를 입력하세요. 예: 2026-11-13').build());
    ch.getRange('A1').setNote(
      '요일이 바뀌는 주에만 한 줄 추가하세요.\n' +
      '· 날짜만 적으면 평소 시간·장소로 안내됩니다.\n' +
      '· 그 주의 평소 목요일 미사는 홈페이지에서 자동으로 빠집니다.\n' +
      '· 그 주에 미사가 없으면 시간 칸에 "없음"을 적으세요.\n' +
      '· 안내 칸에 적은 문장이 홈페이지에 함께 표시됩니다.\n' +
      '· 지난 날짜는 자동으로 표시되지 않으니 지우지 않아도 됩니다.'
    );
    ch.setColumnWidths(1, 5, 140);
    ch.setColumnWidth(4, 320);
  }

  // 미사 날마다 복음 핵심 구절과 요약을 적는 탭
  if (!ss.getSheetByName(탭.복음)) {
    탭만들기_(ss, 탭.복음, [['날짜', '복음', '핵심구절', '요약']]);
    const gs = ss.getSheetByName(탭.복음);
    gs.getRange('A2:A').setNumberFormat('yyyy-mm-dd')
      .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false)
        .setHelpText('미사 날짜를 입력하세요. 예: 2026-11-12').build());
    gs.getRange('A1').setNote(
      '미사가 있는 날마다 한 줄씩 적으세요.\n' +
      '· 홈페이지 첫 화면에 오늘 이후 가장 가까운 날짜의 복음이 나옵니다.\n' +
      '· 복음: 성경 출처 (예: 루카 16,1-8)\n' +
      '· 핵심구절: 매일미사 본문에서 한두 문장\n' +
      '· 요약: 두 줄 정도 (칸 안 줄바꿈은 Ctrl+Enter)\n' +
      '· 지난 날짜는 자동으로 표시되지 않습니다.'
    );
    gs.appendRow([
      new Date(2026, 10, 6), '루카 16,1-8', '',
      '재산을 낭비했다는 고발을 받은 집사가 해고를 앞두고, 주인에게 빚진 이들의 빚을 덜어 주며 앞날을 대비합니다.\n' +
      '예수님은 세상 사람들의 이런 영리함을 들어, 신앙인인 우리도 하느님 나라를 위해 지혜롭게 준비하라고 하십니다.',
    ]);
    gs.setColumnWidth(1, 110);
    gs.setColumnWidth(2, 120);
    gs.setColumnWidth(3, 320);
    gs.setColumnWidth(4, 420);
    gs.getRange('C:D').setWrap(true);
  }
  상담탭_(ss);

  // 처리상태 칸을 목록에서 고를 수 있게
  const sh = ss.getSheetByName(탭.상담);
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(['접수', '연락함', '상담완료'], true).build();
  sh.getRange('J2:J').setDataValidation(rule);

  ss.toast('초기설정이 끝났습니다. 각 탭의 내용을 고쳐 쓰세요.');
}

/* ---------- 내부 도우미 ---------- */
function 표읽기_(sh) {
  if (!sh) return [];
  const v = sh.getDataRange().getDisplayValues();
  const 머리 = (v.shift() || []).map((h) => String(h).trim());
  return v
    .filter((r) => r.some((c) => String(c).trim()))
    .map((r) => {
      const o = {};
      머리.forEach((h, i) => { if (h) o[h] = String(r[i] || '').trim(); });
      return o;
    });
}

function 키값읽기_(sh) {
  const o = {};
  if (!sh) return o;
  sh.getDataRange().getDisplayValues().slice(1).forEach((r) => {
    const k = String(r[0] || '').trim();
    if (k) o[k] = String(r[1] || '').trim();
  });
  return o;
}

function 상담탭_(ss) {
  let sh = ss.getSheetByName(탭.상담);
  if (!sh) {
    sh = ss.insertSheet(탭.상담);
    sh.appendRow(상담_머리글);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 상담_머리글.length).setFontWeight('bold').setBackground('#e6ece6');
  }
  return sh;
}

function 탭만들기_(ss, 이름, 행들) {
  if (ss.getSheetByName(이름)) return; // 이미 있으면 건드리지 않음
  const sh = ss.insertSheet(이름);
  const 폭 = Math.max(...행들.map((r) => r.length));
  const 채움 = 행들.map((r) => r.concat(Array(폭 - r.length).fill('')));
  sh.getRange(1, 1, 채움.length, 폭).setNumberFormat('@').setValues(채움);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, 폭).setFontWeight('bold').setBackground('#e6ece6');
  sh.autoResizeColumns(1, 폭);
}

function 자르기_(v, n) {
  return String(v == null ? '' : v).trim().slice(0, n);
}

// '=', '+', '-', '@'로 시작하는 값이 수식으로 실행되지 않도록
function 안전하게_(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
