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

const 탭 = {
  정보: '기본정보',
  미사: '미사시간',
  변경: '미사변경',
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

    if (알림이메일) {
      MailApp.sendEmail(
        알림이메일,
        '[성당 홈페이지] 새 상담 신청이 있습니다',
        '새 상담 신청이 1건 접수되었습니다.\n아래 스프레드시트의 "상담신청" 탭에서 확인해 주세요.\n\n' + ss.getUrl()
      );
    }
    return json_({ ok: true, id: 번호 });
  } catch (err) {
    return json_({ ok: false, error: '서버 오류: ' + err.message });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
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
    ['연락처', '010-1234-5678'],
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
