/**
 * 부대 성당 홈페이지 - 구글 Apps Script
 *
 * 이 코드는 구글 스프레드시트의 [확장 프로그램 → Apps Script]에 붙여넣어 사용합니다.
 * - 홈페이지가 '기본정보', '미사시간', '공지사항' 탭의 내용을 읽어갑니다.
 * - 상담 신청은 '상담신청' 탭에 저장됩니다. 이 탭은 홈페이지에 절대 공개되지 않습니다.
 *
 * 처음 한 번: 위쪽 함수 선택 목록에서 'setup'을 고르고 [실행]을 누르세요.
 * 복음 자동 채우기: 'setupGospel'을 한 번 실행하세요. (굿뉴스 매일미사에서 복음 출처·핵심구절을 가져옵니다)
 */

// 새 상담 신청이 들어오면 알림을 받을 이메일 (비워두면 알림 없음)
// 알림 메일에는 신청 내용이 담기지 않고 "새 신청이 있다"는 사실만 갑니다.
const 알림이메일 = '';

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
    jubo: 주보_(),
    updated: new Date().toISOString(),
  });
}

/* ---------- 군종교구 주보 (최신 PDF) ----------
 * 군종교구 홈페이지 주보 목록에서 맨 위(최신) 주보의 제목·발행일·PDF 주소를 가져옵니다.
 * 1시간 동안 저장해 두고 다시 씁니다. 못 가져오면 홈페이지 배너가 숨겨집니다.
 */
const 주보_목록 = 'https://www.gunjong.or.kr/parish/weekly.asp';

function 주보_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('jubo');
  if (hit) return JSON.parse(hit);
  try {
    const res = UrlFetchApp.fetch(주보_목록, { muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() !== 200) return null;
    const j = 주보뽑기_(res.getContentText('UTF-8'));
    if (j) cache.put('jubo', JSON.stringify(j), 3600);
    return j;
  } catch (err) {
    console.warn('주보 불러오기 실패: ' + err.message);
    return null;
  }
}

function 주보뽑기_(html) {
  // 주소 안의 점(.)이 %2E로 적혀 있어 풀어서 확인합니다
  const pdf경로 = (t) => [...t.matchAll(/savepath=([^"'&]+)/gi)]
    .map((m) => { try { return decodeURIComponent(m[1].replace(/\+/g, ' ')); } catch (e) { return ''; } })
    .find((p) => /\.pdf$/i.test(p));
  const 줄 = html.split(/<tr[\s>]/i).slice(1).find((t) => pdf경로(t));
  if (!줄) return null;
  const 경로 = pdf경로(줄);
  const 제목 = 글자만_((줄.match(/weekly_detail\.asp[^>]*>([\s\S]*?)<\/a>/i) || [])[1] || '');
  const 날짜 = (줄.match(/(\d{4}-\d{2}-\d{2})/) || [])[1] || '';
  return { title: 제목, date: 날짜, pdf: 'https://www.gunjong.or.kr' + encodeURI(경로) };
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

/* ---------- 복음 자동 채우기 (굿뉴스 매일미사) ----------
 * 날마다(또는 미사가 있는 날만) 복음 출처와 핵심구절을
 * 굿뉴스 매일미사에서 가져와 '복음' 탭에 채웁니다.
 * - 미사시간·미사변경 탭을 고치면 바로, 그리고 매일 새벽 5시에 한 번 실행됩니다.
 * - 직접 적은 칸은 덮어쓰지 않습니다.
 * - 자동으로 만든 줄은 '자동' 칸에 Y가 표시되고, 미사 날짜가 바뀌면 알아서 정리됩니다.
 *   그 줄을 직접 고치면 Y가 지워져 그대로 남습니다.
 */
const 굿뉴스_주소 = 'https://maria.catholic.or.kr/mi_pr/missa/missa.asp?menu=missa&missatype=DA&gomonth=';
const 취소표시 = /없음|취소|쉼|휴/;
// true: 미사와 상관없이 날마다 그날 복음을 채움 (홈페이지에 '오늘의 복음'이 매일 바뀜)
// false: 우리 성당 미사가 있는 날의 복음만 채움
const 복음_매일 = true;
const 복음_며칠앞까지 = 7;

// 처음 한 번 실행: 자동 실행 예약을 만들고 바로 한 번 채웁니다.
function setupGospel() {
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (['복음편집감지', '복음채우기'].includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('복음편집감지').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('복음채우기').timeBased().everyDays(1).atHour(5).inTimezone('Asia/Seoul').create();
  const n = 복음채우기();
  ss.toast(`복음 자동 채우기를 켰습니다. (${n}개 날짜 확인)`);
}

function 복음편집감지(e) {
  const sh = e && e.range && e.range.getSheet();
  if (!sh) return;
  const 이름 = sh.getName();
  if (이름 === 탭.복음) {
    // 사람이 복음·핵심구절 칸을 고치면 그 줄은 자동 관리에서 빠짐
    const 열 = 복음열_(sh);
    if (!열.자동) return;
    for (let r = e.range.getRow(); r <= e.range.getLastRow(); r++) {
      if (r < 2) continue;
      const c1 = e.range.getColumn(), c2 = e.range.getLastColumn();
      const 고친칸 = [열.복음, 열.핵심구절, 열.요약, 열.전례].some((c) => c && c >= c1 && c <= c2);
      if (고친칸) sh.getRange(r, 열.자동).setValue('');
    }
    return;
  }
  if (이름 === 탭.변경 || 이름 === 탭.미사) 복음채우기();
}

function 복음채우기() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return 0;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(탭.복음);
    if (!sh) return 0;
    const 열 = 복음열_(sh, true);
    const 대상 = 복음_매일 ? 날마다_(복음_며칠앞까지) : 미사날짜들_(ss);   // 'yyyy-MM-dd' 목록
    const 오늘 = 날짜키_(new Date());

    const 값 = sh.getDataRange().getValues();

    // 미사가 없어진 날의 자동 줄 정리 (오늘 이후만, 아래에서부터 지움)
    for (let i = 값.length - 1; i >= 1; i--) {
      const k = 날짜키_(값[i][열.날짜 - 1]);
      if (k && k >= 오늘 && 값[i][열.자동 - 1] === 'Y' && !대상.includes(k)) sh.deleteRow(i + 1);
    }

    // 대상 날짜 채우기
    let 확인 = 0;
    const 최신 = sh.getDataRange().getValues();
    const 위치 = {};
    for (let i = 1; i < 최신.length; i++) {
      const k = 날짜키_(최신[i][열.날짜 - 1]);
      if (k && !위치[k]) 위치[k] = i + 1;
    }
    대상.forEach((k) => {
      const r = 위치[k];
      const 기존 = r ? 최신[r - 1] : null;
      const 출처있음 = 기존 && String(기존[열.복음 - 1]).trim();
      const 구절있음 = 기존 && String(기존[열.핵심구절 - 1]).trim();
      const 전례있음 = 기존 && String(기존[열.전례 - 1] || '').trim();
      if (출처있음 && 구절있음 && 전례있음) return;
      const g = 굿뉴스복음_(k);
      확인++;
      if (!g) return;
      if (r) {
        if (!출처있음 && g.출처) sh.getRange(r, 열.복음).setValue(g.출처);
        if (!구절있음 && g.핵심구절) sh.getRange(r, 열.핵심구절).setValue(g.핵심구절);
        if (!전례있음 && g.전례) sh.getRange(r, 열.전례).setValue(g.전례);
      } else {
        const 새줄 = Array(Math.max(열.자동, 열.전례, sh.getLastColumn())).fill('');
        const [y, m, d] = k.split('-').map(Number);
        새줄[열.날짜 - 1] = new Date(y, m - 1, d);
        새줄[열.복음 - 1] = g.출처;
        새줄[열.핵심구절 - 1] = g.핵심구절;
        새줄[열.전례 - 1] = g.전례;
        새줄[열.자동 - 1] = 'Y';
        sh.appendRow(새줄);
        // 날짜 표시 형식은 첫 줄(2행)에 맞춤
        const 형식 = sh.getLastRow() > 2 ? sh.getRange(2, 열.날짜).getNumberFormat() : '';
        sh.getRange(sh.getLastRow(), 열.날짜).setNumberFormat(형식 && 형식 !== '@' ? 형식 : 'yyyy. m. d');
      }
    });
    return 확인;
  } finally {
    lock.releaseLock();
  }
}

// 오늘부터 n일 동안 날마다
function 날마다_(n) {
  const 오늘 = 글날짜_(날짜키_(new Date()));
  const out = [];
  for (let i = 0; i < n; i++) { const d = new Date(오늘); d.setDate(d.getDate() + i); out.push(날짜키_(d)); }
  return out;
}

// 앞으로 3주 동안 미사가 있는 날짜 (홈페이지 계산 방식과 같음)
function 미사날짜들_(ss) {
  const 요일글자 = ['일', '월', '화', '수', '목', '금', '토'];
  const 미사 = 표읽기_(ss.getSheetByName(탭.미사)).filter((r) => /미사/.test(r['이름'] || ''));
  const 변경 = 표읽기_(ss.getSheetByName(탭.변경))
    .map((r) => ({ 날짜: 글날짜_(r['날짜']), 취소: 취소표시.test(r['시간'] || ''), 이름: r['이름'] || '' }))
    .filter((c) => c.날짜);
  const 오늘 = 글날짜_(날짜키_(new Date()));   // 한국 시간 기준 오늘
  const 주키 = (d) => { const x = new Date(d); x.setDate(x.getDate() - ((x.getDay() + 6) % 7)); return 날짜키_(x); };
  const out = new Set();

  for (let i = 0; i < 21; i++) {
    const d = new Date(오늘); d.setDate(d.getDate() + i);
    미사.forEach((r) => {
      const 요일 = String(r['요일'] || '');
      const 해당 = /매일/.test(요일) || (/평일/.test(요일) && d.getDay() >= 1 && d.getDay() <= 5) ||
        (/주일/.test(요일) && d.getDay() === 0) || 요일.replace(/주일/g, '').includes(요일글자[d.getDay()]);
      if (!해당) return;
      const 바뀜 = 변경.some((c) => 주키(c.날짜) === 주키(d) && (!c.이름 || c.이름 === r['이름']));
      if (!바뀜) out.add(날짜키_(d));
    });
  }
  const 끝 = new Date(오늘); 끝.setDate(끝.getDate() + 60);
  변경.forEach((c) => { if (!c.취소 && c.날짜 >= 오늘 && c.날짜 <= 끝) out.add(날짜키_(c.날짜)); });
  return [...out].sort();
}

// 굿뉴스 매일미사에서 그날 복음의 출처와 핵심구절(꺾쇠 안 문장)을 가져옴
function 굿뉴스복음_(날짜키) {
  try {
    const res = UrlFetchApp.fetch(굿뉴스_주소 + 날짜키, { muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() !== 200) return null;
    return 복음뽑기_(res.getContentText('UTF-8'));
  } catch (err) {
    console.warn('굿뉴스 연결 실패 ' + 날짜키 + ': ' + err.message);
    return null;
  }
}

function 복음뽑기_(html) {
  const i = html.search(/<h3[^>]*>\s*복음\s*<\/h3>/);
  if (i < 0) return null;
  const span = (html.slice(i, i + 6000).match(/<span[^>]*>([\s\S]*?)<\/span>/) || [])[1] || '';
  const 줄 = span.split(/<br\s*\/?>/i).map((s) => 글자만_(s)).filter(Boolean);
  const 핵심구절 = ((줄[0] || '').match(/^<\s*([\s\S]+?)\s*>$/) || [])[1] || '';
  const 출처줄 = 줄.find((l) => /거룩한\s*복음입니다/.test(l)) || '';
  const m = 출처줄.match(/✠?\s*(\S+?)(?:가|이)\s*전한\s*거룩한\s*복음입니다\.?\s*(.*)$/);
  const 출처 = m ? (m[1] + ' ' + m[2].trim()).trim() : '';
  if (!출처 && !핵심구절) return null;
  return { 출처, 핵심구절, 전례: 전례뽑기_(html) };
}

// 전례일: '[(녹) 연중 제32주일(평신도 주일)]' → '(녹)연중 제32주일'
// 끝에 붙은 괄호 설명과 ' - 첫째 미사' 같은 미사 구분은 뺍니다.
function 전례뽑기_(html) {
  const raw = 글자만_((html.match(/<p[^>]*class="[^"]*m_t_today01[^"]*"[^>]*>([\s\S]*?)<\/p>/) || [])[1] || '');
  let t = raw.replace(/^\[\s*/, '').replace(/\s*\]$/, '');
  const 색 = (t.match(/^\(([^)]{1,2})\)\s*/) || [])[0] || '';
  t = t.slice(색.length).replace(/\s+-\s+.*$/, '');
  while (/\s*\([^()]*\)\s*$/.test(t)) t = t.replace(/\s*\([^()]*\)\s*$/, '');
  t = t.trim();
  return t ? 색.replace(/\s+$/, '') + t : '';
}

function 글자만_(s) {
  return String(s)
    .replace(/<[^>]+>/g, '')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/&#39;|&apos;/g, "'").replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&')
    .replace(/\s+/g, ' ').trim();
}

// 복음 탭의 열 위치 (없으면 '자동' 열을 만듦)
function 복음열_(sh, 만들기) {
  const 머리 = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getDisplayValues()[0].map((h) => String(h).trim());
  const 찾기 = (n) => { const i = 머리.indexOf(n); return i < 0 ? 0 : i + 1; };
  const 열 = { 날짜: 찾기('날짜') || 1, 복음: 찾기('복음') || 2, 핵심구절: 찾기('핵심구절') || 3, 요약: 찾기('요약'), 자동: 찾기('자동'), 전례: 찾기('전례') };
  if (!열.자동 && 만들기) {
    열.자동 = 머리.length + 1;
    sh.getRange(1, 열.자동).setValue('자동').setFontWeight('bold').setBackground('#e6ece6')
      .setNote('Y: 굿뉴스에서 자동으로 채운 줄 (직접 고치면 지워지고, 그 줄은 그대로 남습니다)');
    머리[열.자동 - 1] = '자동';
  }
  if (!열.전례 && 만들기) {
    열.전례 = 머리.length + 1;
    sh.getRange(1, 열.전례).setValue('전례').setFontWeight('bold').setBackground('#e6ece6')
      .setNote('그날 전례일 (예: (녹)연중 제32주일). 홈페이지 복음 카드 제목 옆에 표시됩니다.');
  }
  return 열;
}

function 날짜키_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
  const d = 글날짜_(v);
  return d ? Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd') : '';
}

function 글날짜_(t) {
  const m = String(t || '').match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
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
