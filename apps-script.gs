/**
 * 안녕하시렵니까 길드 포털 — 서버 스크립트 (최종본 / VERSION 22)
 *
 * 구조:  GitHub Pages(화면)  ──fetch──>  이 스크립트(API)  ──>  Google Sheets(데이터)
 *
 * 설치:
 *   1) 아래 코드를 Code.gs에 전체 붙여넣기
 *   2) SPREADSHEET_ID / SECRET / PW / PW_MEMBER 확인
 *   3) 배포 → 배포 관리 → 연필(수정) → 버전 "새 버전" → 배포
 *      (※ "새 배포"를 누르면 주소가 바뀌므로 주의)
 *
 * 포함 기능: 토큰 인증(비밀번호 서버 보관) · 운영진/길드원 2단계 권한 ·
 *            시트 양방향 동기화 · 3-way 병합 저장 · 드롭다운 목록 실시간 로딩
 */
const PW = 'CHANGE_ME_ADMIN';        // 운영진 비밀번호 (반드시 변경)
const PW_MEMBER = 'CHANGE_ME_MEMBER';  // 길드원 비밀번호 (반드시 변경)
const KEY_SHEET = 'WEB_DATA';
const SPREADSHEET_ID = '1RgqYWLdRor3jA8_V-HsSRjAi_eIxgxJ83dpM-90EsTE'; // 안녕하시렵니까 구글 시트
const VERSION = 22;   // 최종본

const SECRET = 'CHANGE_ME_TO_A_LONG_RANDOM_STRING'; // 토큰 서명용 (외부 비노출, 설정 후 변경 금지)
const TOKEN_HOURS = 12;

function getSS() { return SpreadsheetApp.openById(SPREADSHEET_ID); }

function sign_(s) {
  return Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(s, SECRET));
}
function makeToken_(role) {
  const ts = Date.now();
  return role + '|' + ts + '|' + sign_(role + '|' + ts);
}
function verifyToken_(t) {
  if (!t) return null;
  const p = String(t).split('|');
  if (p.length !== 3) return null;
  if (sign_(p[0] + '|' + p[1]) !== p[2]) return null;
  if (Date.now() - Number(p[1]) > TOKEN_HOURS * 3600 * 1000) return null;
  return (p[0] === 'admin' || p[0] === 'member') ? p[0] : null;
}

// v8: 캐릭터/스킬/경매는 시트 탭에서 직접 읽음(완전 양방향) · 공지는 WEB_DATA에서
function loadData() {
  const ss = getSS();
  const d = { basic: readBasic(ss), skills: readSkills(ss) };
  const a = readAuction(ss);
  d.mau = a.mau;
  try {
    const ws = ss.getSheetByName(KEY_SHEET);
    if (ws) { const v = ws.getRange(1, 1).getValue(); if (v) { const old = JSON.parse(v); d.notice = old.notice || null; d.guildName = old.guildName || null; d.subTitle = old.subTitle || null; d.parties = old.parties || null; d.tabsHidden = old.tabsHidden || []; d.rosterCols = old.rosterCols || null; d.auctionPrices = old.auctionPrices || null; d.tabPerm = old.tabPerm || null; } }
  } catch (err) {}
  if (!d.basic.length && !d.notice) return null;
  return d;
}

function readBasic(ss) {
  const sh = ss.getSheetByName('캐릭터_기본(Basic)');
  if (!sh) return [];
  const vals = sh.getRange(5, 1, 120, 25).getValues();
  const out = [];
  vals.forEach(function (r) {
    const name = String(r[0] || '').trim(); if (!name) return;
    const power = (typeof r[1] === 'number') ? r[1] : (parseInt(String(r[1]).replace(/[^0-9]/g, '')) || null);
    const eq = [];
    for (var i = 3; i < 24; i++) eq.push(r[i] === null || r[i] === undefined ? '' : String(r[i]));
    var lv = String(r[24] || '').trim(); const loc = (lv === '외근' || lv === '출장') ? '외근' : (lv === '아레나' ? '아레나' : '내근');
    out.push({ name: name, power: power, eq: eq, loc: loc });
  });
  return out;
}

function readSkills(ss) {
  const sh = ss.getSheetByName('캐릭터_스킬_마유(skill,MAU)');
  if (!sh) return [];
  const vals = sh.getRange(5, 1, 120, 32).getValues();
  const out = [];
  vals.forEach(function (r) {
    const name = String(r[0] || '').trim(); if (!name) return;
    const v = [];
    for (var i = 3; i < 31; i++) v.push(r[i] === null || r[i] === undefined ? '' : String(r[i]));
    var lv2 = String(r[31] || '').trim(); const loc = (lv2 === '외근' || lv2 === '출장') ? '외근' : (lv2 === '아레나' ? '아레나' : '내근');
    out.push({ name: name, vals: v, loc: loc });
  });
  return out;
}

function readAuction(ss) {
  const sh = ss.getSheetByName('경매 구매');
  const res = { mau: [] };
  if (!sh) return res;
  const vals = sh.getRange(1, 1, 200, 3).getValues();
  var on = false;
  for (var i = 0; i < vals.length; i++) {
    const a = String(vals[i][0] || '').trim();
    if (!on) { if (a === '캐릭명') on = true; continue; }
    if (!a || a.indexOf('경매 시작가') === 0) break;
    res.mau.push({ name: a, memo: String(vals[i][1] || '') });
  }
  return res;
}

// [드랍다운리스트] 탭 읽기 (헤더 3행, 데이터 4행부터 · A~F열)
function loadDropdowns() {
  try {
    const sh = getSS().getSheetByName('드랍다운리스트');
    if (!sh) return null;
    const vals = sh.getRange(4, 1, 150, 6).getValues();
    const pick = function (ci) {
      const arr = [];
      for (var i = 0; i < vals.length; i++) {
        const v = vals[i][ci];
        if (v !== null && v !== undefined && String(v).trim() !== '') arr.push(String(v).trim());
      }
      return arr;
    };
    const dd = { tier: pick(0), drive: pick(1), star: pick(2), defl: pick(3), 'class': pick(4), gang: pick(5) };
    if (!dd.tier.length) return null;
    return dd;
  } catch (err) { return null; }
}

function doGet(e) {
  // GitHub Pages에서 호출하는 읽기 API (화면은 GitHub가 서비스)
  return out({ ok: true, v: VERSION, data: loadData(), dd: loadDropdowns() });
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);

    // (1) 로그인: 비밀번호는 서버에만 존재 — HTML에는 노출되지 않음
    if (p.action === 'auth') {
      const role = (p.pw === PW) ? 'admin' : ((p.pw === PW_MEMBER) ? 'member' : null);
      if (!role) return out({ ok: false, v: VERSION, err: 'bad_pw' });
      return out({ ok: true, v: VERSION, role: role, token: makeToken_(role) });
    }

    // (2) 최신 데이터 조회 (저장 전 병합용) — 토큰 필요
    if (p.action === 'load') {
      if (!verifyToken_(p.token)) return out({ ok: false, v: VERSION, err: 'bad_token' });
      return out({ ok: true, v: VERSION, data: loadData() });
    }

    // (3) 저장 — 토큰 검증
    const role = verifyToken_(p.token);
    if (!role) return out({ ok: false, v: VERSION, err: 'bad_token' });
    const isAdmin = (role === 'admin'), isMember = (role === 'member');
    ensureSheets_();
    const ss = getSS();
    // 길드원 권한: 허용된 탭 데이터만 반영, 나머지는 서버 보관본 유지
    if (isMember) {
      const cur = loadData() || {};
      const perm = (cur.tabPerm) || { skill: 'member', mau: 'member' };
      const merged = JSON.parse(JSON.stringify(cur));
      if (perm.skill === 'member' || perm.mau === 'member') merged.skills = p.data.skills;
      if (perm.basic === 'member') merged.basic = p.data.basic;
      if (perm.notice === 'member') merged.notice = p.data.notice;
      if (perm.party === 'member') merged.parties = p.data.parties;
      if (perm.auction === 'member') { merged.mau = p.data.mau; merged.auctionPrices = p.data.auctionPrices; }
      if (perm.sheet === 'member') merged.freeSheet = p.data.freeSheet;
      p.data = merged;
    }
    let ws = ss.getSheetByName(KEY_SHEET) || ss.insertSheet(KEY_SHEET);
    ws.getRange(1, 1).setValue(JSON.stringify(p.data));
    ws.getRange(1, 2).setValue(new Date());
    const report = {
      basic: mirrorBasic(ss, p.data),
      skills: mirrorSkills(ss, p.data),
      notice: mirrorNotice(ss, p.data),
      auction: mirrorAuction(ss, p.data),
      parties: mirrorParties(ss, p.data)
    };
    return out({ ok: true, v: VERSION, mirrored: report });
  } catch (err) {
    return out({ ok: false, v: VERSION, err: String(err) });
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// ---- 캐릭터_기본(Basic): A=캐릭명, B=전투력, D~X=장비 (C 순위수식 보존) ----
function mirrorBasic(ss, d) {
  const sh = ss.getSheetByName('캐릭터_기본(Basic)');
  if (!sh) return 'tab_not_found';
  if (!d.basic || !d.basic.length) return 'no_data';
  const start = 5, clearRows = Math.max(d.basic.length, 60);
  sh.getRange(start, 1, clearRows, 2).clearContent();
  sh.getRange(start, 4, clearRows, 21).clearContent();
  sh.getRange(start, 1, d.basic.length, 2).setValues(d.basic.map(function (b) { return [b.name, b.power]; }));
  sh.getRange(start, 4, d.basic.length, 21).setValues(d.basic.map(function (b) { return (b.eq || []).concat(new Array(21).fill('')).slice(0, 21); }));
  sh.getRange(4, 25).setValue('위치');
  sh.getRange(start, 25, clearRows, 1).clearContent();
  sh.getRange(start, 25, d.basic.length, 1).setValues(d.basic.map(function (b) { return [b.loc || '내근']; }));
  return d.basic.length + 'rows';
}

// ---- 캐릭터_스킬_마유: A=캐릭명, D~AE=값 (B/C 수식 보존) ----
function mirrorSkills(ss, d) {
  const sh = ss.getSheetByName('캐릭터_스킬_마유(skill,MAU)');
  if (!sh) return 'tab_not_found';
  if (!d.skills || !d.skills.length) return 'no_data';
  const start = 5, clearRows = Math.max(d.skills.length, 60);
  sh.getRange(start, 1, clearRows, 1).clearContent();
  sh.getRange(start, 4, clearRows, 28).clearContent();
  sh.getRange(start, 1, d.skills.length, 1).setValues(d.skills.map(function (s) { return [s.name]; }));
  sh.getRange(start, 4, d.skills.length, 28).setValues(d.skills.map(function (s) { return (s.vals || []).concat(new Array(28).fill('')).slice(0, 28); }));
  sh.getRange(4, 32).setValue('위치');
  sh.getRange(start, 32, clearRows, 1).clearContent();
  sh.getRange(start, 32, d.skills.length, 1).setValues(d.skills.map(function (s) { return [s.loc || '내근']; }));
  return d.skills.length + 'rows';
}

// ---- 공지(Notice): 좌측 스케줄+규칙 / 우측 명단 / 메인캐 ----
function mirrorNotice(ss, d) {
  const sh = ss.getSheetByName('공지(Notice)');
  if (!sh) return 'tab_not_found';
  const n = d.notice;
  if (!n) return 'no_data';

  // 좌측 블록 (B~G): 병합 해제 후 값 재기록
  sh.getRange(2, 2, 60, 6).breakApart();
  sh.getRange(2, 2, 60, 6).clearContent();
  sh.getRange(2, 2).setValue('길드 컨텐츠 스케줄(전원 참여 필요) / 중립시 패널티 없음');
  sh.getRange(3, 2, 1, 5).setValues([['요일', '시간', '컨텐츠명', '공대 참여', '미참여 패널티']]);
  var sched = (n.schedule || []).map(function (r) { return r.length === 6 ? r.slice(1) : r; });
  // 요일순 정렬(월~일)
  var ORD = { '월': 0, '화': 1, '수': 2, '목': 3, '금': 4, '토': 5, '일': 6 };
  sched.sort(function (a, b) {
    var da = ORD[a[0]] !== undefined ? ORD[a[0]] : 9, db = ORD[b[0]] !== undefined ? ORD[b[0]] : 9;
    return da - db || String(a[1]).localeCompare(String(b[1]));
  });
  if (sched.length) sh.getRange(4, 2, sched.length, 5).setValues(sched);
  var rr = 4 + sched.length + 1;
  sh.getRange(rr, 2).setValue(n.rules || '');
  sh.getRange(rr, 2).setWrap(true);

  // 우측 명단 (I~L)
  sh.getRange(2, 9, 90, 6).breakApart();
  sh.getRange(2, 9, 90, 6).clearContent();
  sh.getRange(2, 9, 1, 6).setValues([['캐릭터', '구분', '위치', '단톡방', '공지방', '비고']]);
  if (n.roster && n.roster.length)
    sh.getRange(3, 9, n.roster.length, 6).setValues(n.roster.map(function (r) {
      return [r[0], r[4] || '본캐', r[5] || '내근', r[1], r[2], r[3] || ''];
    }));

  return 'ok';
}

// ---- 경매: MAU 진행현황 + 시작가 ----
function mirrorAuction(ss, d) {
  const sh = ss.getSheetByName('경매 구매');
  if (!sh) return 'tab_not_found';
  sh.getRange(1, 1, 200, 6).breakApart();
  sh.getRange(1, 1, 200, 6).clearContent();
  var r = 1;
  sh.getRange(r, 1).setValue('MAU(오딘) / 런처 진행 현황').setFontWeight('bold'); r += 1;
  sh.getRange(r, 1, 1, 2).setValues([['캐릭명', '진행 현황']]).setFontWeight('bold'); r += 1;
  if (d.mau && d.mau.length) {
    sh.getRange(r, 1, d.mau.length, 2).setValues(d.mau.map(function (m) { return [m.name, m.memo || '']; }));
    r += d.mau.length;
  }
  r += 1;
  sh.getRange(r, 1).setValue('경매 시작가').setFontWeight('bold'); r += 1;
  sh.getRange(r, 1, 1, 2).setValues([['항목', '시작가 / 방식']]).setFontWeight('bold'); r += 1;
  if (d.auctionPrices && d.auctionPrices.length) {
    sh.getRange(r, 1, d.auctionPrices.length, 2)
      .setValues(d.auctionPrices.map(function (p) { return [p[0] || '', p[1] || '']; }));
  }
  return 'ok';
}


// ============================================================
// [1회 실행용] 장비 표기 일괄 변환: 4T_프_15 → 4TP+15, 4T_3 → 4T+3
// 사용법: 편집기 상단 함수 선택에서 convertFormatsOnce 선택 → ▶ 실행
// ============================================================
function convertFormatsOnce() {
  const ss = getSS();
  const cvt = function (v) {
    if (v === null || v === undefined || v === '') return v;
    return String(v).replace(/_프_(\d+)$/, 'P+$1').replace(/_(\d+)$/, '+$1');
  };
  // 1) 드랍다운리스트 A~D열 (장비등급/드라이브/스타게이저/디플렉터)
  const dd = ss.getSheetByName('드랍다운리스트');
  if (dd) {
    const rng = dd.getRange(4, 1, 120, 4);
    const vals = rng.getValues().map(function (row) { return row.map(cvt); });
    rng.setValues(vals);
  }
  // 2) 캐릭터_기본 장비 데이터 (D5~X)
  const cb = ss.getSheetByName('캐릭터_기본(Basic)');
  if (cb) {
    const rng = cb.getRange(5, 4, 80, 21);
    const vals = rng.getValues().map(function (row) { return row.map(cvt); });
    rng.setValues(vals);
  }
  // 3) WEB_DATA JSON 내부 장비값
  const ws = ss.getSheetByName(KEY_SHEET);
  if (ws) {
    const raw = ws.getRange(1, 1).getValue();
    if (raw) {
      try {
        const d = JSON.parse(raw);
        if (d.basic) d.basic.forEach(function (b) { if (b.eq) b.eq = b.eq.map(cvt); });
        ws.getRange(1, 1).setValue(JSON.stringify(d));
      } catch (err) {}
    }
  }
  Logger.log('변환 완료');
}


// ---- 공대 편성: [공대편성] 시트에 표 형태로 기록 (없으면 생성) ----
function mirrorParties(ss, d) {
  if (!d.parties) return 'no_data';
  var sh = ss.getSheetByName('공대편성');
  if (!sh) sh = ss.insertSheet('공대편성');
  sh.clear();
  var r = 1;
  Object.keys(d.parties).forEach(function (key) {
    var g = d.parties[key];
    var teams = g.teams || [];
    sh.getRange(r, 1).setValue(key + ' 공대 편성').setFontWeight('bold');
    r += 1;
    sh.getRange(r, 1).setValue('역할');
    sh.getRange(r, 2, 1, teams.length).setValues([teams.map(function (t, i) { return (g.roles || [])[i] || ''; })]);
    r += 1;
    sh.getRange(r, 1).setValue('공대');
    sh.getRange(r, 2, 1, teams.length).setValues([teams.map(function (t) { return t.name; })]);
    sh.getRange(r, 1, 1, teams.length + 1).setFontWeight('bold');
    r += 1;
    var maxRows = 1;
    teams.forEach(function (t) { maxRows = Math.max(maxRows, (t.members || []).length); });
    var grid = [];
    for (var i = 0; i < maxRows; i++) {
      grid.push([''].concat(teams.map(function (t) { return (t.members || [])[i] || ''; })));
    }
    if (grid.length) { sh.getRange(r, 1, grid.length, teams.length + 1).setValues(grid); r += grid.length; }
    r += 1;
    // 역할별 플레이 방법
    var notes = g.notes || {};
    var uniq = [];
    (g.roles || []).slice(0, teams.length).forEach(function (x) { var l = x || ''; if (uniq.indexOf(l) < 0) uniq.push(l); });
    if (uniq.length) {
      sh.getRange(r, 1).setValue('플레이 방법').setFontWeight('bold');
      r += 1;
      var nrows = uniq.map(function (l) { return [l, notes[l] || '']; });
      sh.getRange(r, 1, nrows.length, 2).setValues(nrows);
      sh.getRange(r, 2, nrows.length, 1).setWrap(true);
      r += nrows.length + 1;
    }
  });
  sh.autoResizeColumns(1, 8);
  return 'ok';
}


// ============================================================
// 빈 구글 시트 자동 구성: 포털이 읽고 쓰는 탭이 없으면 머리글과 함께 만든다
// (저장할 때마다 자동 확인 · 처음 한 번은 편집기에서 setupSheetsOnce ▶ 실행 권장)
// ============================================================
const BASIC_HEADERS = ["본/슈트","본/무기","부/슈트","부/무기","목","가슴","팔","하체","다리","서클릿","귀걸이","목걸이","팔찌","반지","디멘션큐브","사이오닉","하이퍼","퀸텀","드라이브","스타게이저","디플렉터"];
const SKILL_HEADERS = ["난공불락","삼위일체","기달","정면승부","피의광기","4스킬","5스킬","6스킬","7스킬","8스킬","9스킬","무달","동체시력","6스킬(컬)","삼위일체(컬)","기달(컬)","유기체","네더","어비스","크림슨","템페","헤라","오딘","토르","타이폰","케이론PC","케이론IC","카드모스PC"];
const DD_DEFAULT = {
  tier: ["1T+5","1T+6","1T+7","1T+8","1T+9","1T+10","1T+11","1T+12","1T+13","1T+14","1T+15","1TP+5","1TP+6","1TP+7","1TP+8","1TP+9","1TP+10","1TP+11","1TP+12","1TP+13","1TP+14","1TP+15","2T+5","2T+6","2T+7","2T+8","2T+9","2T+10","2T+11","2T+12","2T+13","2T+14","2T+15","2TP+5","2TP+6","2TP+7","2TP+8","2TP+9","2TP+10","2TP+11","2TP+12","2TP+13","2TP+14","2TP+15","3T+5","3T+6","3T+7","3T+8","3T+9","3T+10","3T+11","3T+12","3T+13","3T+14","3T+15","3TP+5","3TP+6","3TP+7","3TP+8","3TP+9","3TP+10","3TP+11","3TP+12","3TP+13","3TP+14","3TP+15","4T+5","4T+6","4T+7","4T+8","4T+9","4T+10","4T+11","4T+12","4T+13","4T+14","4T+15","4TP+5","4TP+6","4TP+7","4TP+8","4TP+9","4TP+10","4TP+11","4TP+12","4TP+13","4TP+14","4TP+15","5T+5","5T+6","5T+7","5T+8","5T+9","5T+10","5T+11","5T+12","5T+13","5T+14","5T+15","5TP+5","5TP+6","5TP+7","5TP+8","5TP+9","5TP+10","5TP+11","5TP+12","5TP+13","5TP+14","5TP+15"],
  drive: ["반중+3","반중+4","반중+5","반중+6","반중+7","반중+8","반중+9","반중+10","반중+11","반중+12","반중+13","반중+14","반중+15","반중P+3","반중P+4","반중P+5","반중P+6","반중P+7","반중P+8","반중P+9","반중P+10","반중P+11","반중P+12","반중P+13","반중P+14","반중P+15","아스+3","아스+4","아스+5","아스+6","아스+7","아스+8","아스+9","아스+10","아스+11","아스+12","아스+13","아스+14","아스+15","아스P+3","아스P+4","아스P+5","아스P+6","아스P+7","아스P+8","아스P+9","아스P+10","아스P+11","아스P+12","아스P+13","아스P+14","아스P+15","보이드+3","보이드+4","보이드+5","보이드+6","보이드+7","보이드+8","보이드+9","보이드+10","보이드+11","보이드+12","보이드+13","보이드+14","보이드+15","보이드P+3","보이드P+4","보이드P+5","보이드P+6","보이드P+7","보이드P+8","보이드P+9","보이드P+10","보이드P+11","보이드P+12","보이드P+13","보이드P+14","보이드P+15","로켓+3","로켓+4","로켓+5","로켓+6","로켓+7","로켓+8","로켓+9","로켓+10","로켓+11","로켓+12","로켓+13","로켓+14","로켓+15","로켓P+3","로켓P+4","로켓P+5","로켓P+6","로켓P+7","로켓P+8","로켓P+9","로켓P+10","로켓P+11","로켓P+12","로켓P+13","로켓P+14","로켓P+15"],
  star: ["탐사+3","탐사+4","탐사+5","탐사+6","탐사+7","탐사+8","탐사+9","탐사+10","탐사+11","탐사+12","탐사+13","탐사+14","탐사+15","탐사P+3","탐사P+4","탐사P+5","탐사P+6","탐사P+7","탐사P+8","탐사P+9","탐사P+10","탐사P+11","탐사P+12","탐사P+13","탐사P+14","탐사P+15","심우주+3","심우주+4","심우주+5","심우주+6","심우주+7","심우주+8","심우주+9","심우주+10","심우주+11","심우주+12","심우주+13","심우주+14","심우주+15","심우주P+3","심우주P+4","심우주P+5","심우주P+6","심우주P+7","심우주P+8","심우주P+9","심우주P+10","심우주P+11","심우주P+12","심우주P+13","심우주P+14","심우주P+15","보급+3","보급+4","보급+5","보급+6","보급+7","보급+8","보급+9","보급+10","보급+11","보급+12","보급+13","보급+14","보급+15","보급P+3","보급P+4","보급P+5","보급P+6","보급P+7","보급P+8","보급P+9","보급P+10","보급P+11","보급P+12","보급P+13","보급P+14","보급P+15"],
  defl: ["표준+3","표준+4","표준+5","표준+6","표준+7","표준+8","표준+9","표준+10","표준+11","표준+12","표준+13","표준+14","표준+15","표준P+3","표준P+4","표준P+5","표준P+6","표준P+7","표준P+8","표준P+9","표준P+10","표준P+11","표준P+12","표준P+13","표준P+14","표준P+15","강화+3","강화+4","강화+5","강화+6","강화+7","강화+8","강화+9","강화+10","강화+11","강화+12","강화+13","강화+14","강화+15","강화P+3","강화P+4","강화P+5","강화P+6","강화P+7","강화P+8","강화P+9","강화P+10","강화P+11","강화P+12","강화P+13","강화P+14","강화P+15"],
  'class': ["퍼니셔","팬텀","인포","사이퍼","드레드","테크","아비","데몰"],
  gang: ["1강","2강","3강","4강","5강","6강","7강","8강","9강","10강","11강","12강","13강","14강","15강"]
};

function ensureSheets_() {
  const ss = getSS();
  var sh = ss.getSheetByName('캐릭터_기본(Basic)');
  if (!sh) {
    sh = ss.insertSheet('캐릭터_기본(Basic)');
    sh.getRange(4, 1, 1, 25).setValues([['캐릭명', '전투력', '순위'].concat(BASIC_HEADERS, ['위치'])]).setFontWeight('bold');
    const f = [];
    for (var r = 5; r <= 124; r++) f.push(['=IF(A' + r + '="","",IFERROR(RANK(B' + r + ',$B$5:$B$124),""))']);
    sh.getRange(5, 3, 120, 1).setFormulas(f);
    sh.setFrozenRows(4);
  }
  sh = ss.getSheetByName('캐릭터_스킬_마유(skill,MAU)');
  if (!sh) {
    sh = ss.insertSheet('캐릭터_스킬_마유(skill,MAU)');
    sh.getRange(4, 1, 1, 32).setValues([['캐릭명', '', ''].concat(SKILL_HEADERS, ['위치'])]).setFontWeight('bold');
    sh.setFrozenRows(4);
  }
  if (!ss.getSheetByName('공지(Notice)')) ss.insertSheet('공지(Notice)');
  if (!ss.getSheetByName('경매 구매')) ss.insertSheet('경매 구매');
  if (!ss.getSheetByName(KEY_SHEET)) ss.insertSheet(KEY_SHEET);
  sh = ss.getSheetByName('드랍다운리스트');
  if (!sh) {
    sh = ss.insertSheet('드랍다운리스트');
    const keys = ['tier', 'drive', 'star', 'defl', 'class', 'gang'];
    sh.getRange(3, 1, 1, 6).setValues([['장비등급', '드라이브', '스타게이저', '디플렉터', '클래스', '강화']]).setFontWeight('bold');
    const n = Math.max.apply(null, keys.map(function (k) { return DD_DEFAULT[k].length; }));
    const rows = [];
    for (var i = 0; i < n; i++) rows.push(keys.map(function (k) { return DD_DEFAULT[k][i] || ''; }));
    sh.getRange(4, 1, n, 6).setValues(rows);
  }
}

function setupSheetsOnce() {
  ensureSheets_();
  Logger.log('시트 구성 완료: ' + getSS().getSheets().map(function (s) { return s.getName(); }).join(', '));
}
