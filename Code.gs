// 動感節拍組合器 － 雲端帳號後端 (Google Apps Script)
// 這份程式放在 GitHub，Apps Script 裡的「啟動程式」會自動抓最新版來執行（不要在這裡寫試算表 ID 和密碼：GitHub 是公開的）。
// 每次修改這份程式，下面的 CODE_VERSION 都要 +1；推上 GitHub 後，在遊戲「老師模式」按「更新伺服器程式」就會生效。
const CODE_VERSION = 1;
// 學生積分存在一份 Google 試算表：Users(每人一列) / Log(每次紀錄) / 排行榜(積分自動排序) / 成績(遊戲內排行榜：每人每種模式只留最佳成績)

const USERS = 'Users';
const LOG = 'Log';
const BOARD = '排行榜';
const SCORES = '成績';
const SCORE_MEASURES = [6, 12, 18, 24, 30], SCORE_TIERS = ['easy', 'mid', 'hard'];

// ★ 老師密碼存在 Apps Script 的「指令碼屬性」TEACHER_PIN（空白 = 老師功能關閉）；可在遊戲老師模式改密碼。
function pin_() { return String(PropertiesService.getScriptProperties().getProperty('TEACHER_PIN') || ''); }
// 班級挑戰：全班累計 Perfect 目標、最少需要幾位同學（名字要寫成「3年2班-小明」，「-」前面就是班級／小組）
const CLASS_GOAL = 10000, CLASS_MIN_MEMBERS = 3;

// 取得試算表：先用指令碼屬性 SHEET_ID；沒有的話專案若綁在試算表就用那份；再不然第一次自動建立一份並記住它
let SS_MEMO = null;
function ss_() {
  if (SS_MEMO) return SS_MEMO;
  const props = PropertiesService.getScriptProperties();
  const sid = props.getProperty('SHEET_ID');
  if (sid) return (SS_MEMO = SpreadsheetApp.openById(sid));
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return (SS_MEMO = ss);
  const id = props.getProperty('SS_ID');
  if (id) return (SS_MEMO = SpreadsheetApp.openById(id));
  ss = SpreadsheetApp.create('動感節拍組合器 - 學生積分');
  props.setProperty('SS_ID', ss.getId());
  return (SS_MEMO = ss);
}

function sheet_(name, header) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#e0f2fe');
  }
  return sh;
}
const usersSheet_ = () => sheet_(USERS, ['姓名', '目前積分', '累計獲得', '獎章', '最後更新']);
const logSheet_ = () => sheet_(LOG, ['時間', '姓名', '模式', '本局分數', '最高連擊', 'Perfect數', '目前積分']);

const scoresSheet_ = () => sheet_(SCORES, ['時間', '姓名', '小節數', '難度', 'BPM', '主題', '成績%', 'Perfect數', '最高連擊', 'key']);

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function findRow_(sh, name) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const names = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < names.length; i++) if (String(names[i][0]) === name) return i + 2;
  return -1;
}

// 班級／小組＝名字裡第一個「-」「－」「_」前面的文字，例如「3年2班-小明」→「3年2班」
function groupOf_(name) {
  const i = String(name).search(/[-－‐_]/);
  if (i < 1) return '';
  const g = name.slice(0, i).trim(), r = name.slice(i + 1).trim();
  return g && r && g.length <= 12 ? g : '';
}
// 讀出全部學生（含統計）。w = 本週（週一日期，由遊戲傳來）
function allUsers_(w) {
  const sh = usersSheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 5).getValues().map(v => {
    let m = {}; try { m = JSON.parse(v[3] || '{}'); } catch (err) {}
    const st = m._st || {}, wp = m._wp || {}, cur = wp.w === w;
    return { name: String(v[0]), group: groupOf_(String(v[0])), points: Number(v[1]) || 0, lifetime: Number(v[2]) || 0,
      games: Number(st.games) || 0, perf: Number(st.perfTotal) || 0, wp: cur ? Number(wp.perf) || 0 : 0, wg: cur ? Number(wp.games) || 0 : 0,
      last: +new Date(v[4]) || 0, lv: st.lv || {}, nt: st.nt || {} };
  }).filter(r => r.name);
}
// 各班彙總（快取 60 秒，避免全班同時開啟時一直重算）
function groups_(w) {
  const cache = CacheService.getScriptCache(), key = 'grp:' + w, hit = cache.get(key);
  if (hit) { try { return JSON.parse(hit); } catch (err) {} }
  const g = {};
  allUsers_(w).forEach(r => {
    if (!r.group) return;
    const o = g[r.group] || (g[r.group] = { g: r.group, n: 0, perf: 0, wp: 0 });
    o.n++; o.perf += r.perf; o.wp += r.wp;
  });
  const out = Object.keys(g).map(k => g[k]);
  try { cache.put(key, JSON.stringify(out), 60); } catch (err) {}
  return out;
}

// ★ 第一次請在編輯器手動執行這個函式一次：授權 + 建好三個分頁，並在「執行記錄」印出試算表網址
function setup() {
  usersSheet_(); logSheet_(); scoresSheet_();
  const ss = ss_();
  let b = ss.getSheetByName(BOARD);
  if (!b) {
    b = ss.insertSheet(BOARD);
    b.getRange(1, 1, 1, 5).setValues([['姓名', '目前積分', '累計獲得', '獎章', '最後更新']]).setFontWeight('bold').setBackground('#fef3c7');
    b.getRange('A2').setFormula('=IFERROR(SORT(FILTER(Users!A2:E, Users!A2:A<>""), 3, FALSE), "")');
    b.setFrozenRows(1);
  }
  const old = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (old && old.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(old);
  console.log('試算表網址：' + ss.getUrl());
}

// 讀取帳號：?action=get&name=王小明
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'get') {
    const name = String(p.name || '').trim();
    const sh = usersSheet_();
    const row = findRow_(sh, name);
    if (row < 0) return json_({ ok: true, user: null });
    const v = sh.getRange(row, 1, 1, 4).getValues()[0];
    let medals = {}; try { medals = JSON.parse(v[3] || '{}'); } catch (err) {}
    return json_({ ok: true, user: { points: Number(v[1]) || 0, lifetime: Number(v[2]) || 0, medals: medals } });
  }
  // 遊戲內排行榜：?action=board&measures=6&tier=easy  → 該組合的成績，由高到低（同分先達成者在前）
  if (p.action === 'board') {
    const m = Number(p.measures), t = String(p.tier || '');
    if (SCORE_MEASURES.indexOf(m) < 0 || SCORE_TIERS.indexOf(t) < 0) return json_({ ok: false, error: 'bad params' });
    const sh = scoresSheet_(), last = sh.getLastRow(), key = m + '-' + t, rows = [];
    if (last >= 2) {
      sh.getRange(2, 1, last - 1, 10).getValues().forEach(v => {
        if (String(v[9]) !== key) return;
        rows.push({ name: String(v[1]), pct: Number(v[6]) || 0, bpm: Number(v[4]) || 0, theme: String(v[5]), perfects: Number(v[7]) || 0, combo: Number(v[8]) || 0, t: +new Date(v[0]) || 0 });
      });
    }
    rows.sort((a, b) => (b.pct - a.pct) || (a.t - b.t));
    return json_({ ok: true, lb: true, rows: rows.slice(0, 50) });
  }
  // 班級挑戰＋組隊週賽：?action=class&name=3年2班-小明&w=2026-10-05
  if (p.action === 'class') {
    const name = String(p.name || '').trim(), w = String(p.w || '').slice(0, 10), grp = groupOf_(name), all = groups_(w);
    const mine = all.filter(x => x.g === grp)[0] || null;
    const teams = all.filter(x => x.n >= CLASS_MIN_MEMBERS).sort((a, b) => (b.wp - a.wp) || (b.perf - a.perf)).slice(0, 10);
    return json_({ ok: true, cls: true, group: grp, members: mine ? mine.n : 0, total: mine ? mine.perf : 0, weekPerf: mine ? mine.wp : 0,
      goal: CLASS_GOAL, minMembers: CLASS_MIN_MEMBERS, teams: teams });
  }
  // 段位榜／積分榜：?action=ranks&name=小明&g=3年2班（g 可省略＝全部）。比的是「累計獲得」，花掉積分不會掉名次
  if (p.action === 'ranks') {
    const name = String(p.name || '').trim(), g = String(p.g || '').slice(0, 12), cache = CacheService.getScriptCache();
    let all = null; const hit = cache.get('ranks');
    if (hit) { try { all = JSON.parse(hit); } catch (err) {} }
    if (!all) {
      const sh = usersSheet_(), last = sh.getLastRow();
      all = last < 2 ? [] : sh.getRange(2, 1, last - 1, 3).getValues().map(v => [String(v[0]), Number(v[2]) || 0]).filter(x => x[0] && x[1] > 0);
      try { cache.put('ranks', JSON.stringify(all), 60); } catch (err) {}
    }
    const list = (g ? all.filter(x => groupOf_(x[0]) === g) : all.slice()).sort((a, b) => (b[1] - a[1]) || (a[0] < b[0] ? -1 : 1));
    const idx = list.findIndex(x => x[0] === name);
    return json_({ ok: true, ranks: true, group: g, total: list.length, rows: list.slice(0, 50).map(x => ({ name: x[0], lt: x[1] })),
      lts: list.slice(0, 2000).map(x => x[1]), me: idx >= 0 ? { rank: idx + 1, lt: list[idx][1] } : null });
  }
  // 老師儀表板：?action=teacher&pin=密碼&w=2026-10-05
  if (p.action === 'teacher') {
    if (!pin_()) return json_({ ok: false, error: 'teacher disabled' });
    if (String(p.pin || '') !== pin_()) return json_({ ok: false, error: 'bad pin' });
    return json_({ ok: true, teacher: true, week: String(p.w || ''), goal: CLASS_GOAL, rows: allUsers_(String(p.w || '').slice(0, 10)) });
  }
  // 測試連線：?action=ping → 伺服器程式版本與試算表名稱
  if (p.action === 'ping') {
    let sheet = ''; try { sheet = ss_().getName(); } catch (err) { sheet = '(讀不到試算表：' + String(err) + ')'; }
    return json_({ ok: true, version: CODE_VERSION, sheet: sheet });
  }
  return json_({ ok: true, msg: 'rhythm game api running' });
}

// 存檔：POST { action:'save', username, totalPoints, lifetime, medals, mode?, score?, combo?, perfects? }
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const name = String(d.username || '').trim();
    if (d.action === 'score') return json_(saveScore_(d, name));
    // 老師動作（要密碼；密碼還沒設定時一律拒絕）
    if (d.action === 'selfUpdate' || d.action === 'setPin') {
      const cur = pin_();
      if (!cur) return json_({ ok: false, error: 'teacher disabled' });
      if (String(d.pin || '') !== cur) return json_({ ok: false, error: 'bad pin' });
      if (d.action === 'setPin') {
        const np = String(d.newPin || '');
        if (np.length < 4 || np.length > 40) return json_({ ok: false, error: 'bad new pin' });
        PropertiesService.getScriptProperties().setProperty('TEACHER_PIN', np);
        return json_({ ok: true, msg: '密碼已更改' });
      }
      if (typeof LOADER === 'undefined') return json_({ ok: false, error: 'no loader' });
      const app = LOADER.reload(), upd = app.CODE_VERSION !== CODE_VERSION;
      return json_({ ok: true, updated: upd, version: app.CODE_VERSION, msg: upd ? '已更新到第 ' + app.CODE_VERSION + ' 版' : '已經是最新版（第 ' + app.CODE_VERSION + ' 版）' });
    }
    if (d.action !== 'save' || !name) return json_({ ok: false, error: 'bad request' });
    const sh = usersSheet_();
    const rowData = [name, Number(d.totalPoints) || 0, Number(d.lifetime) || 0, JSON.stringify(d.medals || {}), new Date()];
    const row = findRow_(sh, name);
    if (row < 0) sh.appendRow(rowData); else sh.getRange(row, 1, 1, 5).setValues([rowData]);
    if (d.mode) logSheet_().appendRow([new Date(), name, d.mode, d.score || 0, d.combo || 0, d.perfects || 0, Number(d.totalPoints) || 0]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 排行榜成績：POST { action:'score', username, measures, tier, bpm, theme, pct, perfects, combo }
// 同一個人、同一種(小節數+難度)只留最佳成績；比舊的差就不更新
function saveScore_(d, name) {
  const m = Number(d.measures), t = String(d.tier || ''), pct = Math.round(Number(d.pct) * 10) / 10, bpm = Number(d.bpm);
  if (!name || name.length > 30) return { ok: false, error: 'bad score name' };
  if (SCORE_MEASURES.indexOf(m) < 0 || SCORE_TIERS.indexOf(t) < 0 || !(pct > 0 && pct <= 100) || !(bpm >= 40 && bpm <= 200)) return { ok: false, error: 'bad score values' };
  const sh = scoresSheet_(), key = m + '-' + t, last = sh.getLastRow();
  const rowData = [new Date(), name, m, t, bpm, String(d.theme || '').slice(0, 20), pct, Math.max(0, Number(d.perfects) || 0), Math.max(0, Number(d.combo) || 0), key];
  if (last >= 2) {
    const vals = sh.getRange(2, 1, last - 1, 10).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][1]) === name && String(vals[i][9]) === key) {
        if (pct <= Number(vals[i][6])) return { ok: true, best: false };
        sh.getRange(i + 2, 1, 1, 10).setValues([rowData]);
        return { ok: true, best: true };
      }
    }
  }
  sh.appendRow(rowData);
  return { ok: true, best: true };
}
