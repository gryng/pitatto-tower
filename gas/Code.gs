/**
 * ピタッとタワー オンラインランキング（Google Apps Script）
 *
 * 使い方
 *   1. 新しいGoogleスプレッドシートを作り、「拡張機能 → Apps Script」を開く
 *   2. このファイルの中身をまるごと貼り付けて保存
 *   3. 「デプロイ → 新しいデプロイ」で種類「ウェブアプリ」を選び、
 *      実行ユーザー「自分」、アクセスできるユーザー「全員」でデプロイ
 *   4. 発行された「ウェブアプリのURL（…/exec）」をゲームの RANK_URL に設定
 *
 * 記録はスプレッドシートの「scores」シートに1プレイ1行で追記される。
 * ランキングは名前ごとの最高スコアで並べる（名前の大文字小文字は区別しない）。
 */
const SHEET_NAME = 'scores';
const TOP_N = 10;

/** ランキングを返す */
function doGet() {
  return json_(buildRanking_(null));
}

/** スコアを登録して、登録後のランキングと順位を返す */
function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad_json' });
  }
  const entry = validate_(data);
  if (!entry) return json_({ ok: false, error: 'invalid' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    getSheet_().appendRow([new Date(), entry.name, entry.points, entry.floors, entry.perfects, entry.maxCombo, entry.version]);
  } finally {
    lock.releaseLock();
  }
  return json_(buildRanking_(entry));
}

/** 送られてきた値の形と、ゲームのルール上ありえる範囲かを確認する */
function validate_(d) {
  if (!d || typeof d !== 'object') return null;
  const name = String(d.name || '').replace(/[\u0000-\u001f\u007f]/g, '').trim().slice(0, 12);
  const points = Math.floor(Number(d.points));
  const floors = Math.floor(Number(d.floors));
  const perfects = Math.floor(Number(d.perfects));
  const maxCombo = Math.floor(Number(d.maxCombo));
  if (!name) return null;
  if (![points, floors, perfects, maxCombo].every(n => Number.isFinite(n) && n >= 0)) return null;
  if (floors > 2000 || perfects > floors || maxCombo > perfects) return null;
  if (points > maxPoints_(floors)) return null;
  return { name, points, floors, perfects, maxCombo, version: String(d.version || '').slice(0, 16) };
}

/**
 * その段数で理論上とれる最高点（全段PERFECT・倍率最大）。
 * ゲーム側の配点（PERFECT 150点 × コンボ倍率最大2 × 高さ倍率、10段ごとに 500×n のボーナス）と合わせること。
 */
function maxPoints_(floors) {
  let max = 0;
  for (let f = 0; f < floors; f++) max += Math.round(150 * 2 * (1 + Math.floor(f / 10) * 0.1));
  for (let k = 1; k * 10 <= floors; k++) max += 500 * k;
  return max;
}

function buildRanking_(entry) {
  const rows = getSheet_().getDataRange().getValues().slice(1);
  const best = {};
  rows.forEach(r => {
    const name = String(r[1]);
    const key = name.toLowerCase();
    const points = Number(r[2]) || 0;
    if (!best[key] || points > best[key].points) best[key] = { name, points, floors: Number(r[3]) || 0 };
  });
  const list = Object.keys(best).map(k => best[k]).sort((a, b) => b.points - a.points);
  const res = { ok: true, top: list.slice(0, TOP_N), players: list.length };
  if (entry) res.rank = list.filter(p => p.points > entry.points).length + 1;
  return res;
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['日時', '名前', 'スコア', '段数', 'PERFECT', '最大コンボ', 'バージョン']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
