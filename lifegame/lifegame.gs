/**
 * ライフゲーム（Googleスプレッドシート用 Apps Script）
 *
 * - 盤面（「ライフゲーム」シート B6:AE35）の塗りつぶしを読み取り、次世代はスクリプト内で計算する。
 * - 3行目の「入力」プルダウンで入力方法を選ぶ：
 *     なし（塗りつぶし） … 塗りつぶしツールで自由に描く
 *     ペン（クリックで反転）… クリックしたセルの生死を切り替える
 *     パターン名        … クリックした位置（左上）にパターンを置く。「向き」で回転
 * - 「全消去」チェックボックスで盤面を消す。
 * - 定型パターンは「パターン」シートで追加・削除できる（O=生、.=死、/=改行）。
 *   盤面で範囲を選択 → メニュー「選択範囲をパターンに登録」でも追加できる。
 *
 * 導入: 拡張機能 → Apps Script → このコードを貼り付けて保存 → シートを再読み込み
 *       → メニュー「ライフゲーム」が表示される（初回は権限の承認が必要）
 */
const BOARD = 'ライフゲーム';
const PATTERNS = 'パターン';
const N = 30;
const BOARD_TOP = 6, BOARD_LEFT = 2;            // 盤面 B6:AE35
const GEN = 'E4', POP = 'M4', STATUS = 'T4', SPEED = 'AD4';
const MODE = 'E3', ROTATE = 'R3', CLEAR = 'Z3';
const MODE_PAINT = 'なし（塗りつぶし）';
const MODE_PEN = 'ペン（クリックで反転）';
const DEAD = '#ffffff', ALIVE = '#1f2937';
const TIME_LIMIT_MS = 5.5 * 60 * 1000;          // GAS の実行上限(6分)より手前で止める

function onOpen() {
  SpreadsheetApp.getUi().createMenu('ライフゲーム')
    .addItem('▶ 再生', 'play')
    .addItem('■ 停止', 'stop')
    .addItem('1世代進める', 'stepOnce')
    .addSeparator()
    .addItem('ランダム配置', 'randomize')
    .addItem('全消去', 'clearAll')
    .addItem('世代を0に戻す', 'resetGen')
    .addSeparator()
    .addItem('選択範囲をパターンに登録', 'registerPattern')
    .addToUi();
}

/* ---------------- 共通 ---------------- */

function board_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(BOARD);
  return { sheet: sheet, grid: sheet.getRange(BOARD_TOP, BOARD_LEFT, N, N) };
}

const isAlive_ = (c) => !!c && c.toLowerCase() !== DEAD;

function writeInfo_(sheet, gen, pop) {
  sheet.getRange(GEN).setValue(gen);
  sheet.getRange(POP).setValue(pop);
}

function inBoard_(row, col) {
  return row >= BOARD_TOP && row < BOARD_TOP + N && col >= BOARD_LEFT && col < BOARD_LEFT + N;
}

/* ---------------- 計算・再生 ---------------- */

/** 色の盤面から次世代の色の盤面を作る。生き残りは元の色、誕生は ALIVE。盤外は死。 */
function nextColors_(colors) {
  const alive = colors.map(r => r.map(c => (isAlive_(c) ? 1 : 0)));
  const out = [];
  let pop = 0, changed = false;
  for (let i = 0; i < N; i++) {
    const row = [];
    for (let j = 0; j < N; j++) {
      let n = 0;
      for (let di = -1; di <= 1; di++) {
        const r = alive[i + di];
        if (!r) continue;
        for (let dj = -1; dj <= 1; dj++) {
          if ((di || dj) && r[j + dj]) n++;
        }
      }
      const me = alive[i][j];
      const live = n === 3 || (me && n === 2);
      if (live !== !!me) changed = true;
      if (live) pop++;
      row.push(live ? (me ? colors[i][j] : ALIVE) : DEAD);
    }
    out.push(row);
  }
  return { colors: out, pop: pop, changed: changed };
}

function play() {
  const b = board_();
  const status = b.sheet.getRange(STATUS);
  const speed = b.sheet.getRange(SPEED);
  status.setValue('再生中');
  let colors = b.grid.getBackgrounds();
  let gen = Number(b.sheet.getRange(GEN).getValue()) || 0;
  const start = Date.now();
  while (true) {
    if (status.getValue() !== '再生中') return;
    if (Date.now() - start > TIME_LIMIT_MS) {
      status.setValue('停止中（時間切れ：もう一度再生）');
      return;
    }
    const res = nextColors_(colors);
    if (!res.changed) {
      status.setValue('停止中（変化なし）');
      return;
    }
    colors = res.colors;
    b.grid.setBackgrounds(colors);
    writeInfo_(b.sheet, ++gen, res.pop);
    SpreadsheetApp.flush();
    const ms = Number(speed.getValue());
    if (ms > 0) Utilities.sleep(Math.min(ms, 5000));
  }
}

function stop() {
  board_().sheet.getRange(STATUS).setValue('停止中');
}

function stepOnce() {
  const b = board_();
  const res = nextColors_(b.grid.getBackgrounds());
  b.grid.setBackgrounds(res.colors);
  writeInfo_(b.sheet, (Number(b.sheet.getRange(GEN).getValue()) || 0) + 1, res.pop);
}

function randomize() {
  const b = board_();
  let pop = 0;
  const out = [];
  for (let i = 0; i < N; i++) {
    out.push([]);
    for (let j = 0; j < N; j++) {
      const live = Math.random() < 0.25;
      if (live) pop++;
      out[i].push(live ? ALIVE : DEAD);
    }
  }
  b.grid.setBackgrounds(out);
  writeInfo_(b.sheet, 0, pop);
}

function clearAll() {
  const b = board_();
  b.sheet.getRange(STATUS).setValue('停止中');   // 再生中なら止める
  b.grid.setBackground(DEAD);
  writeInfo_(b.sheet, 0, 0);
}

function resetGen() {
  board_().sheet.getRange(GEN).setValue(0);
}

/* ---------------- 入力（シンプルトリガー） ---------------- */

/** 「全消去」チェックボックス */
function onEdit(e) {
  const r = e.range;
  if (r.getSheet().getName() !== BOARD || r.getA1Notation() !== CLEAR) return;
  if (r.getValue() === true) {
    clearAll();
    r.setValue(false);
  }
}

/** ペン／パターン配置：盤面のセルをクリックしたとき */
function onSelectionChange(e) {
  const r = e.range;
  const sheet = r.getSheet();
  if (sheet.getName() !== BOARD) return;
  if (r.getNumRows() !== 1 || r.getNumColumns() !== 1) return;  // 範囲選択は無視
  const row = r.getRow(), col = r.getColumn();
  if (!inBoard_(row, col)) return;

  const mode = String(sheet.getRange(MODE).getValue() || MODE_PAINT);
  if (mode === MODE_PAINT) return;

  if (mode === MODE_PEN) {
    r.setBackground(isAlive_(r.getBackground()) ? DEAD : ALIVE);
    return;
  }

  let shape = findPattern_(mode);
  if (!shape) return;
  const turns = Math.round((parseInt(sheet.getRange(ROTATE).getValue(), 10) || 0) / 90) % 4;
  for (let t = 0; t < turns; t++) shape = rotate_(shape);

  // 盤面からはみ出す部分は切り取る
  const h = Math.min(shape.length, BOARD_TOP + N - row);
  const w = Math.min(shape[0].length, BOARD_LEFT + N - col);
  const out = shape.slice(0, h).map(line => line.slice(0, w).map(v => (v ? ALIVE : DEAD)));
  sheet.getRange(row, col, h, w).setBackgrounds(out);
}

/* ---------------- パターン ---------------- */

/** "O" = 生、"." = 死、"/" = 改行 の文字列 → 0/1 の2次元配列（幅をそろえる） */
function parseShape_(text) {
  const lines = String(text).trim().split(/[\/\n]/).map(s => s.trim());
  const w = Math.max(...lines.map(s => s.length));
  return lines.map(s => {
    const row = [];
    for (let j = 0; j < w; j++) row.push(/[Oo*#1■]/.test(s[j] || '') ? 1 : 0);
    return row;
  });
}

function findPattern_(name) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(PATTERNS);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const hit = rows.find(r => String(r[0]) === name && String(r[1]).trim() !== '');
  return hit ? parseShape_(hit[1]) : null;
}

/** 時計回りに90度回転 */
function rotate_(shape) {
  const h = shape.length, w = shape[0].length, out = [];
  for (let i = 0; i < w; i++) {
    out.push([]);
    for (let j = 0; j < h; j++) out[i].push(shape[h - 1 - j][i]);
  }
  return out;
}

/** 盤面で選択した範囲を、余白を除いてパターン一覧に追加する */
function registerPattern() {
  const ui = SpreadsheetApp.getUi();
  const range = SpreadsheetApp.getActiveRange();
  if (!range || range.getSheet().getName() !== BOARD ||
      !inBoard_(range.getRow(), range.getColumn()) ||
      !inBoard_(range.getLastRow(), range.getLastColumn())) {
    ui.alert('盤面の中でパターンにしたい範囲を選択してから実行してください。');
    return;
  }
  const alive = range.getBackgrounds().map(r => r.map(c => (isAlive_(c) ? 1 : 0)));
  const rows = alive.map((r, i) => (r.some(v => v) ? i : -1)).filter(i => i >= 0);
  const cols = alive[0].map((_, j) => (alive.some(r => r[j]) ? j : -1)).filter(j => j >= 0);
  if (!rows.length) {
    ui.alert('選択範囲に塗られたセルがありません。');
    return;
  }
  const text = alive.slice(rows[0], rows[rows.length - 1] + 1)
    .map(r => r.slice(cols[0], cols[cols.length - 1] + 1).map(v => (v ? 'O' : '.')).join(''))
    .join('/');
  const res = ui.prompt('パターンの名前', text, ui.ButtonSet.OK_CANCEL);
  const name = res.getResponseText().trim();
  if (res.getSelectedButton() !== ui.Button.OK || !name) return;
  SpreadsheetApp.getActive().getSheetByName(PATTERNS).appendRow([name, text]);
  ui.alert('「' + name + '」を登録しました。「入力」プルダウンから選べます。');
}
