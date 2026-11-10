/**
 * ライフゲーム（Googleスプレッドシート用 Apps Script）
 *
 * 次世代の計算はワークシート関数（「計算」シート C36:AF65）が行う。
 * このスクリプトは「塗り → 数値」「数値 → 塗り」の受け渡しと、自動再生のループだけを担当する。
 *
 * 導入: 拡張機能 → Apps Script → このコードを貼り付けて保存 → シートを再読み込み
 *       → メニュー「ライフゲーム」が表示される（初回は権限の承認が必要）
 */
const BOARD = 'ライフゲーム';
const CALC = '計算';
const N = 30;
const BOARD_TOP = 6, BOARD_LEFT = 2;          // 盤面 B6:AE35
const STATE = 'C3:AF32', NEXT = 'C36:AF65';   // 計算シート
const GEN = 'E4', STATUS = 'T4', SPEED = 'AD4';
const DEAD = '#ffffff', ALIVE = '#1f2937';
const TIME_LIMIT_MS = 5.5 * 60 * 1000;        // GAS の実行上限(6分)より手前で止める

function onOpen() {
  SpreadsheetApp.getUi().createMenu('ライフゲーム')
    .addItem('▶ 再生', 'play')
    .addItem('■ 停止', 'stop')
    .addItem('1世代進める', 'stepOnce')
    .addSeparator()
    .addItem('ランダム配置', 'randomize')
    .addItem('全消去', 'clearAll')
    .addItem('世代を0に戻す', 'resetGen')
    .addToUi();
}

function sheets_() {
  const ss = SpreadsheetApp.getActive();
  const board = ss.getSheetByName(BOARD);
  return {
    board: board,
    calc: ss.getSheetByName(CALC),
    grid: board.getRange(BOARD_TOP, BOARD_LEFT, N, N),
  };
}

const isPainted_ = (c) => c && c.toLowerCase() !== DEAD;

/** 1世代進める。盤面が変化しなければ false を返す */
function tick_(s) {
  const colors = s.grid.getBackgrounds();
  const cur = colors.map(r => r.map(c => (isPainted_(c) ? 1 : 0)));
  s.calc.getRange(STATE).setValues(cur);              // 塗り → 数値
  const next = s.calc.getRange(NEXT).getValues();     // 関数で計算された次世代
  let changed = false;
  const out = next.map((r, i) => r.map((v, j) => {
    if (v !== cur[i][j]) changed = true;
    if (v !== 1) return DEAD;
    return cur[i][j] === 1 ? colors[i][j] : ALIVE;   // 生き残りは元の色を保つ
  }));
  s.grid.setBackgrounds(out);                          // 数値 → 塗り
  const g = s.board.getRange(GEN);
  g.setValue((Number(g.getValue()) || 0) + 1);
  s.calc.getRange(STATE).setValues(next);              // 生存セル数の表示を最新に
  return changed;
}

function play() {
  const s = sheets_();
  const status = s.board.getRange(STATUS);
  status.setValue('再生中');
  SpreadsheetApp.flush();
  const start = Date.now();
  while (true) {
    if (status.getValue() !== '再生中') return;
    if (Date.now() - start > TIME_LIMIT_MS) {
      status.setValue('停止中（時間切れ：もう一度再生）');
      return;
    }
    if (!tick_(s)) {
      status.setValue('停止中（変化なし）');
      return;
    }
    SpreadsheetApp.flush();
    const ms = Number(s.board.getRange(SPEED).getValue());
    if (ms > 0) Utilities.sleep(Math.min(ms, 5000));
  }
}

function stop() {
  sheets_().board.getRange(STATUS).setValue('停止中');
}

function stepOnce() {
  tick_(sheets_());
}

function randomize() {
  const s = sheets_();
  const out = [];
  for (let i = 0; i < N; i++) {
    out.push([]);
    for (let j = 0; j < N; j++) out[i].push(Math.random() < 0.25 ? ALIVE : DEAD);
  }
  s.grid.setBackgrounds(out);
  resetGen();
}

function clearAll() {
  const s = sheets_();
  s.grid.setBackground(DEAD);
  s.calc.getRange(STATE).setValue(0);
  resetGen();
}

function resetGen() {
  sheets_().board.getRange(GEN).setValue(0);
}
