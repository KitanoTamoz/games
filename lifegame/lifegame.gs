/**
 * ライフゲーム（Googleスプレッドシート用 Apps Script）
 *
 * 盤面（「ライフゲーム」シート B6:AE35）の塗りつぶしを読み取り、次世代の計算はすべてスクリプト内で行う。
 * 速度のため、再生中は盤面をメモリ上で持ち、1世代あたりのシート操作は
 * 「状態の読み取り1回＋塗りの書き込み」だけにしている。
 *
 * 導入: 拡張機能 → Apps Script → このコードを貼り付けて保存 → シートを再読み込み
 *       → メニュー「ライフゲーム」が表示される（初回は権限の承認が必要）
 */
const BOARD = 'ライフゲーム';
const N = 30;
const BOARD_TOP = 6, BOARD_LEFT = 2;    // 盤面 B6:AE35
const GEN = 'E4', POP = 'M4', CTRL = 'T4:AD4';  // CTRL の先頭=状態(T4), 末尾=間隔ms(AD4)
const DEAD = '#ffffff', ALIVE = '#1f2937';
const TIME_LIMIT_MS = 5.5 * 60 * 1000;  // GAS の実行上限(6分)より手前で止める

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

function board_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(BOARD);
  return { sheet: sheet, grid: sheet.getRange(BOARD_TOP, BOARD_LEFT, N, N) };
}

const isAlive_ = (c) => !!c && c.toLowerCase() !== DEAD;

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

function writeInfo_(sheet, gen, pop) {
  sheet.getRange(GEN).setValue(gen);
  sheet.getRange(POP).setValue(pop);
}

function play() {
  const b = board_();
  const ctrl = b.sheet.getRange(CTRL);
  const status = b.sheet.getRange(CTRL.split(':')[0]);
  status.setValue('再生中');
  let colors = b.grid.getBackgrounds();
  let gen = Number(b.sheet.getRange(GEN).getValue()) || 0;
  const start = Date.now();
  while (true) {
    const c = ctrl.getValues()[0];          // 停止指示と間隔を1回の読み取りで取得
    if (c[0] !== '再生中') return;
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
    const ms = Number(c[c.length - 1]);
    if (ms > 0) Utilities.sleep(Math.min(ms, 5000));
  }
}

function stop() {
  const b = board_();
  b.sheet.getRange(CTRL.split(':')[0]).setValue('停止中');
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
  b.grid.setBackground(DEAD);
  writeInfo_(b.sheet, 0, 0);
}

function resetGen() {
  board_().sheet.getRange(GEN).setValue(0);
}
