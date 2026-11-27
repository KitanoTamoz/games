"""ワークシート関数だけで動くライフゲームの .xlsx を生成するスクリプト。

使い方: python3 build_lifegame.py  →  lifegame.xlsx を出力
マクロ・LAMBDA・反復計算は使わず、SUM / IF / OR / AND / INDEX のみで構成する。
"""
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col
from openpyxl.worksheet.datavalidation import DataValidation

SIZE = 30          # 盤面サイズ（SIZE x SIZE）
GENS = 120         # 計算する世代数（0〜GENS）
STRIDE = SIZE + 1  # 1世代ブロックの行数（空行1行を挟む）
C0 = 3             # 盤面の左端列（C列）。B列・右隣列は空けて「盤外=0」にする
CALC_TOP = 2       # 計算シートで第0世代が始まる行

ALIVE = PatternFill("solid", fgColor="1F2937")
DEAD = PatternFill("solid", fgColor="F3F4F6")
thin = Side(style="thin", color="9CA3AF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

cL, cR = col(C0), col(C0 + SIZE - 1)
calc_last = CALC_TOP + GENS * STRIDE + SIZE - 1


def square_cells(ws, r0, c0, n, width=2.6, height=15):
    for c in range(c0, c0 + n):
        ws.column_dimensions[col(c)].width = width
    for r in range(r0, r0 + n):
        ws.row_dimensions[r].height = height


def life_format(ws, rng, hide_text=True):
    font = Font(color="1F2937") if hide_text else None
    ws.conditional_formatting.add(
        rng, CellIsRule(operator="equal", formula=["1"], fill=ALIVE, font=font))
    ws.conditional_formatting.add(
        rng, CellIsRule(operator="notEqual", formula=["1"], fill=DEAD,
                        font=Font(color="F3F4F6") if hide_text else None))


wb = Workbook()

# ---------------- 盤面（表示）シート ----------------
board = wb.active
board.title = "ライフゲーム"
# ---------------- 初期配置シート ----------------
init = wb.create_sheet("初期配置")
# ---------------- 計算シート ----------------
calc = wb.create_sheet("計算")
# ---------------- 個体数推移シート ----------------
pop = wb.create_sheet("個体数推移")

# ===== 初期配置 =====
I_TOP = 3
init["B1"] = "初期配置（第0世代）：生きたセルに 1 を入力してください（空欄・0 は死）"
init["B1"].font = Font(bold=True)
square_cells(init, I_TOP, C0, SIZE)
for r in range(I_TOP, I_TOP + SIZE):
    for c in range(C0, C0 + SIZE):
        init.cell(r, c).border = BOX
        init.cell(r, c).alignment = Alignment(horizontal="center")
life_format(init, f"{cL}{I_TOP}:{cR}{I_TOP + SIZE - 1}", hide_text=False)

# プリセット：グライダー、R-ペントミノ、ブリンカー、ヒキガエル
presets = {
    (1, 1): [(0, 1), (1, 2), (2, 0), (2, 1), (2, 2)],               # グライダー
    (13, 14): [(0, 1), (0, 2), (1, 0), (1, 1), (2, 1)],             # R-ペントミノ
    (3, 22): [(0, 0), (0, 1), (0, 2)],                               # ブリンカー
    (24, 3): [(0, 1), (0, 2), (0, 3), (1, 0), (1, 1), (1, 2)],      # ヒキガエル
}
for (br, bc), cells in presets.items():
    for dr, dc in cells:
        init.cell(I_TOP + br + dr, C0 + bc + dc, 1)

# ===== 計算 =====
calc["A1"] = "各世代 = 前世代の3x3範囲の合計 S が 3、または S=4 かつ自分が生 → 1"
calc.column_dimensions["A"].width = 10
calc.column_dimensions[col(C0 + SIZE + 1)].width = 10
square_cells(calc, 1, C0, SIZE, width=2.6, height=15)
pop_col = col(C0 + SIZE + 1)  # 盤面右の空列のさらに右

for g in range(GENS + 1):
    top = CALC_TOP + g * STRIDE
    calc.cell(top, 1, f"世代 {g}").font = Font(bold=True)
    calc[f"{pop_col}{top}"] = f"=SUM({cL}{top}:{cR}{top + SIZE - 1})"
    for i in range(SIZE):
        r = top + i
        for j in range(SIZE):
            c = C0 + j
            if g == 0:
                src = f"初期配置!{col(c)}{I_TOP + i}"
                f = f"=IF({src}=1,1,0)"
            else:
                pr = r - STRIDE  # 前世代の同じ位置
                rng = f"{col(c - 1)}{pr - 1}:{col(c + 1)}{pr + 1}"
                me = f"{col(c)}{pr}"
                f = f"=IF(OR(SUM({rng})=3,AND(SUM({rng})=4,{me}=1)),1,0)"
            calc.cell(r, c, f)
life_format(calc, f"{cL}{CALC_TOP}:{cR}{calc_last}", hide_text=False)
calc.freeze_panes = "B2"

# ===== 盤面（表示） =====
B_TOP = 7
board["B1"] = "ライフゲーム（ワークシート関数版）"
board["B1"].font = Font(bold=True, size=16)
board["B2"] = "「初期配置」シートに 1 を入れて配置 → 下の世代番号を変えると盤面が進みます。"
board["B3"] = "ルール：生きたセルは隣接2〜3個で生存、死んだセルは隣接ちょうど3個で誕生（盤外は常に死）"
board["B5"] = "世代"
board["B5"].font = Font(bold=True)
GEN_CELL = "F5"
board.merge_cells("B5:E5")
board.merge_cells("F5:H5")
board[GEN_CELL] = 0
board[GEN_CELL].font = Font(bold=True, size=14, color="B91C1C")
board[GEN_CELL].alignment = Alignment(horizontal="center")
board[GEN_CELL].border = BOX
board["J5"] = "生存セル数"
board.merge_cells("J5:N5")
board["O5"] = f"=INDEX(計算!{pop_col}:{pop_col},{CALC_TOP}+{GEN_CELL}*{STRIDE})"
board.merge_cells("O5:Q5")
board["O5"].font = Font(bold=True, size=14)
board["O5"].alignment = Alignment(horizontal="center")
board["S5"] = f"（0〜{GENS} を入力）"

dv = DataValidation(type="whole", operator="between", formula1="0",
                    formula2=str(GENS), showErrorMessage=True,
                    errorTitle="世代", error=f"0〜{GENS} の整数を入力してください")
board.add_data_validation(dv)
dv.add(GEN_CELL)

square_cells(board, B_TOP, 2, SIZE)
board.column_dimensions["A"].width = 2.6
for i in range(SIZE):
    for j in range(SIZE):
        cell = board.cell(B_TOP + i, 2 + j)
        cell.value = (f"=INDEX(計算!${cL}${CALC_TOP}:${cR}${calc_last},"
                      f"${GEN_CELL}*{STRIDE}+{i + 1},{j + 1})")
        cell.alignment = Alignment(horizontal="center")
life_format(board, f"B{B_TOP}:{col(1 + SIZE)}{B_TOP + SIZE - 1}", hide_text=True)

# ===== 個体数推移 =====
pop["A1"], pop["B1"] = "世代", "生存セル数"
pop["A1"].font = pop["B1"].font = Font(bold=True)
for g in range(GENS + 1):
    pop.cell(g + 2, 1, g)
    pop.cell(g + 2, 2, f"=INDEX(計算!{pop_col}:{pop_col},{CALC_TOP}+A{g + 2}*{STRIDE})")
chart = LineChart()
chart.title = "生存セル数の推移"
chart.x_axis.title = "世代"
chart.y_axis.title = "セル数"
chart.add_data(Reference(pop, min_col=2, min_row=1, max_row=GENS + 2), titles_from_data=True)
chart.set_categories(Reference(pop, min_col=1, min_row=2, max_row=GENS + 2))
chart.legend = None
chart.width, chart.height = 22, 10
pop.add_chart(chart, "D2")

out = __file__.rsplit("/", 1)[0] + "/lifegame.xlsx"
wb.save(out)
print("saved", out)
