import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

st.set_page_config(page_title="翡閣進項發票比對工具", page_icon="🧾", layout="centered")

st.title("🧾 翡閣進項發票比對工具")
st.caption("上傳內帳與外帳 Excel，自動產生差異報告")

st.divider()

col1, col2 = st.columns(2)
with col1:
    st.markdown("**📁 內帳（進銷項發票）**"
    inner_file = st.file_uploader("上傳內帳 xlsx", type=["xlsx"], key="inner",
                                   label_visibility="collapsed")
with col2:
    st.markdown("**📁 外帳（會計師申報）**")
    outer_file = st.file_uploader("上傳外帳 xlsx", type=["xlsx"], key="outer",
                                   label_visibility="collapsed")

if inner_file and outer_file:
    st.divider()
    if st.button("🔍 開始比對", use_container_width=True, type="primary"):
        with st.spinner("比對中，請稍候…"):

            # ── 讀取內帳 ──────────────────────────────────────────────────
            try:
                df_inner = pd.read_excel(inner_file, sheet_name="進項", header=0)
                df_inner["發票號碼"] = df_inner["發票號碼"].astype(str).str.strip()
            except Exception as e:
                st.error(f"讀取內帳失敗：{e}\n請確認工作表名稱為「進項」")
                st.stop()

            # ── 讀取外帳 ──────────────────────────────────────────────────
            try:
                outer_raw = pd.read_excel(outer_file, sheet_name=0, header=None)
                df_outer = outer_raw.iloc[2:].copy()
                df_outer.columns = outer_raw.iloc[1].tolist()
                df_outer = df_outer.reset_index(drop=True)
                df_outer["發票號碼"] = df_outer["發票號碼"].astype(str).str.strip()
            except Exception as e:
                st.error(f"讀取外帳失敗：{e}")
                st.stop()

            # ── 差異計算 ───────────────────────────────────────────────────
            inner_all = set(df_inner["發票號碼"])
            outer_all = set(df_outer["發票號碼"])
            only_inner = inner_all - outer_all
            only_outer = outer_all - inner_all
            both = inner_all & outer_all

            # ── 摘要顯示 ───────────────────────────────────────────────────
            c1, c2, c3 = st.columns(3)
            c1.metric("內帳有、外帳沒有", len(only_inner))
            c2.metric("外帳有、內帳沒有", len(only_outer))
            c3.metric("兩邊吻合", len(both))

            # ── 產生 Excel ─────────────────────────────────────────────────
            wb = _build_report(df_inner, df_outer, only_inner, only_outer, both)

            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)

        st.success("比對完成！")
        st.download_button(
            label="⬇️ 下載差異報告 Excel",
            data=buf,
            file_name="進項發票比對報告.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


# ── 報告產生函式 ───────────────────────────────────────────────────────────
def _build_report(df_inner, df_outer, only_inner, only_outer, both):
    wb = Workbook()

    FONT = "Arial"
    HDR_FT  = Font(name=FONT, bold=True, color="FFFFFF", size=11)
    FT_NORM = Font(name=FONT, size=10)
    FT_BOLD = Font(name=FONT, bold=True, size=10)

    FILL_HDR   = PatternFill("solid", start_color="2F4F4F")
    FILL_INNER = PatternFill("solid", start_color="FFD7D7")
    FILL_OUTER = PatternFill("solid", start_color="D7EBFF")
    FILL_BOTH  = PatternFill("solid", start_color="E8F5E9")
    FILL_VOID  = PatternFill("solid", start_color="F5F5F5")
    FILL_DIFF  = PatternFill("solid", start_color="FFF3CD")
    FILL_SECT  = PatternFill("solid", start_color="EAEAEA")

    AL_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
    AL_L = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    def border():
        s = Side(style="thin", color="CCCCCC")
        return Border(left=s, right=s, top=s, bottom=s)

    def hdr_row(ws, headers, widths, fill=FILL_HDR):
        ws.row_dimensions[1].height = 22
        for i, (h, w) in enumerate(zip(headers, widths), 1):
            c = ws.cell(row=1, column=i)
            c.value = h; c.font = HDR_FT; c.fill = fill
            c.alignment = AL_C; c.border = border()
            ws.column_dimensions[get_column_letter(i)].width = w

    def write_cell(ws, r, c, v, fill, ft=None, al=None):
        cell = ws.cell(row=r, column=c)
        cell.value = v
        cell.fill = fill
        cell.font = ft or FT_NORM
        cell.border = border()
        cell.alignment = al or AL_L

    # ── 工作表1：差異摘要 ─────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "差異摘要"
    ws1.sheet_view.showGridLines = False

    ws1.merge_cells("A1:C1")
    t = ws1["A1"]
    t.value = "進項發票比對報告 — 差異摘要"
    t.font = Font(name=FONT, bold=True, size=14, color="FFFFFF")
    t.fill = FILL_HDR; t.alignment = AL_C
    ws1.row_dimensions[1].height = 28

    inner_valid = len([i for i in only_inner
                       if df_inner[df_inner["發票號碼"]==i]["發票狀態"].values[0]=="開立已確認"])
    inner_void  = len(only_inner) - inner_valid

    rows = [
        ("", ""),
        ("📋 比對統計", None),
        ("內帳進項筆數", len(df_inner)),
        ("外帳不重複發票號", len(set(df_outer["發票號碼"]))),
        ("兩邊吻合", len(both)),
        ("", ""),
        ("⚠️ 差異狀況", None),
        ("內帳有、外帳沒有（共）", len(only_inner)),
        ("  └ 開立已確認", inner_valid),
        ("  └ 作廢已確認", inner_void),
        ("外帳有、內帳沒有", len(only_outer)),
        ("", ""),
        ("📌 顏色說明", None),
        ("淡紅底", "內帳有、外帳沒有"),
        ("淡藍底", "外帳有、內帳沒有"),
        ("淡綠底", "兩邊都有（吻合）"),
        ("淡黃底", "兩邊都有但金額不同"),
        ("淺灰底", "作廢發票 / 非正式發票（收據、無）"),
    ]
    FILLS_LEGEND = {
        "淡紅底": FILL_INNER, "淡藍底": FILL_OUTER,
        "淡綠底": FILL_BOTH, "淡黃底": FILL_DIFF, "淺灰底": FILL_VOID,
    }
    ws1.column_dimensions["A"].width = 34
    ws1.column_dimensions["B"].width = 14

    r = 3
    for lbl, val in rows:
        if val is None:
            c = ws1.cell(row=r, column=1)
            c.value = lbl; c.font = FT_BOLD; c.fill = FILL_SECT
        else:
            c1 = ws1.cell(row=r, column=1)
            c1.value = lbl; c1.font = FT_NORM
            fill = FILLS_LEGEND.get(lbl, PatternFill())
            c1.fill = fill; c1.border = border() if lbl in FILLS_LEGEND else None
            c2 = ws1.cell(row=r, column=2)
            c2.value = val; c2.font = FT_BOLD
            c2.alignment = Alignment(horizontal="center")
        r += 1

    # ── 工作表2：內帳有_外帳沒有 ─────────────────────────────────────────
    ws2 = wb.create_sheet("內帳有_外帳沒有")
    ws2.sheet_view.showGridLines = False
    hdr_row(ws2,
        ["發票號碼","發票狀態","發票日期","賣方名稱","銷售額","營業稅","總計"],
        [14, 12, 12, 30, 10, 10, 10])

    diff2 = df_inner[df_inner["發票號碼"].isin(only_inner)].copy()
    diff2["_s"] = diff2["發票狀態"].map({"作廢已確認": 0, "開立已確認": 1})
    diff2 = diff2.sort_values(["_s","發票日期"])
    for r, (_, row) in enumerate(diff2.iterrows(), 2):
        is_void = row["發票狀態"] == "作廢已確認"
        fill = FILL_VOID if is_void else FILL_INNER
        for c, v in enumerate([
            row["發票號碼"], row["發票狀態"], str(row["發票日期"])[:10],
            row["賣方名稱"], row["銷售額合計"], row["營業稅"], row["總計"]
        ], 1):
            write_cell(ws2, r, c, v, fill)

    # ── 工作表3：外帳有_內帳沒有 ─────────────────────────────────────────
    ws3 = wb.create_sheet("外帳有_內帳沒有")
    ws3.sheet_view.showGridLines = False
    hdr_row(ws3,
        ["發票號碼","憑證日期","收支日期","對象","發票金額","稅額","銷售額","附註說明"],
        [14, 12, 12, 22, 11, 8, 11, 40])

    diff3 = df_outer[df_outer["發票號碼"].isin(only_outer)].copy()
    diff3 = diff3.sort_values("憑證日期")
    for r, (_, row) in enumerate(diff3.iterrows(), 2):
        inv = str(row["發票號碼"])
        is_no_inv = inv in ("收據", "無", "nan", "")
        fill = FILL_VOID if is_no_inv else FILL_OUTER
        for c, v in enumerate([
            inv,
            str(row["憑證日期"])[:10] if pd.notna(row["憑證日期"]) else "",
            str(row["收支日期"])[:10] if pd.notna(row["收支日期"]) else "",
            str(row["對象"]) if pd.notna(row["對象"]) else "",
            row["發票金額"], row["稅額"], row["銷售額"],
            str(row["附註說明"]) if pd.notna(row["附註說明"]) else "",
        ], 1):
            write_cell(ws3, r, c, v, fill)

    # ── 工作表4：兩邊都有_金額核對 ───────────────────────────────────────
    ws4 = wb.create_sheet("兩邊都有_金額核對")
    ws4.sheet_view.showGridLines = False
    hdr_row(ws4,
        ["發票號碼","內帳日期","內帳賣方","內帳總計","外帳日期","外帳對象","外帳發票金額","金額差異"],
        [14, 12, 28, 11, 12, 22, 13, 10])

    r = 2
    for inv in sorted(both):
        inner_row = df_inner[df_inner["發票號碼"] == inv].iloc[0]
        outer_rows = df_outer[df_outer["發票號碼"] == inv]
        i_total = float(inner_row["總計"]) if pd.notna(inner_row["總計"]) else 0
        for _, o_row in outer_rows.iterrows():
            o_total = float(o_row["發票金額"]) if pd.notna(o_row["發票金額"]) else 0
            diff = i_total - o_total
            fill = FILL_DIFF if abs(diff) > 0.5 else FILL_BOTH
            for c, v in enumerate([
                inv, str(inner_row["發票日期"])[:10], inner_row["賣方名稱"], i_total,
                str(o_row["憑證日期"])[:10] if pd.notna(o_row["憑證日期"]) else "",
                str(o_row["對象"]) if pd.notna(o_row["對象"]) else "",
                o_total, diff if abs(diff) > 0.5 else "",
            ], 1):
                write_cell(ws4, r, c, v, fill)
            r += 1

    return wb
