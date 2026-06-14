import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io


# ── 報告產生函式（必須在呼叫前定義）────────────────────────────────────────
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

    def sc(row, col, default=""):
        """安全取欄位值，欄位不存在或為空時返回預設值"""
        if col in row.index and pd.notna(row[col]):
            return row[col]
        return default

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

    # 計算已確認/作廢數量（若無「發票狀態」欄則視所有為已確認）
    if "發票狀態" in df_inner.columns:
        inner_valid = 0
        for inv in only_inner:
            rows = df_inner[df_inner["發票號碼"] == inv]
            if not rows.empty and rows.iloc[0]["發票狀態"] == "開立已確認":
                inner_valid += 1
        inner_void = len(only_inner) - inner_valid
    else:
        inner_valid = len(only_inner)
        inner_void = 0

    rows_data = [
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
    for lbl, val in rows_data:
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
    if "發票狀態" in diff2.columns:
        diff2["_s"] = diff2["發票狀態"].map({"作廢已確認": 0, "開立已確認": 1}).fillna(1)
    else:
        diff2["_s"] = 1
    sort_cols = ["_s"]
    if "發票日期" in diff2.columns:
        sort_cols.append("發票日期")
    diff2 = diff2.sort_values(sort_cols)
    for r, (_, row) in enumerate(diff2.iterrows(), 2):
        is_void = sc(row, "發票狀態") == "作廢已確認"
        fill = FILL_VOID if is_void else FILL_INNER
        for c, v in enumerate([
            sc(row, "發票號碼"),
            sc(row, "發票狀態"),
            str(sc(row, "發票日期"))[:10],
            sc(row, "賣方名稱"),
            sc(row, "銷售額合計"),
            sc(row, "營業稅"),
            sc(row, "總計"),
        ], 1):
            write_cell(ws2, r, c, v, fill)

    # ── 工作表3：外帳有_內帳沒有 ─────────────────────────────────────────
    ws3 = wb.create_sheet("外帳有_內帳沒有")
    ws3.sheet_view.showGridLines = False
    hdr_row(ws3,
        ["發票號碼","憑證日期","收支日期","對象","發票金額","稅額","銷售額","附註說明"],
        [14, 12, 12, 22, 11, 8, 11, 40])

    diff3 = df_outer[df_outer["發票號碼"].isin(only_outer)].copy()
    sort_col3 = "憑證日期" if "憑證日期" in diff3.columns else diff3.columns[0]
    diff3 = diff3.sort_values(sort_col3)
    for r, (_, row) in enumerate(diff3.iterrows(), 2):
        inv = str(sc(row, "發票號碼"))
        is_no_inv = inv in ("收據", "無", "nan", "")
        fill = FILL_VOID if is_no_inv else FILL_OUTER
        for c, v in enumerate([
            inv,
            str(sc(row, "憑證日期"))[:10] if sc(row, "憑證日期") != "" else "",
            str(sc(row, "收支日期"))[:10] if sc(row, "收支日期") != "" else "",
            sc(row, "對象"),
            sc(row, "發票金額"),
            sc(row, "稅額"),
            sc(row, "銷售額"),
            sc(row, "附註說明"),
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
        i_total_raw = sc(inner_row, "總計")
        try:
            i_total = float(i_total_raw) if i_total_raw != "" else 0
        except (ValueError, TypeError):
            i_total = 0
        for _, o_row in outer_rows.iterrows():
            o_total_raw = sc(o_row, "發票金額")
            try:
                o_total = float(o_total_raw) if o_total_raw != "" else 0
            except (ValueError, TypeError):
                o_total = 0
            diff = i_total - o_total
            fill = FILL_DIFF if abs(diff) > 0.5 else FILL_BOTH
            i_date = sc(inner_row, "發票日期")
            o_date = sc(o_row, "憑證日期")
            for c, v in enumerate([
                inv,
                str(i_date)[:10] if i_date != "" else "",
                sc(inner_row, "賣方名稱"),
                i_total,
                str(o_date)[:10] if o_date != "" else "",
                sc(o_row, "對象"),
                o_total,
                diff if abs(diff) > 0.5 else "",
            ], 1):
                write_cell(ws4, r, c, v, fill)
            r += 1

    return wb


# ── 頁面設定 ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="翡閣進項發票比對工具", page_icon="🧾", layout="centered")

st.title("🧾 翡閣進項發票比對工具")
st.caption("上傳內帳與外帳 Excel，自動產生差異報告")

st.divider()

col1, col2 = st.columns(2)
with col1:
    st.markdown("**📁 內帳（進銷項發票���**")
    inner_file = st.file_uploader("上傳內帳 xlsx", type=["xlsx"], key="inner",
                                   label_visibility="collapsed")
with col2:
    st.markdown("**📁 外帳（會計師申報）**")
    outer_file = st.file_uploader("上傳外帳 xlsx", type=["xlsx"], key="outer",
                                   label_visibility="collapsed")

if inner_file and outer_file:
    st.divider()

    # ── 偵測內帳工作表與欄位 ──────────────────────────────────────────────
    inner_file.seek(0)
    xl_inner = pd.ExcelFile(inner_file)
    inner_sheets = xl_inner.sheet_names
    default_sheet = "進項" if "進項" in inner_sheets else inner_sheets[0]
    inner_sheet = st.selectbox(
        "📋 請選擇內帳工作表",
        options=inner_sheets,
        index=inner_sheets.index(default_sheet),
    )

    inner_file.seek(0)
    df_inner_preview = pd.read_excel(inner_file, sheet_name=inner_sheet, header=0, nrows=0)
    inner_cols = df_inner_preview.columns.tolist()

    inv_guess = next((c for c in inner_cols if "發票" in str(c) and "號" in str(c)), None) or \
                next((c for c in inner_cols if "發票" in str(c)), None) or inner_cols[0]
    inner_inv_col = st.selectbox(
        "🔢 請選擇「發票號碼」欄位（內帳）",
        options=inner_cols,
        index=inner_cols.index(inv_guess),
    )

    # ── 偵測外帳欄位 ──────────────────────────────────────────────────────
    outer_file.seek(0)
    outer_raw_preview = pd.read_excel(outer_file, sheet_name=0, header=None, nrows=3)
    outer_cols = outer_raw_preview.iloc[1].tolist()
    outer_cols_str = [str(c) for c in outer_cols]

    outer_inv_guess = next((c for c in outer_cols_str if "發票" in c and "號" in c), None) or \
                      next((c for c in outer_cols_str if "發票" in c), None) or outer_cols_str[0]
    outer_inv_col = st.selectbox(
        "🔢 請選擇「發票號碼」欄位（外帳）",
        options=outer_cols_str,
        index=outer_cols_str.index(outer_inv_guess),
    )

    if st.button("🔍 開始比對", use_container_width=True, type="primary"):
        with st.spinner("比對中，請稍候…"):

            # ── 讀取內帳 ──────────────────────────────────────────────────
            try:
                inner_file.seek(0)
                df_inner = pd.read_excel(inner_file, sheet_name=inner_sheet, header=0)
                df_inner.columns = [str(c) for c in df_inner.columns]
                if inner_inv_col != "發票號碼":
                    df_inner = df_inner.rename(columns={inner_inv_col: "發票號碼"})
                df_inner["發票號碼"] = df_inner["發票號碼"].astype(str).str.strip()
            except Exception as e:
                st.error(f"讀取內帳失敗：{e}")
                st.stop()

            # ── 讀取外帳 ──────────────────────────────────────────────────
            try:
                outer_file.seek(0)
                outer_raw = pd.read_excel(outer_file, sheet_name=0, header=None)
                df_outer = outer_raw.iloc[2:].copy()
                df_outer.columns = [str(c) for c in outer_raw.iloc[1].tolist()]
                df_outer = df_outer.reset_index(drop=True)
                if outer_inv_col != "發票號碼":
                    df_outer = df_outer.rename(columns={outer_inv_col: "發票號碼"})
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
            try:
                wb = _build_report(df_inner, df_outer, only_inner, only_outer, both)
            except Exception as e:
                st.error(f"產生報告失敗：{e}")
                st.stop()

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
