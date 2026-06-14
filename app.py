import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io


# ── 自動偵測檔案類型並讀取 ────────────────────────────────────────────────
def detect_and_load(f):
    """
    掃描前4列，自動判斷是「財政部電子發票」還是「會計軟體記錄」。
    返回 (df, file_type)，file_type = 'gov' 或 'acc'
    """
    f.seek(0)
    raw = pd.read_excel(f, sheet_name=0, header=None)

    GOV_KEYS = {"發票狀態", "賣方名稱", "銷售額合計", "買方統一編號"}
    ACC_KEYS = {"交易模式", "收支日期", "會計項目", "資金帳戶"}

    file_type = None
    header_row = 0

    for i in range(min(4, len(raw))):
        vals = set(str(v) for v in raw.iloc[i].tolist())
        if len(GOV_KEYS & vals) >= 2:
            file_type = "gov"
            header_row = i
            break
        if len(ACC_KEYS & vals) >= 1:
            file_type = "acc"
            header_row = i
            break

    # 寬鬆備援
    if file_type is None:
        for i in range(min(4, len(raw))):
            vals = set(str(v) for v in raw.iloc[i].tolist())
            if "發票狀態" in vals or "賣方名稱" in vals:
                file_type = "gov"; header_row = i; break
            if "交易模式" in vals or "收支日期" in vals:
                file_type = "acc"; header_row = i; break

    if file_type is None:
        sample = [str(v) for v in raw.iloc[0].tolist()[:8]]
        raise ValueError(f"無法判斷檔案格式，前幾欄：{sample}")

    df = raw.iloc[header_row + 1:].copy()
    df.columns = [str(c) for c in raw.iloc[header_row].tolist()]
    df = df.reset_index(drop=True)

    if "發票號碼" not in df.columns:
        guess = next((c for c in df.columns if "發票" in c and "號" in c), None)
        if guess:
            df = df.rename(columns={guess: "發票號碼"})
        else:
            raise ValueError(f"找不到發票號碼欄位，現有欄位：{list(df.columns)[:10]}")

    df["發票號碼"] = df["發票號碼"].astype(str).str.strip()
    return df, file_type


# ── 報告產生函式 ──────────────────────────────────────────────────────────
def _build_report(df_gov, df_acc, only_gov, only_acc, both):
    wb = Workbook()

    FONT = "Arial"
    HDR_FT  = Font(name=FONT, bold=True, color="FFFFFF", size=11)
    FT_NORM = Font(name=FONT, size=10)
    FT_BOLD = Font(name=FONT, bold=True, size=10)

    FILL_HDR  = PatternFill("solid", start_color="2F4F4F")
    FILL_GOV  = PatternFill("solid", start_color="FFD7D7")
    FILL_ACC  = PatternFill("solid", start_color="D7EBFF")
    FILL_BOTH = PatternFill("solid", start_color="E8F5E9")
    FILL_DIFF = PatternFill("solid", start_color="FFF3CD")
    FILL_VOID = PatternFill("solid", start_color="F5F5F5")
    FILL_SECT = PatternFill("solid", start_color="EAEAEA")

    AL_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
    AL_L = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    def border():
        s = Side(style="thin", color="CCCCCC")
        return Border(left=s, right=s, top=s, bottom=s)

    def hdr_row(ws, headers, widths):
        ws.row_dimensions[1].height = 22
        for i, (h, w) in enumerate(zip(headers, widths), 1):
            c = ws.cell(row=1, column=i)
            c.value = h; c.font = HDR_FT; c.fill = FILL_HDR
            c.alignment = AL_C; c.border = border()
            ws.column_dimensions[get_column_letter(i)].width = w

    def wcell(ws, r, c, v, fill):
        cell = ws.cell(row=r, column=c)
        cell.value = v; cell.fill = fill
        cell.font = FT_NORM; cell.border = border()
        cell.alignment = AL_L

    def sc(row, col, default=""):
        if col in row.index and pd.notna(row[col]):
            return row[col]
        return default

    # 統計外帳有效/作廢
    if "發票狀態" in df_gov.columns:
        valid_gov = set(df_gov[df_gov["發票狀態"] == "開立已確認"]["發票號碼"]) & only_gov
        void_gov  = set(df_gov[df_gov["發票狀態"] == "作廢已確認"]["發票號碼"]) & only_gov
    else:
        valid_gov = only_gov
        void_gov  = set()

    # ── Sheet 1：差異摘要 ──
    ws1 = wb.active
    ws1.title = "差異摘要"
    ws1.sheet_view.showGridLines = False
    ws1.merge_cells("A1:C1")
    t = ws1["A1"]
    t.value = "進項發票比對報告 — 差異摘要"
    t.font = Font(name=FONT, bold=True, size=14, color="FFFFFF")
    t.fill = FILL_HDR; t.alignment = AL_C
    ws1.row_dimensions[1].height = 28
    ws1.column_dimensions["A"].width = 38
    ws1.column_dimensions["B"].width = 14

    rows_data = [
        ("", ""),
        ("📋 比對統計", None),
        ("外帳（財政部電子發票）筆數", len(df_gov)),
        ("內帳（會計軟體）進項筆數", len(df_acc)),
        ("兩邊吻合", len(both)),
        ("", ""),
        ("⚠️ 差異狀況", None),
        ("外帳有、內帳沒有（共）", len(only_gov)),
        ("  └ 開立已確認", len(valid_gov)),
        ("  └ 作廢已確認", len(void_gov)),
        ("內帳有、外帳沒有", len(only_acc)),
        ("", ""),
        ("📌 顏色說明", None),
        ("淡紅底", "外帳有、內帳沒有（應盡快補記）"),
        ("淡藍底", "內帳有、外帳沒有"),
        ("淡綠底", "兩邊都有（吻合）"),
        ("淡黃底", "兩邊都有但金額不同"),
        ("淺灰底", "作廢發票 / 收據"),
    ]
    FILLS_LEGEND = {
        "淡紅底": FILL_GOV, "淡藍底": FILL_ACC,
        "淡綠底": FILL_BOTH, "淡黃底": FILL_DIFF, "淺灰底": FILL_VOID,
    }
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

    # ── Sheet 2：外帳有_內帳沒有 ──
    ws2 = wb.create_sheet("外帳有_內帳沒有")
    ws2.sheet_view.showGridLines = False
    hdr_row(ws2,
        ["發票號碼", "發票狀態", "發票日期", "賣方名稱", "銷售額合計", "營業稅", "總計"],
        [14, 12, 12, 32, 12, 10, 12])
    diff2 = df_gov[df_gov["發票號碼"].isin(only_gov)].copy()
    if "發票日期" in diff2.columns:
        diff2 = diff2.sort_values("發票日期")
    for r, (_, row) in enumerate(diff2.iterrows(), 2):
        is_void = sc(row, "發票狀態") == "作廢已確認"
        fill = FILL_VOID if is_void else FILL_GOV
        for c, v in enumerate([
            sc(row, "發票號碼"), sc(row, "發票狀態"),
            str(sc(row, "發票日期"))[:10], sc(row, "賣方名稱"),
            sc(row, "銷售額合計"), sc(row, "營業稅"), sc(row, "總計"),
        ], 1):
            wcell(ws2, r, c, v, fill)

    # ── Sheet 3：內帳有_外帳沒有 ──
    ws3 = wb.create_sheet("內帳有_外帳沒有")
    ws3.sheet_view.showGridLines = False
    hdr_row(ws3,
        ["發票號碼", "憑證日期", "收支日期", "對象", "發票金額", "稅額", "銷售額", "附註說明"],
        [14, 12, 12, 24, 12, 8, 12, 36])
    diff3 = df_acc[df_acc["發票號碼"].isin(only_acc)].copy()
    if "憑證日期" in diff3.columns:
        diff3 = diff3.sort_values("憑證日期")
    for r, (_, row) in enumerate(diff3.iterrows(), 2):
        inv = str(sc(row, "發票號碼"))
        is_receipt = inv in ("收據", "無", "nan", "")
        fill = FILL_VOID if is_receipt else FILL_ACC
        for c, v in enumerate([
            inv,
            str(sc(row, "憑證日期"))[:10] if sc(row, "憑證日期") != "" else "",
            str(sc(row, "收支日期"))[:10] if sc(row, "收支日期") != "" else "",
            sc(row, "對象"), sc(row, "發票金額"),
            sc(row, "稅額"), sc(row, "銷售額"), sc(row, "附註說明"),
        ], 1):
            wcell(ws3, r, c, v, fill)

    # ── Sheet 4：兩邊都有_金額核對 ──
    ws4 = wb.create_sheet("兩邊都有_金額核對")
    ws4.sheet_view.showGridLines = False
    hdr_row(ws4,
        ["發票號碼", "外帳日期", "外帳賣方", "外帳金額", "內帳憑證日期", "內帳對象", "內帳金額", "差異"],
        [14, 12, 30, 12, 13, 24, 12, 10])
    r = 2
    for inv in sorted(both):
        o_rows = df_gov[df_gov["發票號碼"] == inv]
        i_rows = df_acc[df_acc["發票號碼"] == inv]
        if o_rows.empty or i_rows.empty:
            continue
        o_row = o_rows.iloc[0]; i_row = i_rows.iloc[0]
        try: o_amt = float(sc(o_row, "總計")) if sc(o_row, "總計") != "" else 0
        except: o_amt = 0
        try: i_amt = float(sc(i_row, "發票金額")) if sc(i_row, "發票金額") != "" else 0
        except: i_amt = 0
        diff = o_amt - i_amt
        fill = FILL_DIFF if abs(diff) > 0.5 else FILL_BOTH
        o_date = str(sc(o_row, "發票日期"))[:10] if sc(o_row, "發票日期") != "" else ""
        i_date = str(sc(i_row, "憑證日期"))[:10] if sc(i_row, "憑證日期") != "" else ""
        for c, v in enumerate([
            inv, o_date, sc(o_row, "賣方名稱"), o_amt,
            i_date, sc(i_row, "對象"), i_amt,
            diff if abs(diff) > 0.5 else "",
        ], 1):
            wcell(ws4, r, c, v, fill)
        r += 1

    return wb


# ── 頁面 ──────────────────────────────────────────────────────────────────
st.set_page_config(page_title="翡閣進項發票比對工具", page_icon="🧾", layout="centered")
st.title("🧾 翡閣進項發票比對工具")
st.caption("上傳兩個 Excel，系統自動辨識格式並產生差異報告")
st.info("💡 兩個檔案可任意順序上傳，不需手動選擇哪個是內帳/外帳", icon="ℹ️")
st.divider()

col1, col2 = st.columns(2)
with col1:
    st.markdown("**📁 檔案一**")
    file_a = st.file_uploader("上傳第一個 xlsx", type=["xlsx"], key="file_a",
                               label_visibility="collapsed")
with col2:
    st.markdown("**📁 檔案二**")
    file_b = st.file_uploader("上傳第二個 xlsx", type=["xlsx"], key="file_b",
                               label_visibility="collapsed")

if file_a and file_b:
    st.divider()
    if st.button("🔍 開始比對", use_container_width=True, type="primary"):
        with st.spinner("自動辨識格式並比對中…"):
            try:
                df_a, type_a = detect_and_load(file_a)
            except Exception as e:
                st.error(f"檔案一讀取失敗：{e}")
                st.stop()
            try:
                df_b, type_b = detect_and_load(file_b)
            except Exception as e:
                st.error(f"檔案二讀取失敗：{e}")
                st.stop()

            if type_a == type_b:
                st.error(f"兩個檔案格式相同（皆為 {'財政部電子發票' if type_a == 'gov' else '會計軟體'}），請確認是否上傳正確")
                st.stop()

            if type_a == "gov":
                df_gov, df_acc = df_a, df_b
                gov_name, acc_name = file_a.name, file_b.name
            else:
                df_gov, df_acc = df_b, df_a
                gov_name, acc_name = file_b.name, file_a.name

            st.success(f"✅ 外帳（財政部）= **{gov_name}** ／ 內帳（會計）= **{acc_name}**")

            EXCLUDE = {"nan", "收據", "無", ""}
            gov_set = set(df_gov["發票號碼"]) - EXCLUDE
            acc_set = set(df_acc["發票號碼"]) - EXCLUDE
            only_gov = gov_set - acc_set
            only_acc = acc_set - gov_set
            both = gov_set & acc_set

            c1, c2, c3 = st.columns(3)
            c1.metric("外帳有、內帳沒有", len(only_gov))
            c2.metric("內帳有、外帳沒有", len(only_acc))
            c3.metric("兩邊吻合", len(both))

            try:
                wb = _build_report(df_gov, df_acc, only_gov, only_acc, both)
            except Exception as e:
                st.error(f"產生報告失敗：{e}")
                st.stop()

            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)

        st.download_button(
            label="⬇️ 下載差異報告 Excel",
            data=buf,
            file_name="進項發票比對報告.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
