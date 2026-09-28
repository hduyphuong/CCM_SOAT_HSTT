"""EAC v2 — EAC = AC + ETC (AC = chi phí thực tế lũy kế · ETC = ước tính phần còn lại), ETC theo BẢN CHẤT từng mã NS (anh Phương chốt 28/09/2026).
Công thức nằm TRONG Excel (truy vết được):
   KL : ETC = Σ(mã NS, ĐVT) MAX(KL NS − KL đội đã làm, 0) × ĐG bình quân đã trả (chưa làm ⇒ ĐG NS)      — sheet R2b_EAC_KL
   TG : ETC = MAX(chi bình quân tháng thực tế, NS / tổng số tháng) × số tháng còn lại
   VT : ETC = đã chi ÷ %hoàn thành doanh thu CĐT − đã chi (chưa chi ⇒ NS)
   NS : ETC = MAX(NS − đã chi, 0)
Mốc dự án ở N1_DanhMuc!V3 (khởi công) · W3 (hoàn thành dự kiến); cut-off N1!E3 (trống ⇒ hôm nay).
Chạy: python eac_v2.py <khung.xlsx> [--ghi]"""
import sys, os, shutil, datetime as dt, pythoncom, win32com.client as w32
sys.stdout.reconfigure(encoding="utf-8")
BAT_DAU, KET_THUC = dt.date(2026, 3, 5), dt.date(2027, 1, 5)
def cach(ma):
    if ma.startswith(("NTP_", "DTC_")): return "KL"
    if ma.startswith("NCC_"): return "NS"
    # vật tư: tạm NS (anh chốt 28/09) — "VT" theo % doanh thu cho số không tin được; chờ C theo định mức × KL còn lại
    if ma in ("Prelim_2", "Prelim_3", "Prelim_5", "Prelim_6", "Prelim_9.1", "Prelim_9.2"): return "TG"
    return "NS"
ser = lambda d: (d - dt.date(1899, 12, 30)).days
khung = sys.argv[1]; ghi = "--ghi" in sys.argv
if ghi:
    os.makedirs(os.path.join(os.path.dirname(khung), "_backup"), exist_ok=True)
    shutil.copy2(khung, os.path.join(os.path.dirname(khung), "_backup", f"{os.path.splitext(os.path.basename(khung))[0]}_truoc_EAC_v2_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"))
pythoncom.CoInitialize(); xl = w32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False; wb = xl.Workbooks.Open(os.path.abspath(khung))
try:
    w1, w2, wr = wb.Worksheets("N1_DanhMuc"), wb.Worksheets("N2_NganSach"), wb.Worksheets("R2_NS_TheoCongTac")
    h2 = [str(w2.Cells(1, c).Value or "") for c in range(1, 20)]; col = lambda t: chr(64 + h2.index(t) + 1)
    cM, cD, cK, cG = col("Mã NS"), col("ĐVT"), col("KL"), col("Giá trị")
    w1.Range("V2").Value = "Ngày khởi công"; w1.Range("W2").Value = "Ngày hoàn thành dự kiến"
    w1.Range("V3").Value = ser(BAT_DAU); w1.Range("W3").Value = ser(KET_THUC); w1.Range("V3:W3").NumberFormat = "dd/mm/yyyy"
    ma_kl = []
    for r in range(3, 80):
        ma = w1.Range(f"L{r}").Value
        if ma:
            c = cach(str(ma)); w1.Range(f"O{r}").Value = c
            if c == "KL": ma_kl.append(str(ma))
    cap = []
    for r in range(2, w2.Cells(w2.Rows.Count, 1).End(-4162).Row + 1):
        m, d = w2.Range(f"{cM}{r}").Value, w2.Range(f"{cD}{r}").Value
        if m in ma_kl and d and (m, str(d).strip().lower()) not in {(a, b.lower()) for a, b in cap}: cap.append((m, str(d).strip()))
    try: ws = wb.Worksheets("R2b_EAC_KL"); ws.Cells.Clear()
    except Exception: ws = wb.Worksheets.Add(After=wr); ws.Name = "R2b_EAC_KL"
    for j, t in enumerate(["Mã NS", "ĐVT", "KL ngân sách", "GT ngân sách", "KL đội đã làm", "GT đã làm", "ĐG dự báo", "KL còn lại", "Chi phí còn lại (ETC)"]): ws.Cells(1, j + 1).Value = t
    N2, N9 = "N2_NganSach!", "N9_TT_DoiTac!"
    for i, (m, d) in enumerate(cap, 2):
        ws.Range(f"A{i}").Value = m; ws.Range(f"B{i}").Value = d
        ws.Range(f"C{i}").Formula = f'=SUMIFS({N2}${cK}$2:${cK}$5000,{N2}${cM}$2:${cM}$5000,$A{i},{N2}${cD}$2:${cD}$5000,$B{i})'
        ws.Range(f"D{i}").Formula = f'=SUMIFS({N2}${cG}$2:${cG}$5000,{N2}${cM}$2:${cM}$5000,$A{i},{N2}${cD}$2:${cD}$5000,$B{i})'
        ws.Range(f"E{i}").Formula = f'=SUMIFS({N9}$H$2:$H$5000,{N9}$M$2:$M$5000,$A{i},{N9}$G$2:$G$5000,$B{i},{N9}$D$2:$D$5000,"THUC_HIEN")+SUMIFS({N9}$H$2:$H$5000,{N9}$M$2:$M$5000,$A{i},{N9}$G$2:$G$5000,$B{i},{N9}$D$2:$D$5000,"DIEU_CHINH")'
        ws.Range(f"F{i}").Formula = f'=SUMIFS({N9}$K$2:$K$5000,{N9}$M$2:$M$5000,$A{i},{N9}$G$2:$G$5000,$B{i},{N9}$D$2:$D$5000,"THUC_HIEN")+SUMIFS({N9}$K$2:$K$5000,{N9}$M$2:$M$5000,$A{i},{N9}$G$2:$G$5000,$B{i},{N9}$D$2:$D$5000,"DIEU_CHINH")'
        ws.Range(f"G{i}").Formula = f"=IF(E{i}>0,F{i}/E{i},IF(C{i}>0,D{i}/C{i},0))"
        ws.Range(f"H{i}").Formula = f"=MAX(C{i}-E{i},0)"
        ws.Range(f"I{i}").Formula = f"=H{i}*G{i}"
    cut = 'IF(N1_DanhMuc!$E$3="",TODAY(),N1_DanhMuc!$E$3)'
    phu = [("% hoàn thành doanh thu CĐT", '=IFERROR((SUMIFS(N8_TT_CDT!$K$2:$K$5000,N8_TT_CDT!$A$2:$A$5000,"HD-CDT",N8_TT_CDT!$D$2:$D$5000,"THUC_HIEN")+SUMIFS(N8_TT_CDT!$K$2:$K$5000,N8_TT_CDT!$A$2:$A$5000,"HD-CDT",N8_TT_CDT!$D$2:$D$5000,"DIEU_CHINH"))/SUMIFS(N4_HD_CDT!$H$2:$H$100,N4_HD_CDT!$A$2:$A$100,"HD-CDT"),0)'),
           ("Tháng đã qua", f"=MAX(1,({cut}-N1_DanhMuc!$V$3)/30.4375)"),
           ("Tháng còn lại", f"=MAX(0,(N1_DanhMuc!$W$3-{cut})/30.4375)"),
           ("Tổng số tháng", "=MAX(1,(N1_DanhMuc!$W$3-N1_DanhMuc!$V$3)/30.4375)")]
    for r_, (t, fx) in enumerate(phu, 1): ws.Range(f"K{r_}").Value = t; ws.Range(f"L{r_}").Formula = fx
    ws.Range("C:F").NumberFormat = '#,##0;[Red]-#,##0;"–"'; ws.Range("H:I").NumberFormat = '#,##0;[Red]-#,##0;"–"'; ws.Range("G:G").NumberFormat = "#,##0"
    ws.Range("L1").NumberFormat = "0.0%"; ws.Range("L2:L4").NumberFormat = "0.0"; ws.Rows(1).Font.Bold = True
    B = "'R2b_EAC_KL'!"; n = 0
    for r in range(2, wr.Cells(wr.Rows.Count, 1).End(-4162).Row + 1):
        if not str(wr.Range(f"K{r}").Formula).startswith('=IF($A'): continue
        wr.Range(f"K{r}").Formula = (f'=IF($A{r}="","",IF(J{r}="KL",IF(COUNTIF({B}$A:$A,$A{r})=0,MAX(F{r}-H{r},0),SUMIFS({B}$I:$I,{B}$A:$A,$A{r})),'
                                     f'IF(J{r}="TG",MAX(H{r}/{B}$L$2,F{r}/{B}$L$4)*{B}$L$3,'
                                     f'IF(J{r}="VT",IF(H{r}=0,F{r},IF({B}$L$1>0,MAX(H{r}/{B}$L$1-H{r},0),MAX(F{r}-H{r},0))),MAX(F{r}-H{r},0)))))'); n += 1
    xl.CalculateFullRebuild()
    print(f"R2 cột K đã thay {n} dòng · %HT doanh thu {ws.Range('L1').Value:.1%} · tháng đã qua {ws.Range('L2').Value:.2f} · còn lại {ws.Range('L3').Value:.2f} · tổng {ws.Range('L4').Value:.2f} · cặp KL {len(cap)}")
    print(f"{'Mã NS':<12}{'cách':<5}{'NS':>16}{'AC':>15}{'ETC':>16}{'EAC':>16}{'NS−EAC':>16}")
    for r in range(2, wr.Cells(wr.Rows.Count, 1).End(-4162).Row + 1):
        a, F_, H_, K_, L_ = wr.Range(f"A{r}").Value, *((wr.Range(f"{c}{r}").Value) for c in "FHKL")
        if a and isinstance(F_, float) and (F_ or H_): print(f"{a:<12}{wr.Range(f'J{r}').Value or '':<5}{F_:>16,.0f}{(H_ or 0):>15,.0f}{(K_ or 0):>16,.0f}{(L_ or 0):>16,.0f}{F_ - (L_ or 0):>16,.0f}")
    R1 = wb.Worksheets("R1_CVR")
    print("R1:", [(R1.Range(f"A{r}").Value, round(R1.Range(f"C{r}").Value or 0)) for r in range(1, 40) if R1.Range(f"A{r}").Value and ("EAC" in str(R1.Range(f"A{r}").Value) or "còn lại" in str(R1.Range(f"A{r}").Value))])
    print("R2b:", [(ws.Range(f"A{i}").Value, ws.Range(f"B{i}").Value, round(ws.Range(f"C{i}").Value or 0, 1), round(ws.Range(f"E{i}").Value or 0, 1), round(ws.Range(f"G{i}").Value or 0), round(ws.Range(f"I{i}").Value or 0)) for i in range(2, len(cap) + 2)])
    if ghi: wb.Save(); print("ĐÃ LƯU")
finally:
    wb.Close(False); xl.Quit()
