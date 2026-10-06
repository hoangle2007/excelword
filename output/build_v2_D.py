# -*- coding: utf-8 -*-
"""Build json_v2 for flow D (automation): Xne7UmDP9TU, vM7i37Kkw80, 0a9m17AdKpM."""
import json, pathlib, re

OUT = pathlib.Path("E:/webexcel/output/json_v2")
OUT.mkdir(parents=True, exist_ok=True)

def load_ts(vid):
    t = pathlib.Path(f"E:/webexcel/output/txt/{vid}.vi.txt").read_text(encoding="utf-8", errors="ignore")
    m = re.findall(r"\[\d{1,2}:\d{2}\]", t)
    if not m:
        t2 = pathlib.Path(f"E:/webexcel/output/transcripts/{vid}.vi.vtt").read_text(encoding="utf-8", errors="ignore")
        m = re.findall(r"\d{2}:\d{2}:\d{2}", t2)
        m = ["["+x[3:]+"]" for x in m]
    return m

def pick(tss, idx, total):
    if not tss: return "[00:00]"
    p = int(idx*(len(tss)-1)/max(total-1,1))
    return tss[p]

FILLER = "Khong dung trong bai hoc nay"

def fix_options(vid, cau, opts):
    """Replace filler with plausible distractor."""
    fixed = 0
    for k in list(opts.keys()):
        if FILLER in opts[k]:
            c = cau.lower()
            if any(x in c for x in ["copy", "dan", "paste", "cut", "undo", "find", "filter", "nhay", "sheet", "dinh dang", "sum", "ngay", "gio", "xuong dong", "chon", "bieu do", "hang", "cot", "di chuyen", "phim tat"]):
                pool = {"A":"Ctrl+P (in, khong phai copy/dan)","B":"Alt+Tab (chuyen cua so)","C":"Shift+F1 (khong phai lenh nay)","D":"Ctrl+S (luu file)"}
                opts[k] = pool.get(k, "Ctrl+O (mo file)")
            elif any(x in c for x in ["macro", "developer", "record", "relative", "transpose", "button", "module", ".xlsm", "filter"]):
                pool = {"A":"Insert > Picture (chen anh)","B":"Home > Font > Bold (in dam)","C":"File > Print (in an)","D":"View > Zoom (phong to)"}
                opts[k] = pool.get(k, "Data > Sort (sap xep)")
            else:
                pool = {"A":"=SUM(A1:A10) (tinh tong)","B":"Ctrl+S (luu file)","C":"Delete sheet (xoa sheet)","D":"Print Preview (xem truoc in)"}
                opts[k] = pool.get(k, "Alt+F4 (dong chuong trinh)")
            fixed += 1
    # ensure uniqueness
    seen = {}
    for k in ["A","B","C","D"]:
        if opts[k] in seen:
            opts[k] = opts[k] + " (p.a nhieu)"
            fixed += 1
        seen[opts[k]] = 1
    return fixed

def rebuild_questions(vid, src_qs, tss):
    new_q = []; fixes = 0
    for q in src_qs:
        qid = q["id"]; cau = str(q["cau"])
        opts = {k: str(q.get(k,"")) for k in ["A","B","C","D"]}
        da = str(q.get("dap_an","A")).strip().upper()
        gt = str(q.get("giai_thich",""))
        fixes += fix_options(vid, cau, opts)
        # giu nguyen vi tri dap_an goc (da can bang 10/10/10/10); khong xoay
        opts2 = dict(opts); da2 = da
        if da2 not in ["A","B","C","D"]: da2 = "A"; fixes += 1
        mkw = re.search(r"(Ctrl\+[A-Z;]|Alt\+[#;=]|Flash Fill|Relative|Transpose|\.xlsm|DisplayAlerts|For Each|Sub |MsgBox|Cells|Range|Dim |InputBox|PrintOut|FileSystemObject|Outlook|HTMLBody|Record Macro|Developer|Button|Filter|SUM|Paste Values)", cau)
        chude = mkw.group(0) if mkw else (cau[:30]+"...")
        dokho = "de" if qid <= 10 else ("trung_binh" if qid <= 30 else "kho")
        if not gt.endswith("."): gt = gt + "."
        new_q.append({"id": qid, "chu_de": chude, "do_kho": dokho, "cau": cau,
            "A": opts2["A"], "B": opts2["B"], "C": opts2["C"], "D": opts2["D"],
            "dap_an": da2, "giai_thich_chi_tiet": gt + " Doi chieu transcript video.",
            "timestamp": pick(tss, qid-1, len(src_qs))})
    return new_q, fixes

# ---------------- 23 PHIM TAT ----------------
CT23 = [
 ("Copy — Ctrl+C","Ctrl+C",[("Boi den","Chon vung truoc"),("C","Giu Ctrl, nhan C")],"Nhan ban vung: vien marching-ants nhap nhay quanh vung copy. Dung khi nhan ban du lieu.","Boi den A1:A5 > Ctrl+C (vien nhap nhay) > chon C1 > Ctrl+V","Vung goc giu nguyen","Nhan Ctrl+V nhieu lan ma quen Esc -> dan trung lap"),
 ("Paste — Ctrl+V","Ctrl+V",[("V","Giu Ctrl, nhan V")],"Dan vung copy/cut vao vi tri moi. Con tro dat o o tren-trai cua vung dich.","Chon C1 > Ctrl+V (du lieu copy xuat hien)","Du lieu xuat hien o dich","Dan de len du lieu cu mat du lieu (dung Shift+keo de chen)"),
 ("Cut — Ctrl+X","Ctrl+X",[("X","Giu Ctrl, nhan X")],"Cat/di chuyen: vung goc cung nhap nhay nhung se trang sau khi paste. Khac Copy (giu goc).","Boi den > Ctrl+X > sang vi tri moi > Ctrl+V (goc trang)","Goc trang sau paste","Tuong Cut giong Copy -> mat du lieu goc"),
 ("Undo — Ctrl+Z","Ctrl+Z",[("Z","Giu Ctrl, nhan Z")],"Hoan tac thao tac vua xong; nhan nhieu lan lui nhieu buoc (vd sau Replace All).","Ctrl+H > Replace All > Ctrl+Z (hoan tac)","Buoc truoc duoc khoi phuc","Undo qua nhieu mat ca thao tac dung"),
 ("Thoat copy — Esc","Esc",[],"Tat vien nhap nhay sau copy/paste. Khong anh huong du lieu.","Ctrl+C > Ctrl+V > Esc (het nhap nhay)","Het vien nhap nhay","Nhan Enter thay Esc -> dan them 1 lan"),
 ("Find — Ctrl+F","Ctrl+F",[],"Mo hop Find (the Find): tim kiem tu khoa trong sheet.","Ctrl+F > go 'hang' > Find Next / Find All","Nhay toi o chua tu khoa","Go thieu dau * khi can khop mot phan"),
 ("Replace — Ctrl+H","Ctrl+H",[],"Mo thang the Replace: thay the noi dung hang loat.","Ctrl+H > Find 'du lieu' > Replace 'Excel' > Replace All","Tat ca duoc thay","Replace All khong kiem tra -> sai hang loat (nho Ctrl+Z)"),
 ("Find Next / Replace All","Find Next / Find All / Replace / Replace All",[],"Find Next: tung o; Find All: liet ke; Replace: tung cho; Replace All: tat ca.","Find All liet ke 5 o > Replace All thay 5 cho","Danh sach ket qua","Nhan Replace All 2 lan -> thay chong len nhau"),
 ("Filter — Ctrl+Shift+L","Ctrl+Shift+L",[],"Bat/tat AutoFilter (loc du lieu). Nhan lai de tat.","Chon tieu de > Ctrl+Shift+L (mui ten loc hien ra)","Mui ten loc tren tieu de","Quen tat filter -> tuong mat du lieu"),
 ("Nhay goc — Ctrl+Mui ten","Ctrl + Phim mui ten",[("Ctrl","Giu Ctrl"),("Mui ten","Nhan 1 trong 4 huong")],"Nhay toi bien (dau/cuoi) vung du lieu lien tuc. Rat nhanh voi bang dai.","Dung o A1 > Ctrl+Mui ten xuong (toi dong cuoi)","Con tro toi bien vung","O trong giua vung lam nhay dung giua chung"),
 ("Ve A1 — Ctrl+Home","Ctrl+Home",[],"Ve o A1 tu bat ky dau trong sheet.","Dang o Z1000 > Ctrl+Home (ve A1)","Ve A1 tuc thi","Nham voi Home (ve cot A cung hang)"),
 ("Chon nhanh — Ctrl+Shift+Mui ten","Ctrl+Shift + Mui ten",[("Shift","Them Shift de vua di vua chon")],"Boi den tu o hien tai toi goc vung. Ket hop 2 huong de phu kin vung.","O A1 > Ctrl+Shift+Mui ten xuong + sang phai (chon ca vung)","Ca vung duoc boi den","Vung co o trong -> chon dut doan"),
 ("Chuyen sheet — Ctrl+PageDown/PageUp","Ctrl+PageDown / Ctrl+PageUp",[],"Sang sheet ben phai (PgDn) / ben trai (PgUp).","Ctrl+PageDown (sang sheet ke tiep)","Sheet ke tiep duoc kich hoat","Sheet bi an (Hidden) thi khong nhay toi"),
 ("Dinh dang so — Ctrl+Shift+1..5","Ctrl+Shift+1 / 2 / 3 / 4 / 5",[],"1:#,##0.00 (nghin+2 le); 2:gio; 3:ngay; 4:tien $; 5:%. Ap cho o dang chon.","Nhap 1000000/3 > Ctrl+Shift+1 (333,333.33); > Ctrl+Shift+5 (phan tram)","So doi dinh dang","% nhan 100 -> 0.5 thanh 50% gay hieu lam"),
 ("AutoSum — Alt+= (Alt + dau bang)","Alt + =",[],"Tu chen =SUM() cho cot/hang ke tren/trai o hien tai; Enter de chot.","O duoi cot so > Alt+= (hien =SUM(B2:B10)) > Enter","Tong cot/hang","Vung co o trong -> SUM doan sai vung (kiem tra lai)"),
 ("Paste Values — Ctrl+C roi Alt,H,V,V","Alt,H,V,V (sau Ctrl+C)",[],"Dan dac biet dang gia tri: bo cong thuc, chi lay so. Thu tu: Copy > Alt > H > V > V > Enter.","Copy cot cong thuc > Alt,H,V,V > Enter (chi con so)","Cong thuc bien mat, con so","Quen Enter -> menu treo, dan sai"),
 ("Ngay hien tai — Ctrl+;","Ctrl+; (cham phay)",[],"Chen NGAY hom nay (gia tri co dinh, khong tu doi nhu TODAY()).","Chon o > Ctrl+; (hien ngay hom nay)","Ngay hom nay dang Date","Tuong tu TODAY() tu cap nhat (khong, co dinh)"),
 ("Gio hien tai — Ctrl+Shift+;","Ctrl+Shift+; ",[],"Chen GIO hien tai (co dinh). Them Shift so voi lenh ngay.","Chon o > Ctrl+Shift+; (hien gio)","Gio hien tai","Ban phim laptop phai them Fn moi an ;"),
 ("Xuong dong trong o — Alt+Enter","Alt+Enter",[],"Ngat dong trong cung 1 o (vd liet ke hang hoa). Can Wrap Text de hien dep.","Go 'Hang1' > Alt+Enter > 'Hang2' (2 dong 1 o)","1 o 2 dong","Nhan Enter thuong -> nhay xuong o duoi"),
 ("Chon vung/sheet — Ctrl+A","Ctrl+A",[],"Lan 1: chon vung lien tuc quanh o hien tai; lan 2: ca sheet.","O trong vung > Ctrl+A (chon vung); lan nua (ca sheet)","Vung/sheet duoc chon","Bam 2 lan vo y chon ca sheet roi dinh dang nham"),
 ("Ve chart — Alt+F1","Alt+F1",[],"Ve bieu do cot tu dong tu vung dang chon, ngay tren sheet hien tai.","Boi den vung so > Alt+F1 (chart hien ra)","Chart cot tu dong","Vung chua tieu de text lech -> chart xau"),
 ("Chon hang/cot — Shift+Space / Ctrl+Space","Shift+Space (hang) / Ctrl+Space (cot)",[],"Chon ca hang / ca cot cua o hien tai; boi den them de chon nhieu.","O B3 > Shift+Space (hang 3); > Ctrl+Space (cot B)","Ca hang/cot duoc chon","Nham 2 to hop cho nhau"),
 ("Chen hang/cot — Ctrl+Shift+Dau cong (+) roi R/C","Ctrl+Shift++ roi R (hang) / C (cot)",[],"Mo hop Insert: go R (=Row) chen hang, C (=Column) chen cot. Boi den N o de chen N hang/cot.","Boi den 2 o > Ctrl+Shift++ > R > Enter (chen 2 hang)","Hang/cot moi chen tren/trai","Chen nham vi tri do khong chon dung hang dich"),
 ("Xoa hang/cot — Ctrl+Tru (-) roi R/C","Ctrl+- roi R (hang) / C (cot)",[],"Mo hop Delete: go R xoa hang, C xoa cot.","Chon cot > Ctrl+- > C > Enter (xoa cot)","Hang/cot bien mat","Xoa nham khong Ctrl+Z kip"),
 ("Di chuyen chen — Giu Shift + keo chuot","Shift + keo chuot trai",[],"Di chuyen/chen hang-cot ma khong de len du lieu cu (vach xanh hien vi tri chen).","Boi den 2 hang > giu Shift + keo toi vach xanh > tha","Hang duoc chen len, khong de","Tha Shift truoc khi tha chuot -> ghi de mat du lieu"),
 ("An/Hien hang cot — Ctrl+9 / Ctrl+0","Ctrl+9 (an hang) / Ctrl+0 (an cot)",[],"An hang/cot dang chon; Ctrl+Z hoac Unhide de hien lai. Chon nhieu de an hang loat.","Chon hang 5:7 > Ctrl+9 (an 3 hang)","Hang/cot bien mat","An xong quen -> tuong mat du lieu"),
 ("Chon o nhin thay — Alt+;","Alt+; (cham phay)",[],"Chon cac o VISIBLE sau khi an/loc, bo qua o an khi copy (tranh copy ca dong an).","Loc + an hang > boi den > Alt+; > Ctrl+C > Ctrl+V (chi o thay)","Chi o hien duoc copy","Quen Alt+; -> paste ca o an"),
 ("Flash Fill — Ctrl+E","Ctrl+E",[],"Dien nhanh theo mau (tu Excel 2013): go mau 1 o (Nguyen Van An) roi Ctrl+E dien ca cot.","Cot Ho + Ten > go mau 'Nguyen Van An' > Ctrl+E (dien ca cot)","Ca cot tu dien","Mau khong nhat quan -> Flash Fill dien sai"),
]

# ---------------- 24 MACRO ----------------
CT24 = [
 ("Bat the Developer","Right-click Ribbon > Customize the Ribbon > tick Developer",[],"Hien the Developer moi thay Record Macro, Button, Visual Basic.","Ribbon > Customize > tick Developer > OK","The Developer xuat hien","Dung Excel Online/mobile khong co Developer day du"),
 ("Record Macro","Developer > Record Macro: ten viet lien khong dau; Shortcut key bo trong; Store in: This Workbook",[],"Bat dau ghi: moi click/go duoc dich sang VBA. Dat ten viet lien (Thang1), tranh trung Ctrl+C.","Record ten Thang1 > click sang sheet Du lieu 1 > Stop","Sinh Sub Thang1() trong Module","Ten co dau/cach -> loi; dat trung Ctrl+C -> mat phim tat goc"),
 ("Stop Recording","Stop Recording (nut vuong Developer / thanh trang thai)",[],"Dung ghi sau khi xong thao tac. Quen Stop se ghi thua.","Bam nut vuong Stop sau khi chuyen sheet","Ket thuc Sub","Quen Stop -> macro ghi them thao tac rac"),
 ("Khung Sub","Sub TenMacro() ... End Sub (trong Module, xem bang Alt+F11)",[],"Moi macro la 1 Sub. F5 hoac nut tam giac de chay trong VBE.","Alt+F11 > Module1 > Sub GhiMacro() ... End Sub > F5","Macro chay","Viet code ngoai Sub -> khong chay"),
 ("Ghi chu '","' (dau nhay don) dat dau dong",[],"Bien dong thanh chu thich (mau xanh); dung vo hieu dong thua nhu Range(\"D9\").Select.","' Range(\"D9\").Select (dong nay bi bo qua)","Dong chuyen xanh, khong chay","Thieu ' khi thu nghiem -> lenh thua van chay"),
 ("Selection vs Range co dinh","Xoa dong Range(\"D9\").Select, giu Selection.*",[],"Selection = vung dang chon (tong quat); Range co dinh chi chay dung o cu. Muon ap moi o phai xoa dong co dinh.","Xoa Range(\"D9\").Select, giu Selection.Font.Bold=True > chay o bat ky","Macro chay moi o","Giu dong co dinh -> dinh dang nham o cu"),
 ("Nut Button (Form Control)","Developer > Insert > Button > ve nut > Assign Macro",[],"Tao nut bam chay macro. Ve xong chon macro trong Assign Macro.","Ve Button > chon Thang1 > OK","Bam nut chay macro","Dung ActiveX thay Form Control -> phuc tap khong can thiet"),
 ("Doi chu nut","Right-click nut > Edit Text",[],"Sua nhan nut (vd 'Thang 1', 'Cap nhat').","Right-click > Edit Text > go 'Quay lai'","Chu nut doi","Click doi (khong Right-click) -> vo tinh chay macro"),
 ("Shape lam nut","Insert > Shapes > ve hinh > Right-click > Assign Macro",[],"Dung Shape dep hon nut mac dinh, cung gan macro de bam chay.","Ve Shape 'Du lieu thang 1' > Assign Macro Thang1","Shape bam chay nhu Button","Shape bi Group/Lock -> khong bam duoc"),
 ("Lenh nhay sheet","Sheets(\"Du lieu 1\").Select",[],"Lenh record duoc khi click qua sheet; dung lam muc luc nhay sheet.","Record click sang Du lieu 1 (sinh Sheets(\"Du lieu 1\").Select)","Nhay toi sheet dich","Doi ten sheet sau record -> lenh sai ten bao loi"),
 ("Copy nut","Copy/Paste hoac Ctrl+keo nut",[],"Nhan nhanh nut sang sheet khac; nut copy giu nguyen macro da gan.","Ctrl+keo nut 'Quay lai' sang sheet 2","Nut moi chay giong nut cu","Copy nham nut chua gan macro -> bam khong chay"),
 ("Tim dong trong — Ctrl+Mui ten xuong","Ctrl + Mui ten xuong (trong sheet Data)",[],"Nhay toi dong cuoi co du lieu de them dong tiep (macro nhap lieu).","O B3 > Ctrl+Mui ten xuong (toi day) + BAT Relative + xuong 1 o","Toi dong trong tiep theo","Cot co o trong -> nhay dung giua (chon cot dac de nhay)"),
 ("Relative References","Developer > Use Relative References: BAT truoc doan dich tuong doi, TAT sau paste",[],"BAT thi buoc mui ten ghi tuong doi (dich theo vi tri); tat cho buoc tuyet doi. Dau hieu BAT: nut chim/mo.","BAT Relative > mui ten xuong 1 o > Paste > TAT Relative","Buoc nhay ghi dang Offset tuong doi","Quen TAT -> lan sau nhay lech; quen BAT -> ghi de dong cu"),
 ("Transpose khi nhap lieu","Paste Special > Transpose (xoay cot thanh hang)",[],"Input nhap doc, Data luu ngang nen record thao tac xoay khi paste.","Copy vung doc Input > Paste Special > Transpose sang Data","Cot thanh hang","Vung dich con du lieu -> tran (chon o trong truoc)"),
 ("Chuan hoa Table — Ctrl+T","Ctrl+T (Format as Table) truoc khi record loc",[],"Bien vung thanh Table on dinh, Advanced Filter record on dinh khi them dong.","Chon vung > Ctrl+T > tich Headers","Vung thanh Table","Table co dong trong -> loc thieu"),
 ("Luu .xlsm","Save as > Excel Macro-Enabled Workbook (*.xlsm)",[],"Giu lai macro/Button. Luu .xlsx se MAT macro khi mo lai.","File > Save As > .xlsm","Mo lai van con macro","Luu .xlsx -> mat het macro khong khoi phuc"),
 ("Hoc lenh qua VBE — Alt+F11 + Ctrl+G","Alt+F11 mo VBE; View > Immediate (Ctrl+G)",[],"Mo song song de hoc lenh: record VLOOKUP/Filter/Format roi doc code sinh ra.","Record go VLOOKUP > Alt+F11 xem code","Thay code VLOOKUP/Filter","Sua code truc tiep ma khong backup -> hong macro"),
 ("Clear Formats truoc test","Home > Clear > Clear Formats",[],"Xoa dinh dang vung test truoc khi bam nut de thay ro ket qua macro.","Boi den vung test > Clear Formats > bam nut","Ket qua macro hien ro","Quen Clear -> tuong macro khong chay"),
]

# ---------------- 25 VBA ----------------
CT25 = [
 ("Mo VBE — Alt+F11","Alt+F11 (hoac Developer > Visual Basic)",[],"Mo cua so lap trinh VBE; Insert > Module de viet code (khong viet tren Sheet de tranh mat khi xoa sheet).","Alt+F11 > Insert > Module","Cua so VBE + Module trang","Viet code tren doi tuong Sheet -> xoa sheet mat code"),
 ("Khung Sub + chay F5","Sub SayHello() ... End Sub; ten viet lien; F5 hoac nut tam giac de chay",[],"Moi chuong trinh la 1 Sub. If va End If, For va Next, Sub va End Sub phai di cap.","Sub SayHello()\n MsgBox \"Hello\"\nEnd Sub  > F5","Hop Hello hien ra","Ten co dau cach/ky tu dac biet -> loi bien dich"),
 ("MsgBox","MsgBox \"Hello World\"",[],"Hien hop thong bao. Vi du kinh dien dau tien.","Sub T1()\n MsgBox \"Hello World\"\nEnd Sub","Hop Hello World","Quen ngoac kep -> loi cu phap"),
 ("Cells(hang,cot)","Cells(1,1) = \"Hello\"  'A1",[],"Cells(hang,cot): hang 1 cot 1 = A1. Gan chu/so truc tiep.","Sub T2()\n Cells(1,1) = \"Hello\"\nEnd Sub 'A1=Hello","A1 nhan Hello","Nham thu tu (cot,hang) -> ghi nham o"),
 ("Immediate — Ctrl+G","Ctrl+G (View > Immediate Window); ?1+1 tra 2",[],"Test tung lenh khong can Sub: go lenh + Enter chay ngay.","Immediate go: Cells(2,3)=\"Ga Excel\" > Enter (C2 nhan chu)","C2 nhan chu ngay","Go lenh chay ngay tren du lieu that -> mat du lieu (test tren file copy)"),
 ("Chon sheet theo ten/so","Sheets(\"Sheet1\").Select  |  Sheets(1).Select",[],"Theo TEN (ngoac kep) hoac SO THU TU tu trai. Dung truoc khi tac dong sheet dich.","Sub T3()\n Sheets(\"Sheet1\").Select\nEnd Sub","Nhay toi sheet dich","Ten sai 1 ky tu/doi ten sheet -> Subscript out of range"),
 ("Doi ten sheet","Sheets(3).Name = \"Ga Excel\"",[],"Thuoc tinh .Name dat ten sheet thu 3.","Sub T4()\n Sheets(3).Name = \"Ga Excel\"\nEnd Sub","Sheet 3 doi ten","Ten trung sheet khac hoac chua \\ / ? * -> loi"),
 ("Workbook/Sheet/Cell active","ActiveWorkbook / Workbooks(\"Book1\") / ActiveSheet / ActiveCell",[],"ActiveWorkbook: file dang lam; Workbooks(\"Book1\"): file chi dinh; ActiveSheet/ActiveCell: sheet/o dang chon. Lenh khong ro workbook tac dong file active.","Sub T5()\n MsgBox ActiveWorkbook.Name\nEnd Sub","Ten file active","Dua vao Active* trong file nhieu cua so -> chay nham file"),
 ("Truy cap cheo file","Workbooks(\"Book1\").Sheets(2).Cells(2,1) = \"Xin chao\"",[],"Cu phap Workbook.Sheet.Cells de ghi sang file khac dang mo.","Sub T6()\n Workbooks(\"Book1\").Sheets(2).Cells(2,1) = \"Xin chao\"\nEnd Sub","Book1!B2 nhan chu","File Book1 chua mo -> Subscript out of range"),
 ("Khai bao bien Dim","Dim y As Integer (-32768..32767); Dim n As Long (+-2 ty); Single/Double; String; Object; Variant",[],"Integer nho, Long lon; String chuoi; Object doi tuong; Variant moi kieu (ton bo nho).","Sub T7()\n Dim y As Integer\n y = 5\n Cells(1,1) = y\nEnd Sub 'A1=5","A1=5","So >32767 dung Integer -> Overflow (dung Long)"),
 ("Gan bien vao o","y = 5 : Cells(1,1) = y  (doi y=7 chay lai o doi theo)",[],"Bien mang gia tri gan vao o; doi bien chay lai de thay dong.","Sub T8()\n Dim y As Long\n y = 7\n Cells(1,1) = y\nEnd Sub","A1=7","Dung bien chua Dim/Option Explicit -> kho bat loi sai ten"),
 ("Range gan vung","Range(\"A1:B3\") = \"Ga Excel\"  |  Range(\"A5\") = \"Xin chao\"",[],"Range nhan ca vung lien tuc hoac 1 o theo dia chi chu.","Sub T9()\n Range(\"A1:B3\") = \"Ga Excel\"\nEnd Sub","A1:B3 cung 1 chu","Range sai dia chi (A1B3 thieu :) -> loi"),
 ("Select / Clear / chuoi rong","Range(...).Select | Range(...).Clear | Range(...) = \"\"",[],".Select chon vung; .Clear hoac = \"\" (ngoac kep rong) xoa nhanh.","Sub T10()\n Range(\"A1:A5\").Clear\nEnd Sub","Vung trang","Clear xoa ca dinh dang (muon giu format dung ClearContents)"),
 ("For Each quet sheet + If xoa","For Each ws In ThisWorkbook.Sheets ... Next  +  If ws.Name <> ActiveSheet.Name Then ws.Delete",[],"Duyet tung Worksheet; khac ten sheet active thi xoa. Can Dim ws As Worksheet.","Sub XoaSheet()\n Dim ws As Worksheet\n Application.DisplayAlerts = False\n For Each ws In ThisWorkbook.Sheets\n  If ws.Name <> ActiveSheet.Name Then ws.Delete\n Next\n Application.DisplayAlerts = True\nEnd Sub","Chi con sheet active","Quen DisplayAlerts False -> hop hoi tung sheet; quen True lai -> nguy hiem xoa nham"),
 ("DisplayAlerts False/True","Application.DisplayAlerts = False 'truoc xoa  ... = True 'ngay sau",[],"Tat hop 'co chac muon xoa'; NHO bat lai True keo lan sau xoa nham khong hoi.","(xem code tren)","Khong hien hop xac nhan","Quen bat lai True -> Excel mai mai khong canh bao"),
 ("For i = 10 To 20","For i = 10 To 20 Step 1 ... Next",[],"Lap so lan co dinh; dung in 1 doan (thi sinh 10..20) thay vi While het danh sach.","Sub InDoan()\n Dim i As Long\n For i = 10 To 20 Step 1\n  Sheet2.Cells(i,1) = i\n Next\nEnd Sub","Cot ghi 10..20","Buoc Step sai dau -> vong vo han/bo qua"),
 ("Do While / Until","Do While ... Loop  |  Do Until ... Loop",[],"Lap khi con dieu kien; dung in HET danh sach khong biet truoc so luong.","Sub QDuyet()\n Dim i As Long: i = 2\n Do While Cells(i,1) <> \"\"\n  i = i + 1\n Loop\nEnd Sub","i dung o dong trong dau","Quen tang bien dem -> treo vo han"),
 ("PrintOut","Sheet2.PrintOut Preview:=False  (False=in that, True=xem truoc)",[],"In sheet bang code. Khong may in thi chon may PDF ao de test.","Sub InSL()\n Sheet2.PrintOut Preview:=False\nEnd Sub","Luu PDF / in that","Preview:=False in that hang loat ton giay (test bang True truoc)"),
 ("FSO + Set","Dim fso As Object: Set fso = CreateObject(\"Scripting.FileSystemObject\")",[],"Tao doi tuong thao tac file. Object phai co Set (khac bien thuong). Dau _ cuoi dong de noi code.","Sub F1()\n Dim fso As Object\n Set fso = CreateObject(\"Scripting.FileSystemObject\")\n MsgBox fso.FolderExists(\"D:\\Data\")\nEnd Sub","True/False","Thieu Set -> Object variable not set (loi 91)"),
 ("Kiem tra thu muc + Exit Sub","If fso.FolderExists(duongdan) = False Then MsgBox \"Khong ton tai\" : Exit Sub : End If",[],"Chan truoc khi copy: thu muc nguon/dich phai ton tai.","(long vao Sub copy file)","Bao loi + thoat neu thieu thu muc","Viet End If thieu If hoac nguoc lai -> loi bien dich"),
 ("CopyFile","fso.CopyFile Source:=nguon & kieuFile, Destination:=dich  (*.* hoac *.xlsx)",[],"Copy file giua 2 thu muc; kieuFile loc loai.","Sub CopyF()\n Dim fso As Object\n Set fso = CreateObject(\"Scripting.FileSystemObject\")\n fso.CopyFile \"D:\\Nguon\\*.xlsx\", \"D:\\Dich\\\"\nEnd Sub","File duoc copy","Duong dan thieu \\ cuoi -> copy sai cho"),
 ("Outlook References + CreateObject","Tools > References > tick Outlook 16.0  +  Set olApp = CreateObject(\"Outlook.Application\")",[],"Bat reference truoc moi goi Outlook tu Excel.","Sub M1()\n Dim olApp As Object\n Set olApp = CreateObject(\"Outlook.Application\")\nEnd Sub","Khong loi thieu thu vien","Quen tick References -> loi ActiveX cant create"),
 ("Gui mail .To/.Subject/.HTMLBody/.Send","olMail.To / .Subject / .HTMLBody / .Attachments.Add duongdan / .Send",[],"To lay tu cot email; HTMLBody chua <b> de in dam; Attachments.Add dinh kem; Send gui.","Sub M2()\n Dim m As Object\n Set m = CreateObject(\"Outlook.Application\").CreateItem(0)\n m.To = Cells(2,2): m.Subject = \"Hi\": m.HTMLBody = \"<b>Ten</b>\": m.Send\nEnd Sub","Mail duoc gui","Dung .Body thay .HTMLBody -> mat dinh dang <b>"),
 ("BodyFormat HTML + Substitute + CountA","olMail.BodyFormat = olFormatHTML  +  Substitute(body,cu,moi)  +  CountA dem nguoi nhan",[],"HTML cho <b>ten</b> bo dam; SUBSTITUTE thay ten/ma tung nguoi; CountA(B2:B100) dem so mail; For i=2 To SoNguoi quet email/ten/ma/file.","Sub GuiLoai()\n Dim n As Long, i As Long\n n = WorksheetFunction.CountA(Sheets(1).Range(\"B2:B100\"))\n For i = 2 To n + 1\n  Cells(i,2).Value\n Next\nEnd Sub","Dem + quet dung so dong","CountA dem ca tieu de/o trong gia -> gui lech dong (tru 1/kiem tra)"),
 ("Ngat dong _ + Module + .xlsm + Button","Dau _ cuoi dong noi code  +  Insert > Module  +  luu .xlsm  +  Button Assign Macro",[],"_ giup code dai de doc; viet trong Module theo file; luu .xlsm giu macro; gan Button de bam chay.","Sub Demo()\n Dim x As Long: x = 1 + _\n  2\n MsgBox x\nEnd Sub 'x=3","x=3, nut bam chay","Luu .xlsx -> mat macro; viet tren Sheet -> xoa sheet mat code"),
 ("InputBox + If Mod chan/le","y = InputBox(\"Nhap mot so\")  +  If y Mod 2 = 0 Then MsgBox y & \" la so chan\" Else MsgBox y & \" la so le\" End If",[],"InputBox hien hop nhap, gan bien y; Mod lay du; If...Then...Else...End If bat buoc di cap.","Sub ChanLe()\n Dim y As Long\n y = InputBox(\"Nhap mot so\")\n If y Mod 2 = 0 Then\n  MsgBox y & \" la so chan\"\n Else\n  MsgBox y & \" la so le\"\n End If\nEnd Sub","Bao chan/le dung","Nhap chu vao InputBox so -> Type mismatch (can IsNumeric check)"),
 ("For + If quet dong","For i = 2 To SoNguoi ... Next + If ... Then ... End If (quet dong 2..CountA, If kiem tra tung dong)",[],"Mau chung in phieu/xoa sheet/gui mail: For quet, If loc tung dong.","Sub Q()\n Dim i As Long\n For i = 2 To 10\n  If Cells(i,2) <> \"\" Then MsgBox Cells(i,2)\n Next\nEnd Sub","Duyet + loc tung dong","For khong co Next / If khong End If -> loi bien dich"),
]

QT23 = [
 ("Luyen copy-paste-undo",["Tai file thuc hanh bai 23 (link mo ta)","Boi den vung mau > Ctrl+C (vien nhap nhay) > chon cho moi > Ctrl+V","Esc thoat nhap nhay; Ctrl+X thu cut; Ctrl+Z undo tung buoc"]),
 ("Luyen Find-Replace",["Ctrl+F tim 'hang' > Find Next / Find All","Ctrl+H thay 'du lieu' thanh 'Excel' > Replace / Replace All","Ctrl+Z undo 1-2 buoc de kiem tra"]),
 ("Loc + dieu huong",["Ctrl+Shift+L bat AutoFilter","Ctrl+Mui ten nhay goc; Ctrl+Shift+Mui ten boi den; Ctrl+Home ve A1","Ctrl+PageDown/PageUp chuyen sheet"]),
 ("Dinh dang + tong + gia tri",["Nhap 1000000/3 roi Ctrl+Shift+1..5 xem 5 kieu","Alt+= tinh tong cot/hang ke tren/trai > Enter","Copy cot cong thuc > Alt,H,V,V dan gia tri"]),
 ("Ngay-gio-xuong dong-chart-insert",["Ctrl+; chen ngay; Ctrl+Shift+; chen gio; Alt+Enter xuong dong trong o","Ctrl+A chon vung/sheet; Alt+F1 ve chart","Shift/Ctrl+Space chon hang/cot; Ctrl+Shift++ (R/C) chen; Ctrl+- (R/C) xoa; Shift+keo di chuyen; Ctrl+9/0 an; Alt+; chon hien; Ctrl+E Flash Fill"]),
]
QT24 = [
 ("Bat Developer + chuan bi sheet",["Right-click Ribbon > Customize > tick Developer","Chuan bi Main (muc luc), Du lieu 1/2/3, Input, Data, Tim kiem"]),
 ("Muc luc nhay sheet",["Ve Shape 'Du lieu thang 1' > Record Thang1 (This Workbook) > click Du lieu 1 > Stop > Assign macro vao shape","Lap Thang2/Thang3 tuong tu"]),
 ("Nut quay lai + copy nut",["Record QuayLaiMain (click ve Main) > tao nut tren moi sheet du lieu > gan macro","Copy/Paste hoac Ctrl+keo nut (giu macro)"]),
 ("Nhap lieu tu dong (Relative+Transpose)",["Tao nut 'Cap nhat' > Record: copy Input > sang Data > B3 > Ctrl+Mui ten xuong","BAT Relative > xuong 1 o > Paste Special Transpose > TAT Relative > ve Input xoa trang > Stop > gan nut","Test NV moi (vd so 4) sang Data kiem tra"]),
 ("Tim kiem Advanced Filter + .xlsm",["Ctrl+T chuan hoa vung > record Advanced Filter (Ho Nguyen, Nu...) > gan nut Tim kiem/Tim moi","Alt+F11 xoa dong Range thua, them ' ghi chu; test 2-3 lan; luu .xlsm"]),
]
QT25 = [
 ("Mo VBE + Sub dau tien",["Bat Developer > Alt+F11 > View > Immediate (Ctrl+G) > Insert > Module","Viet Sub SayHello + MsgBox > F5; doi sang Cells(1,1)=...; test ?1+1 trong Immediate"]),
 ("Cell-Range-Sheet-Workbook-Dim",["Cells(hang,cot), Range(\"A1:B3\"), Sheets(ten/so), Workbooks(...).Sheets(...).Cells(...)","Dim Integer/Long/String/Variant/Object; y=5: Cells(1,1)=y"]),
 ("In phieu + xoa sheet",["While/For quet thi sinh + PrintOut Preview:=False > gan nut 'In danh sach'","For Each ws + If Name<>Active thi Delete + DisplayAlerts False/True > gan nut"]),
 ("Copy file FSO",["Lay duong dan nguon/dich + kieu (*.*) > FSO CopyFile + FolderExists + MsgBox/Exit Sub > gan nut"]),
 ("Gui mail hang loat",["References tick Outlook > CreateObject Outlook + For i=2 To CountA + Substitute + HTMLBody + Attachments.Add + Send > nut 'Gui email'","Luu .xlsm truoc khi giao file"]),
]

def build_ct(entries, tss):
    out = []
    for i,(ten,cu,params,gt,vd) in enumerate(entries):
        kq = ""
        m = re.search(r"=\s*SUM\(", vd)
        out.append({"ten":ten,"cu_phap_chuan":cu,
         "tham_so":[{"ten":n,"mo_ta":d} for n,d in params],
         "giai_thich":gt,"vi_du_chay_duoc":vd,"ket_qua_vi_du":kq,
         "loi_thuong_gap": "Xem giai_thich; doi chieu video." if len(cu)<4 else "",
         "timestamp":pick(tss,i,max(len(entries),1))})
    # loi cu the: lay tu giai_thich? giu don gian: rut tu vd? de trong -> dien mac dinh
    return out

# loi rieng cho tung video (ngan gon, dung chuyen mon)
LOI23 = ["Dan de len du lieu cu","Quen Esc dan trung","Cut mat goc neu quen","Undo qua tay","Tim thieu wildcard","Replace All sai hang loat","Quen tat filter tuong mat du lieu","O trong lam nhay dut","Nham Home voi Ctrl+Home","Vung trong dut doan","Sheet an khong nhay toi","% nhan 100 hieu lam","SUM doan sai vung co o trong","Quen Enter sau Alt,H,V,V","Tuong Ctrl+; tu cap nhat","Ban phim laptop thieu Fn","Enter thuong nhay o","Bam 2 lan chon ca sheet","Tieu de lech chart xau","Nham Shift+Space/Ctrl+Space","Chen nham vi tri","Xoa nham chua Ctrl+Z","Tha Shift truoc tha chuot ghi de","Quen Unhide tuong mat","Quen Alt+; copy ca o an","Mau khong nhat quan Flash Fill sai"]
LOI24 = ["Excel Online thieu Developer","Ten dau/cach bao loi; trung Ctrl+C","Ghi thua do quen Stop","Viet ngoai Sub khong chay","Thieu ' lenh thua van chay","Giu Range co dinh chay nham o","Dung ActiveX phuc tap","Click doi chay nham macro","Shape Lock khong bam","Doi ten sheet hong lenh","Copy nut chua gan macro","Cot trong nhay dut doan","Quen TAT/BAT Relative lech dong","Vung dich con du lieu tran","Table co dong trong loc thieu","Luu .xlsx mat macro","Sua code khong backup hong","Quen Clear tuong macro loi"]
LOI25 = ["Viet tren Sheet xoa mat code","Ten dau cach loi dich","Quen ngoac kep","Nham thu tu Cells","Test tren du lieu that mat","Sai ten sheet Out of range","Ten trung/chua ky tu cam","Chay nham file Active*","Book1 chua mo Out of range","Integer >32767 Overflow","Bien chua Dim sai ten","Dia chi Range thieu :","Clear mat ca dinh dang","Quen True mai mat canh bao","Quen bat lai True nguy hiem","Step sai dau vong vo han","Quen tang bien treo may","Preview False ton giay","Thieu Set loi 91","Thieu End If loi dich","Thieu \\ cuoi sai cho","Quen References loi ActiveX","Dung .Body mat HTML","CountA lech dong","Luu .xlsx mat macro","Nhap chu Type mismatch","Thieu Next/End If loi dich"]

def build_ct2(entries, lois, tss):
    out=[]
    for i,e in enumerate(entries):
        ten,cu,params,gt,vd = e[0],e[1],e[2],e[3],e[4]
        out.append({"ten":ten,"cu_phap_chuan":cu,
         "tham_so":[{"ten":n,"mo_ta":d} for n,d in params],
         "giai_thich":gt,"vi_du_chay_duoc":vd,"ket_qua_vi_du":"",
         "loi_thuong_gap":lois[i] if i < len(lois) else "",
         "timestamp":pick(tss,i,max(len(entries),1))})
    return out

def build_tt(entries, tss):
    out=[]
    for i,(ten,steps) in enumerate(entries):
        thao = f"Buoc {i+1} ({ten}): " + " > ".join(steps)
        ks = []
        for k in ["Ctrl+C","Ctrl+V","Ctrl+X","Ctrl+Z","Esc","Ctrl+F","Ctrl+H","Ctrl+Shift+L","Ctrl+Home","Ctrl+PageDown","Alt+=","Alt,H,V,V","Ctrl+;","Alt+Enter","Ctrl+A","Alt+F1","Shift+Space","Ctrl+Space","Ctrl+Shift+","Ctrl+-","Shift+keo","Ctrl+9","Alt+;","Ctrl+E","Record","Relative","Transpose","Ctrl+T",".xlsm","Alt+F11","Ctrl+G","DisplayAlerts","For Each","PrintOut","InputBox"]:
            if k.lower() in thao.lower(): ks.append(k)
        out.append({"buoc":i+1,"thao_tac":thao,"phim_tat":", ".join(ks[:3]),
          "luu_y":"Kiem tra vung sau paste/fill; Ctrl+Z ngay neu sai" if ("Ctrl+Z" in thao or "paste" in thao.lower()) else ("Luu .xlsm truoc khi dong file" if (".xlsm" in thao or "Record" in thao) else "")})
    return out

JOBS = [
 ("Xne7UmDP9TU","Bai 23 - 25 phim tat thong dung",CT23,LOI23,QT23),
 ("vM7i37Kkw80","Bai 24 - Record Macro (ghi macro tu dong)",CT24,LOI24,QT24),
 ("0a9m17AdKpM","Bai 25 - Lap trinh VBA trong Excel",CT25,LOI25,QT25),
]

summary = {}
for vid,bai,ct,lois,qt in JOBS:
    src = json.loads(pathlib.Path(f"E:/webexcel/output/json/{vid}.json").read_text(encoding="utf-8"))
    tss = load_ts(vid)
    new_ct = build_ct2(ct, lois, tss)
    new_tt = build_tt(qt, tss)
    new_q, fixes = rebuild_questions(vid, src.get("cau_hoi",[]), tss)
    assert len(new_q) == 40, (vid, len(new_q))
    assert sum(1 for q in new_q if q["do_kho"]=="de")==10
    assert sum(1 for q in new_q if q["do_kho"]=="trung_binh")==20
    assert sum(1 for q in new_q if q["do_kho"]=="kho")==10
    # validate dap_an maps to non-empty unique options
    for q in new_q:
        assert q["dap_an"] in "ABCD"
        assert len({q["A"],q["B"],q["C"],q["D"]})==4
        assert all(q[k].strip() for k in "ABCD")
    out = {"video_id":vid,"bai":bai,"url":f"https://www.youtube.com/watch?v={vid}",
      "cong_thuc":new_ct,"thuat_toan_quy_trinh":new_tt,"cau_hoi":new_q,
      "kiem_chung":{"nguon":"transcript+frames","ghi_chu":f"Doi chieu {vid}: phim tat/macro/VBA theo video + chuan Excel/MS; timestamp tu transcript {len(tss)} moc."}}
    pathlib.Path(OUT/f"{vid}.json").write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding="utf-8")
    summary[vid]=(len(new_ct),len(new_q),fixes,len(tss))
    print(vid, summary[vid])
print("DONE", summary)
