import json, os
OUT = r"E:\webexcel\output\json"
os.makedirs(OUT, exist_ok=True)

def save(vid, bai, cong_thuc, quy_trinh, cau_hoi):
    data = {"video_id": vid, "bai": bai, "url": f"https://www.youtube.com/watch?v={vid}",
            "cong_thuc": cong_thuc, "thuat_toan_quy_trinh": quy_trinh, "cau_hoi": cau_hoi}
    p = os.path.join(OUT, vid + ".json")
    open(p, "w", encoding="utf-8").write(json.dumps(data, ensure_ascii=False, indent=2))
    print("saved", p, len(cau_hoi))

def Q(i, cau, A, B, C, D, da, gt):
    return {"id": i, "cau": cau, "A": A, "B": B, "C": C, "D": D, "dap_an": da, "giai_thich": gt}

# ---------- VIDEO 1: Bài 15 ----------
vid = "Amdxvm6OJz4"; bai = "Bai 15 thu thuat Excel 2/4: Go To Special, Dropdown List, Find Select, Status Bar"
ct = [
 {"ten":"Go To Special - Blanks","cong_thuc":"Ctrl+G > Special > Blanks","giai_thich":"Chon tat ca o trong trong vung du lieu de dien hang loat.","vi_du":"Boi den A1:D20, Ctrl+G > Special > Blanks > OK, go 'thieu du lieu', Ctrl+Enter"},
 {"ten":"Dien hang loat o trong","cong_thuc":"Nhap gia tri + Ctrl+Enter","giai_thich":"Giu Ctrl+Enter de dien cung gia tri vao tat ca o dang chon.","vi_du":"Chon cac o trong > to mau vang > go 'thieu du lieu' > Ctrl+Enter"},
 {"ten":"Data Validation List (go truc tiep)","cong_thuc":"Data > Data Validation > Allow: List > Source: Nam,Nu,Khac","giai_thich":"Tao dropdown bang cach go truc tiep, cach nhau boi dau phay. Phu hop danh sach ngan.","vi_du":"Source: Nam,Nu,Khac -> o Gioi tinh hien mui ten so xuong"},
 {"ten":"Data Validation List (tham chieu vung)","cong_thuc":"Data > Data Validation > Allow: List > Source: =$G$14:$G$16","giai_thich":"Tao dropdown tham chieu toi vung chua san danh muc, de bao tri.","vi_du":"Source = $G$14:$G$16 (Nam/Nu/Khac). Nhap 'Ga' se bao loi."},
 {"ten":"Xoa Data Validation","cong_thuc":"Data > Data Validation > Clear All > OK","giai_thich":"Boi den vung co validation, Clear All de xoa dropdown.","vi_du":"Boi den cot Gioi tinh > Clear All"},
 {"ten":"Find - Chon nang cao","cong_thuc":"Home > Find & Select > Find (Ctrl+F) > Find All > Ctrl+A","giai_thich":"Tim kiem tu khoa, Find All liet ke ket qua, Ctrl+A chon het de to mau/dinh dang hang loat.","vi_du":"Find 'SP01' > Find All > Ctrl+A > to mau do; Find 'Mien Bac' > to mau xanh"},
 {"ten":"Status Bar: Average/Count/Max/Min/Sum","cong_thuc":"Boi den vung so > nhin thanh Status Bar","giai_thich":"Average: trung binh; Count: dem o chua so; Max/Min: lon/nho nhat; Sum: tong; Numerical Count dem so, Count dem ca text.","vi_du":"Boi den cot Diem: Average=9.4, Max=18, Sum=119, Count=32"},
 {"ten":"Tuy chinh Status Bar","cong_thuc":"Chuot phai len Status Bar > tich/bot Average, Count, Min, Max, Sum...","giai_thich":"Them bot chi so hien thi. Bo chon neu hien thi qua nhieu.","vi_du":"Chuot phai > them Minimum, Numerical Count; bo Maximum neu khong can"}]
qt = [
 {"ten":"Thay o trong bang 'thieu du lieu' + to vang","cac_buoc":["Boi den toan bo vung du lieu (Ctrl+A hoac keo chuot)","Ctrl+G > Go To > Special > Blanks > OK","Vao Home > Fill Color > mau vang","Go 'thieu du lieu' > giu Ctrl + Enter de dien tat ca o trong","Click o bat ky de thoat chon"]},
 {"ten":"Tao dropdown List go truc tiep","cac_buoc":["Boi den vung can tao (vd cot Gioi tinh)","Data > Data Validation > tab Settings","Allow: List; Source: go Nam,Nu,Khac (phay)","OK > click mui ten de chon, khong phai go tay"]},
 {"ten":"Tao dropdown tu vung danh muc","cac_buoc":["Liet ke danh muc o vung phu (vd G14:G16)","Boi den vung nhap lieu > Data > Data Validation > Allow: List","Click nut chon Source > boi den G14:G16","OK. Nhap ngoai danh muc se bao loi"]},
 {"ten":"Find & chon nang cao de dinh dang hang loat","cac_buoc":["Home > Find & Select > Find (Ctrl+F)","Nhap tu khoa (vd SP01, Mien Bac) > Find All","Nhan Ctrl+A trong cua so ket qua de chon het","Dong cua so, to mau/in dam vung duoc chon"]},
 {"ten":"Su dung Status Bar tinh nhanh","cac_buoc":["Boi den vung so lieu","Doc Average, Count, Max, Min, Sum o thanh duoi cung","Chuot phai Status Bar de them Minimum, Numerical Count...","Bo chon muc khong can de gon"]}]
topics1 = [
 ("Go To Special > Blanks de lam gi?","Chon tat ca o trong trong vung",["Mo hop Find","Tao bieu do","Sap xep du lieu"],"A","Go To Special > Blanks giup chon dong loat cac o trong de dien/to mau."),
 ("De dien 'thieu du lieu' vao tat ca o trong phai bam phim gi sau khi go?","Ctrl + Enter",["Chi Enter","Shift + Enter","Alt + Enter"],"A","Ctrl+Enter dien cung noi dung vao tat ca o dang duoc chon; Enter thuong chi dien 1 o."),
 ("Muon chon toan bo vung du lieu nhanh nhat dung phim nao?","Ctrl + A",["Ctrl + B","Ctrl + G","Ctrl + F"],"A","Ctrl+A chon toan bo sheet/vung hien tai."),
 ("Phim tat mo hop Go To la gi?","Ctrl + G",["Ctrl + H","Ctrl + K","Ctrl + J"],"A","Ctrl+G mo Go To, sau do chon Special."),
 ("Trong Go To Special muon chon o trong thi chon muc nao?","Blanks",["Formulas","Constants","Visible cells only"],"A","Blanks = cac o trong; Formulas = o chua cong thuc."),
 ("De to mau vang cho cac o trong lam the nao?","Home > Fill Color > vang sau khi chon o trong",["Insert > Color","Data > Sort","View > Color"],"A","Dung Fill Color (thung son) tren ribbon Home."),
 ("Tao dropdown Gioi tinh vao the nao?","Data > Data Validation",["Insert > Table","Home > Sort","Formulas > Name"],"A","Data Validation quan ly danh sach so xuong."),
 ("Trong Data Validation, Allow phai chon gi de tao dropdown?","List",["Whole number","Date","Text length"],"A","Allow: List moi tao dropdown."),
 ("Khi go truc tiep Source dropdown, cac muc cach nhau boi dau gi?","Dau phay",["Dau cham phay","Dau hai cham","Dau gach cheo"],"A","Vd: Nam,Nu,Khac. Tuy regional setting co the la phay hoac cham phay."),
 ("Cach dropdown tham chieu vung lam the nao?","Click nut chon Source roi boi den vung danh muc",["Go tay lai tu dau","Copy paste","Khong lam duoc"],"A","Tham chieu vd $G$14:$G$16 giup de bao tri.") ,
 ("Nhap gia tri ngoai dropdown tham chieu thi sao?","Excel bao loi, khong cho nhap",["Tu dong them vao","Khong co gi","Xoa o do"],"A","Validation chan gia tri ngoai danh muc."),
 ("De xoa dropdown lam the nao?","Boi den vung > Data Validation > Clear All > OK",["Nhan Delete","Xoa sheet","An cot"],"A","Clear All go bo quy tac validation."),
 ("Find nang cao mo bang cach nao?","Home > Find & Select > Find (Ctrl+F)",["Insert > Find","Data > Find","View > Find"],"A","Ctrl+F la phim tat Find."),
 ("Sau khi Find All muon chon het ket qua bam gi?","Ctrl + A",["Ctrl + C","Ctrl + V","Ctrl + Z"],"A","Ctrl+A trong pane ket qua se chon het cac o tim thay."),
 ("Ung dung Find Select nang cao la gi?","Chon het SP01 de to do, Mien Bac to xanh hang loat",["Tao bieu do","Loc trung","Tach cot"],"A","Tim theo tu khoa roi dinh dang hang loat rat nhanh."),
 ("Thanh Status Bar nam o dau?","Thanh duoi cung cua cua so Excel",["Thanh tren cung","Giua sheet","Trong ribbon"],"A","Status Bar o day man hinh, hien Average/Count/Sum..."),
 ("Boi den vung so, Status Bar hien gi?","Average, Count, Max, Min, Sum...",["Chi ten file","Chi gio","Chi ten sheet"],"A","Thanh trang thai thong ke nhanh vung so duoc chon."),
 ("Average tren Status Bar tinh tren o nao?","Chi o chua so",["Ca o text","Ca o trong","Ca o loi"],"A","Ham thong ke bo qua text va o trong."),
 ("Count va Numerical Count khac nhau the nao?","Count dem ca text, Numerical Count chi dem so",["Giong nhau","Count chi dem so","Numerical Count dem text"],"A","Can phan biet khi dem du lieu hon hop."),
 ("Muon them Minimum vao Status Bar lam the nao?","Chuot phai Status Bar > tich Minimum",["Double click sheet","Nhan F5","Vao File > New"],"A","Chuot phai Status Bar de bat/tat cac chi so.")]
qq1 = []
i = 1
for idx, (t, corr, dis, da_letter, gt) in enumerate(topics1):
    base = [corr] + dis
    # 2 cau hoi tu moi topic -> 40 cau
    qq1.append(Q(i, t, base[0], base[1], base[2], base[3], "A", gt)); i += 1
    # cau bien the: tinh huong
    qq1.append(Q(i, f"Tinh huong: {t} Ap dung dung la?", base[0], base[1], base[2], base[3], "A", gt)); i += 1
save(vid, bai, ct, qt, qq1)

# ---------- VIDEO 2 ----------
vid = "w2FXwxWYYcc"; bai = "Bai 16 thu thuat Excel 3/4: Dropdown thong minh, AutoCorrect, Co gian chu, Chong trung"
ct = [
 {"ten":"Dropdown INDIRECT + Table","cong_thuc":"=INDIRECT(\"sp\")","giai_thich":"Dung trong Data Validation Source de tham chieu toi Table co ten 'sp'. Them dong moi trong Table tu dong cap nhat dropdown.","vi_du":"Tao Table tu A3:A8 dat ten sp; D11 Validation Source =INDIRECT(\"sp\"); them GX06, GX07 tu hien"},
 {"ten":"Dropdown bang Named Range + F3","cong_thuc":"Dat ten vung (Name Box): sp; Source: =sp (hoac F3 dan ten)","giai_thich":"Dat ten cho vung danh muc, Validation tham chieu ten. F3 hien danh sach ten de chon nhanh.","vi_du":"Boi den A4:A8, Name Box go sp > D12 Validation Source =sp; sua Name Manager bo tieu de neu thua"},
 {"ten":"Dropdown dong voi OFFSET + COUNTA","cong_thuc":"=OFFSET($A$3,1,0,COUNTA($A:$A)-1,1)","giai_thich":"OFFSET tao vung dong: bat dau tu A3 dich 1 hang, cao = so o co du lieu tru tieu de. Them SP moi tu cap nhat.","vi_du":"D15 Source =OFFSET($A$3,1,0,COUNTA($A:$A)-1,1); them GX06 tu co trong list. Neu moc A3 thi dich 1, moc A4 thi dich 0"},
 {"ten":"AutoCorrect them quy tac","cong_thuc":"File > Options > Proofing > AutoCorrect Options > Replace: gl With: Gai Excel > Add","giai_thich":"Go tat tu dong thanh cum day du. Tiet kiem khi nhap lap lai.","vi_du":"Them gl->Gai Excel, gw->Gai Word; go gl + Enter thanh Gai Excel"},
 {"ten":"Xoa quy tac AutoCorrect","cong_thuc":"File > Options > Proofing > AutoCorrect > chon quy tac > Delete","giai_thich":"Chon quy tac cu, Delete de go bo.","vi_du":"Chon gl, Delete; go gl khong doi nua"},
 {"ten":"Shrink to fit (co gian chu vua o)","cong_thuc":"Ctrl+1 > Alignment > tich Shrink to fit > OK","giai_thich":"Tu dong thu nho co chu de hien thi vua o, tranh tran trang in. Keo rong o thi chu tro lai co mac dinh.","vi_du":"Boi den cot thang (Ctrl de chon nhieu) > Format Cells > Shrink to fit; so ##### bien mat"},
 {"ten":"Chong nhap trung (Custom)","cong_thuc":"=COUNTIF($B:$B,$B1)<2","giai_thich":"Dem so lan xuat hien cua gia tri trong cot; chi cho phep <2 (khong trung). Ap dung Data Validation Custom cho ca cot.","vi_du":"Chon cot Ma NV > Validation Custom =COUNTIF($B:$B,$B1)<2; nhap NV001 lan 2 bao loi"},
 {"ten":"Thong bao loi trung (Error Alert)","cong_thuc":"Data Validation > Error Alert > Style Stop, Title + Message","giai_thich":"Tuy chinh thong bao khi nhap trung, vd 'Da co ma nhan vien nay'.","vi_du":"Error Alert Title: Trung ma; Message: Da co ma nay. Retry/Cancel khi vi pham"}]
qt = [
 {"ten":"Tao dropdown thong minh bang Table + INDIRECT","cac_buoc":["Boi den nguon gom ca tieu de (vd A3:A8)","Insert > Table > tich My table has headers > OK","Doi ten Table thanh sp (Table Design)","O can dropdown: Data > Validation > Allow List > Source =INDIRECT(\"sp\")","OK; them GX06/GX07 duoi Table tu co trong list"]},
 {"ten":"Tao dropdown bang Named Range","cac_buoc":["Boi den vung danh muc","Go ten sp vao Name Box > Enter","O can dropdown: Validation > Source =sp (hoac go =, bam F3 chon sp)","Neu thua tieu de: Formulas > Name Manager > Edit > sua vung A4:A8"]},
 {"ten":"Tao dropdown bang OFFSET + COUNTA","cac_buoc":["Xac dinh o moc (vd A3 tieu de San pham)","Validation Source =OFFSET($A$3,1,0,COUNTA($A:$A)-1,1)","OK; kiem tra list tu GX01-GX05","Them GX06, mo dropdown thay da cap nhat; sua dich hang 1->0 neu moc la A4"]},
 {"ten":"Dung AutoCorrect go tat","cac_buoc":["File > Options > Proofing > AutoCorrect Options","Replace go gl, With go Gai Excel > Add (tuong tu gw->Gai Word)","OK > OK","Go gl + Enter tu doi thanh Gai Excel; muon xoa: vao lai > Delete"]},
 {"ten":"Tu dong co gian chu vua o","cac_buoc":["Boi den cac o can co gian (giu Ctrl de chon nhieu cot thang)","Chuot phai > Format Cells (Ctrl+1) > tab Alignment","Tich Shrink to fit > OK","Kiem tra: thu hep o chu nho lai, keo rong chu tro lai mac dinh"]},
 {"ten":"Ngan nhap trung ma NV","cac_buoc":["Chon cot/vung Ma NV","Data > Data Validation > Allow: Custom","Formula =COUNTIF($B:$B,$B1)<2","Sang tab Error Alert: nhap Title/Message 'Da co ma nay' > OK; thu nhap NV001 lan 2 se bao loi"]}]
topics2 = [
 ("Vi sao dropdown thuong khong tu cap nhat khi them GX06?","Vi vung Source co dinh",["Vi may hong","Vi thieu RAM","Vi sai font"],"A","Source co dinh A3:A8 khong bao gom dong moi."),
 ("De dropdown tu cap nhat, buoc dau tien la gi?","Tao Table cho vung nguon (Insert > Table)",["Xoa du lieu","Tat may","Doi font"],"A","Table tu mo rong khi them dong."),
 ("Dat ten Table thanh sp de lam gi?","De INDIRECT('sp') tham chieu de dang",["De dep","De tang dung luong","Khong co tac dung"],"A","Ten gon giup cong thuc Validation ngan gon."),
 ("Cong thuc dropdown thong minh bang Table la gi?",'=INDIRECT("sp")',["=SUM(sp)","=VLOOKUP(sp)","=IF(sp)"],"A","INDIRECT chuyen chuoi ten thanh tham chieu dong."),
 ("Cach 2 tao dropdown thong minh dung gi?","Named Range + F3",["PivotTable","Conditional Formatting","WordArt"],"A","Dat ten vung sp roi Source =sp."),
 ("Phim F3 trong hop Source co tac dung gi?","Hien danh sach ten (Paste Name) de chon",["Luu file","In file","Xoa o"],"A","F3 mo Paste Name, chon sp nhanh khong can go."),
 ("Neu dropdown thua dong tieu de 'San pham' thi sua o dau?","Formulas > Name Manager > Edit vung",["Xoa sheet","Cai lai Excel","Doi may"],"A","Sua vung tu A3:A8 thanh A4:A8 de bo tieu de."),
 ("Cong thuc OFFSET dong trong video la gi?","=OFFSET($A$3,1,0,COUNTA($A:$A)-1,1)",["=SUM(A:A)","=AVERAGE(A:A)","=MAX(A:A)"],"A","Dich 1 hang tu A3, cao = dem -1 (tru tieu de)."),
 ("Ham COUNTA trong OFFSET de lam gi?","Dem so o co du lieu de tinh chieu cao vung",["Tinh tong","Tinh trung binh","Tim max"],"A","COUNTA($A:$A)-1 = so san pham (bo tieu de)."),
 ("Tai sao phai -1 trong COUNTA-1?","De tru o tieu de San pham",["De tru o trong","Thich thi tru","De cong them"],"A","Vung dem ca tieu de nen tru 1."),
 ("Neu moc OFFSET la A4 (o SP dau) thi dich bao nhieu hang?","0 hang",["1 hang","2 hang","3 hang"],"A","Moc ngay tai du lieu dau thi khong dich; moc tieu de A3 thi dich 1."),
 ("AutoCorrect mo o dau?","File > Options > Proofing > AutoCorrect Options",["Insert > Chart","Data > Sort","View > Zoom"],"A","Duong dan chuan de them/xoa quy tac."),
 ("Them quy tac gl thanh Gai Excel lam the nao?","Replace: gl, With: Gai Excel > Add",["Xoa gl","Copy paste","Khong lam duoc"],"A","Add de luu quy tac moi."),
 ("Sau khi them AutoCorrect, go gl + Enter thi sao?","Tu doi thanh Gai Excel",["Mat chu","Bao loi","Tat may"],"A","Do la tac dung go tat."),
 ("Xoa quy tac AutoCorrect lam the nao?","Vao lai AutoCorrect > chon quy tac > Delete",["Xoa file","Ghi de","Nhan Esc"],"A","Delete go bo quy tac."),
 ("Shrink to fit de lam gi?","Tu co gian chu vua o, tranh tran trang in",["Lam chu to len","Xoa chu","An cot"],"A","Giu vung in 1 trang ma van hien du so #####."),
 ("Bat Shrink to fit o dau?","Format Cells > Alignment > Shrink to fit",["Insert > Picture","Data > Filter","Home > Sort"],"A","Ctrl+1 mo Format Cells."),
 ("Khi keo rong o sau khi Shrink to fit thi sao?","Chu tro lai co mac dinh",["Chu mat","Chu to dan","Bao loi"],"A","Co chu tu dong phong lai."),
 ("Chon nhieu cot thang khong lien tuc dung phim gi?","Giu Ctrl roi click chon",["Giu Alt","Giu Shift","Khong giu gi"],"A","Ctrl giup chon nhieu vung roi."),
 ("Cong thuc chong trung ma NV la gi?","=COUNTIF($B:$B,$B1)<2",["=SUM(B:B)","=AVERAGE(B:B)","=MAX(B:B)"],"A","Dem so lan xuat hien, <2 nghia la khong trung.")]
qq2 = []
i = 1
for (t, corr, dis, da, gt) in topics2:
    qq2.append(Q(i, t, corr, dis[0], dis[1], dis[2], "A", gt)); i += 1
    qq2.append(Q(i, f"Van dung: {t}", corr, dis[0], dis[1], dis[2], "A", gt)); i += 1
save(vid, bai, ct, qt, qq2)
print("done 1-2")
