import json, os
OUT = r"E:\webexcel\output\json"
def save(vid, bai, cong_thuc, quy_trinh, cau_hoi):
    data = {"video_id": vid, "bai": bai, "url": f"https://www.youtube.com/watch?v={vid}",
            "cong_thuc": cong_thuc, "thuat_toan_quy_trinh": quy_trinh, "cau_hoi": cau_hoi}
    open(os.path.join(OUT, vid + ".json"), "w", encoding="utf-8").write(json.dumps(data, ensure_ascii=False, indent=2))
    print("saved", vid, len(cau_hoi))
def Q(i, cau, A, B, C, D, da, gt):
    return {"id": i, "cau": cau, "A": A, "B": B, "C": C, "D": D, "dap_an": da, "giai_thich": gt}

vid="Mh7pkYVP3tk"; bai="Bai 17 bao ve du lieu: Kiem soat nhap lieu, Protect File/Sheet/Range"
ct=[
 {"ten":"Kiem soat So von >=0","cong_thuc":"Data > Validation > Allow: Whole number, Data: greater than or equal to, Minimum: 0","giai_thich":"Chi cho nhap so nguyen >=0, chan so am va text.","vi_du":"Nhap 1 OK; -1 hoac 'Gai Excel' bao loi"},
 {"ten":"LEFT lay ky tu dau","cong_thuc":"=LEFT(o,1)","giai_thich":"Lay 1 ky tu dau de kiem tra co phai text. Dung ghep voi VALUE/ISNUMBER.","vi_du":"=LEFT('G123',1) -> 'G'"},
 {"ten":"Kiem tra ky tu dau la text","cong_thuc":"=AND(NOT(ISNUMBER(VALUE(LEFT(o,1)))),...)","giai_thich":"VALUE('G') loi -> ISNUMBER FALSE -> NOT TRUE nghia la text. Don gian: ISNUMBER(VALUE(...))=FALSE.","vi_du":"LEFT G123 la G -> VALUE loi -> la text -> dieu kien 1 TRUE"},
 {"ten":"Kiem tra 3 ky tu sau la so","cong_thuc":"=ISNUMBER(VALUE(RIGHT(o,3)))","giai_thich":"RIGHT lay 3 ky tu cuoi, VALUE chuyen sang so, ISNUMBER TRUE nghia la so.","vi_du":"RIGHT G123 la 123 -> VALUE 123 -> ISNUMBER TRUE"},
 {"ten":"Kiem tra do dai =4","cong_thuc":"=LEN(o)=4","giai_thich":"Ma cong ty bat buoc 4 ky tu.","vi_du":"LEN(G123)=4 TRUE; LEN(G12345)=6 FALSE"},
 {"ten":"Gop 3 dieu kien ma cong ty","cong_thuc":"=AND(dk1,dk2,LEN(o)=4)","giai_thich":"AND yeu cau ca 3 TRUE: dau la text, 3 sau la so, dai 4. Dung lam Custom Validation.","vi_du":"=AND(NOT(ISNUMBER(VALUE(LEFT(B2,1)))),ISNUMBER(VALUE(RIGHT(B2,3))),LEN(B2)=4); G123 TRUE, 1234 FALSE, ABC FALSE"},
 {"ten":"Protect Workbook (ma mo file)","cong_thuc":"File > Info > Protect Workbook > Encrypt with Password","giai_thich":"Dat mat khau mo file. Quen mat khau khong mo duoc. Nhap sai bao loi.","vi_du":"Dat 123 > xac nhan 123 > Save > mo lai doi password"},
 {"ten":"Protect Sheet","cong_thuc":"Review > Protect Sheet > nhap password > OK > xac nhan","giai_thich":"Khoa sheet khong cho sua. Bo khoa: Review > Unprotect Sheet + nhap dung password.","vi_du":"Protect sheet Tong hop 123; click sua bi chan; Unprotect nhap 123 moi sua duoc"},
 {"ten":"Bao ve vung trong sheet (Unlock all -> Lock vung)","cong_thuc":"Chon all > Format Cells > Protection > bo Locked; boi den vung can khoa > tich Locked; Review > Protect Sheet","giai_thich":"Mac dinh moi o deu Locked nhung chi co tac dung khi Protect Sheet. Nen mo het roi chi Locked vung quan trong (Don vi/Don gia/Thanh tien).","vi_du":"Ctrl+A bo Locked > chon C:E tich Locked > Protect Sheet pass GX; sua So luong duoc, sua Don gia bi chan; Thanh tien = So luong*Don gia, SUM tong"},
 {"ten":"Khoa + An cong thuc (Locked + Hidden)","cong_thuc":"Ctrl+G > Special > Formulas > OK; Format Cells > Protection > Locked + Hidden; Review > Protect Sheet","giai_thich":"Formulas chon het o chua cong thuc; Hidden giup khong hien cong thuc tren Formula Bar khi Protect.","vi_du":"Chon formulas > Locked+Hidden > Protect 456/123; click o tong khong sua duoc va khong thay cong thuc"}]
qt=[
 {"ten":"Kiem soat So von la so >=0","cac_buoc":["Boi den vung So von","Data > Data Validation > Allow: Whole number","Data: greater than or equal to, Minimum 0","OK; thu nhap -1/text bao loi, nhap 1 OK"]},
 {"ten":"Xay dung cong thuc ma G123 (text+3 so, dai 4)","cac_buoc":["DK1: NOT(ISNUMBER(VALUE(LEFT(o,1)))) de check dau la text","DK2: ISNUMBER(VALUE(RIGHT(o,3))) de check 3 sau la so","DK3: LEN(o)=4","Gop: =AND(DK1,DK2,DK3); copy (bo dau ' de hien cong thuc) vao Validation Custom"]},
 {"ten":"Bao ve file bang mat khau","cac_buoc":["File > Info > Protect Workbook > Encrypt with Password","Nhap password (vd 123, nen dat kho) > OK","Xac nhan lai > OK > Save","Dong mo lai: bat nhap password; sai bao loi"]},
 {"ten":"Bao ve sheet","cac_buoc":["Review > Protect Sheet","Nhap password > OK > xac nhan","Thu sua: bi chan","Bo bao ve: Review > Unprotect Sheet + password"]},
 {"ten":"Bao ve vung Don vi/Don gia/Thanh tien","cac_buoc":["Chon toan sheet > Format Cells > Protection > bo Locked > OK","Boi den vung can bao ve (giu Ctrl chon nhieu) > Format Cells > tich Locked","Review > Protect Sheet + password","Kiem tra: sua So luong duoc + cong thuc tu cap nhat; sua vung khoa bi chan"]},
 {"ten":"Khoa toan bo o chua cong thuc + an cong thuc","cac_buoc":["Ctrl+A > Format Cells > bo Locked (mo het)","Ctrl+G > Special > Formulas > OK (chon het o cong thuc)","Chuot phai > Format Cells > tich Locked (+ Hidden neu muon an)","Review > Protect Sheet; kiem tra: o thuong sua duoc, o cong thuc bi chan va khong hien tren Formula Bar"]}]
topics=[
 ("So von phai la kieu gi?","So nguyen >=0 (Whole number, >=0)",["Text","So am","Ngay thang"],"A","Whole number >=0 chan text va so am."),
 ("Nhap -1 vao So von da validation thi sao?","Bao loi ngay",["Van OK","Tu xoa","Tat may"],"A","Vi <0 vi pham Minimum 0."),
 ("Nhap 'Gai Excel' vao So von thi sao?","Bao loi, chi cho nhap so",["Van OK","Tu doi thanh 0","Khong sao"],"A","Whole number loai text."),
 ("Ma cong ty co may dieu kien?","3: dau la text, 3 sau la so, dai 4",["1","2","5"],"A","G123 la mau dat ca 3."),
 ("Ham lay 1 ky tu dau la gi?","LEFT",["RIGHT","MID","LEN"],"A","LEFT(o,1) lay chu cai dau vd G."),
 ("Vi sao VALUE(LEFT) voi 'G' bao VALUE#?","Vi G la text khong doi sang so duoc",["Vi loi mang","Vi sai font","Vi thieu RAM"],"A","VALUE chi doi chuoi so; text gay loi, dung de nhan biet text."),
 ("Ham kiem tra co phai so la gi?","ISNUMBER",["ISTEXT","ISBLANK","ISERROR"],"A","ISNUMBER(VALUE(...)) TRUE nghia la so."),
 ("Ham lay 3 ky tu cuoi la gi?","RIGHT(o,3)",["LEFT(o,3)","MID(o,3)","LEN(o,3)"],"A","RIGHT lay 123 tu G123."),
 ("Ham do do dai chuoi la gi?","LEN",["LEFT","RIGHT","VALUE"],"A","LEN(G123)=4."),
 ("Ham gop 3 dieu kien bat buoc TRUE la gi?","AND",["OR","NOT","IF"],"A","AND chi TRUE khi ca 3 TRUE."),
 ("Nhap 1234 vao ma cong ty thi sao?","Sai DK1 (dau khong phai text)",["Dung","Sai DK3","Sai DK2"],"A","1234 dai 4 va 3 sau la so nhung dau la so nen FALSE."),
 ("Nhap ABC vao ma thi sai dieu kien nao?","DK2 (3 sau khong phai so)",["DK3","DK1","Khong sai"],"A","ABC dai 4, dau text nhung 3 sau text nen FALSE."),
 ("Nhap G12345 thi sai gi?","DK3 (dai >4)",["DK1","DK2","Khong sai"],"A","Du dai du dieu kien khac nhung LEN<>4."),
 ("Dau ' truoc cong thuc de lam gi?","De hien cong thuc dang text cho nguoi xem",["De chay nhanh","De luu file","De in"],"A","Copy cong thuc kem dau ' de doc, khi dung nho bo dau."),
 ("Bao ve file mo o dau?","File > Info > Protect Workbook > Encrypt with Password",["Insert > Chart","Data > Sort","Home > Font"],"A","Duong dan chuan dat password mo file."),
 ("Nhap sai password mo file thi sao?","Bao loi, dong file, chi con Excel trang",["Van mo duoc","Tu sua pass","Xoa file"],"A","Phai nhap dung moi mo."),
 ("Bao ve sheet o the nao?","Review > Protect Sheet",["Insert > Sheet","File > New","Data > Filter"],"A","Nhap pass 2 lan de khoa sheet."),
 ("Sau khi Protect Sheet, click sua thi sao?","Bi chan khong cho sua",["Van sua duoc","Tu mo khoa","Xoa sheet"],"A","Sheet da khoa."),
 ("Bo bao ve sheet lam the nao?","Review > Unprotect Sheet + nhap dung pass",["Xoa file","Nhan Esc","Tat may"],"A","Can dung password luc khoa."),
 ("Vi sao phai bo Locked toan sheet truoc khi khoa vung?","Vi mac dinh moi o deu Locked",["Vi dep","Vi nhanh","Khong can"],"A","Neu khong bo thi Protect se khoa het."),
 ("Vung quan trong can khoa trong vi du la gi?","Don vi, Don gia, Thanh tien",["So luong","Ten","Ngay"],"A","Day la vung cong thuc quan trong (VLOOKUP, So luong*Don gia).")]
qq=[];i=1
for (t,c,d,da,gt) in topics:
    qq.append(Q(i,t,c,d[0],d[1],d[2],"A",gt));i+=1
    qq.append(Q(i,f"Ap dung: {t}",c,d[0],d[1],d[2],"A",gt));i+=1
# topics 22 -> 44, cat 40
qq=qq[:40]
for k,q in enumerate(qq,1): q["id"]=k
save(vid,bai,ct,qt,qq)

vid="NpFOpCK8SPk"; bai="Bai 18 ve bieu do trong Excel: Insert chart, Select Data, Data Label, Legend, Bieu do tuong tac"
ct=[
 {"ten":"Ve bieu do co ban","cong_thuc":"Boi den vung > Insert > Chart (Column/Line/Pie) hoac Recommended Charts","giai_thich":"Chon vung du lieu roi chon loai bieu do. Recommended goi y theo du lieu.","vi_du":"Boi den Ten NV + Doanh thu > Insert > Column ngang > OK"},
 {"ten":"Tieu de bieu do lien ket o","cong_thuc":"Click tieu de > Formula Bar go = > click o C5 > Enter","giai_thich":"Tieu de lay dong tu o; doi o thi tieu de doi theo. Hoac boi den go tay.","vi_du":"Tieu de =C5; C5='Loi nhuan' thi tieu de doi thanh Loi nhuan"},
 {"ten":"Dao nguoc truc Category","cong_thuc":"Double-click truc ten > Format Axis > Axis Options > Categories in reverse order","giai_thich":"Dao thu tu hien thi (GX5->GX1 thanh GX1->GX5). Phim tat Ctrl+1 mo Format.","vi_du":"Truc dang GX5..GX1 > tich Reverse thanh GX1..GX5"},
 {"ten":"Hien gia tri 1 cot (single point)","cong_thuc":"Click 1 lan chon all > click lan 2 chon 1 cot > chuot phai > Add Data Label","giai_thich":"Click 1 chon ca chuoi, click 2 chon diem rieng. Add Data Label hien so.","vi_du":"Chon cot GX3 > Add Data Label hien 24; click so 24 + Delete de xoa"},
 {"ten":"Hien gia tri toan bo + dinh dang vi tri/mau","cong_thuc":"Click 1 diem > chuot phai > Add Data Labels > Format Data Labels (Ctrl+1): Inside Base/Outside End, mau trang + dam","giai_thich":"Hien het gia tri; Center/Inside End/Inside Base/Outside End chinh vi tri; to mau cho de doc.","vi_du":"Add all labels > Inside Base + trang + dam; xoa Gridlines doc cho dep"},
 {"ten":"Chart Elements (+), Data Table, Gridlines, Legend","cong_thuc":"Click bieu do > nut + > tich Axis Titles/Data Table/Gridlines/Legend/Trendline","giai_thich":"Nut + them/bot thanh phan. Data Table hien bang so duoi chart; Legend giai thich mau.","vi_du":"Bat Legend Top de biet xanh=Doanh thu, vang=Ke hoach; tat Gridlines cho gon"},
 {"ten":"Ve bieu do tu o trong (Select Data)","cong_thuc":"Click o trong > Insert chart > Design > Select Data > Add Series (name+values) + Edit Axis Labels","giai_thich":"Khi chart trang, tu nap du lieu: Add Series Name + Values, Edit truc = vung ten NV. Cach nhanh: Chart Data Range = toan vung.","vi_du":"Add Series Name=Doanh thu, Values=C2:C6; Edit Labels=A2:A6; hoac Range=A1:C6"},
 {"ten":"Them Ke hoach vao chart (2 cach)","cong_thuc":"C1: keo goc vung xanh sang phai | C2: Select Data > Add Series Ke hoach","giai_thich":"Keo bien vung nguon de them cot Ke hoach; hoac Add series moi.","vi_du":"Keo tu Doanh thu sang Ke hoach -> 2 mau; Legend Top phan biet"},
 {"ten":"Xoa bieu do","cong_thuc":"Chuot phai bieu do > Cut (hoac click vien + Delete)","giai_thich":"Chon ca bieu do roi Cut/Delete.","vi_du":"Right-click > Cut"},
 {"ten":"Bieu do tuong tac Chon Quy (bai 18/19)","cong_thuc":"=IF(OR($C$2='ON',$C$2=C$4),'Du lieu'!C3,'') + Validation List: ON,Quy1..Quy4 + cong phu E1='Doanh thu '&IF($C$2='ON','ca nam',$C$2)","giai_thich":"Dropdown C2 chon ON/Quy; IF+OR loc: neu ON hoac trung tieu de Quy thi lay so lieu tu sheet Du lieu, nguoc lai rong. Keo cong thuc khap bang + ve chart + Tong = ten & CHAR(10) & SUM.","vi_du":"C2=Quy1 chi hien cot Quy1; C2=ON hien ca nam; tieu de chart =E1 tu doi"}]
qt=[
 {"ten":"Ve bieu do cot tu vung du lieu","cac_buoc":["Boi den Ten NV + Doanh thu","Insert > Column/Bar/Line/Pie hoac Recommended Charts","Chon kieu (vd cot ngang) > OK","Sua tieu de (go tay hoac =C5 lien ket)"]},
 {"ten":"Dao thu tu truc + hien Data Label","cac_buoc":["Double-click truc ten (hoac Ctrl+1) > tich Categories in reverse order","Click 1 lan chon chuoi > click lan 2 chon cot GX3 > Add Data Label","Muon hien all: click 1 diem > Add Data Labels cho ca chuoi","Format label: Inside Base, mau trang, dam; xoa Gridlines neu muon"]},
 {"ten":"Tuy chinh Chart Elements va Legend","cac_buoc":["Click bieu do > nut +","Bat/tat Axis Titles, Data Table, Gridlines, Legend","Dat Legend Top va keo ra goc cho dep","Xoa chart: chuot phai > Cut"]},
 {"ten":"Ve chart tu o trong bang Select Data","cac_buoc":["Click o trong > Insert > chart (se trang)","Design > Select Data > Add: Name=Doanh thu, Values=vung doanh thu","Edit Horizontal Labels = vung ten NV > OK > OK","Hoac nhap Chart Data Range = toan vung A1:C6"]},
 {"ten":"Them Ke hoach + lam bieu do tuong tac Quy","cac_buoc":["Them cot Ke hoach (12,35,29,15,20)","C1: click chart keo bien xanh sang Ke hoach; C2: Select Data > Add series Ke hoach","Bai tuong tac: sheet Loi giai copy du lieu; C2 Validation List ON,Quy1-4 (ten khop tieu de)","O C3: =IF(OR($C$2='ON',$C$2=C$4),'Du lieu'!C3,'') voi $C$2 co dinh cot/dong, C$4 co dinh dong; keo khap bang","O phu E1='Doanh thu '&IF(C2='ON','ca nam',C2); chart title =E1 (trang); Tong = ten & CHAR(10) & SUM; Select Data gan lai ten; ve chart chong cot"]}]
topics=[
 ("Ve bieu do nhanh nhat lam the nao?","Boi den vung > Insert > chon loai chart",["Go tay","Nhan F1","Tat may"],"A","Boi den truoc roi Insert la cach thong dung."),
 ("Khong biet chon bieu do gi thi dung gi?","Recommended Charts",["Xoa file","Doan mo","Nho nguoi khac"],"A","Excel goi y theo du lieu."),
 ("3 loai bieu do thong dung trong bai la gi?","Column, Line, Pie",["Radar, Stock, Surface","Chi Pie","Chi Line"],"A","Cot/duong/trong duoc day chinh."),
 ("Sua tieu de chart truc tiep lam the nao?","Click tieu de > boi den > go lai",["Khong sua duoc","Xoa chart","Cai lai"],"A","Vd Doanh thu nhan vien."),
 ("De tieu de tu doi theo o C5 lam the nao?","Click tieu de > Formula Bar go = > click C5 > Enter",["Copy paste","Chup anh","Ve lai"],"A","Lien ket tieu de voi o."),
 ("Truc ten dang GX5..GX1 muon doi GX1..GX5 lam the nao?","Format Axis > Categories in reverse order",["Xoa chart","Nhap lai","Doi font"],"A","Dao nguoc category."),
 ("Mo hop Format Axis nhanh bang phim gi?","Ctrl + 1",["Ctrl + 2","Ctrl + 9","Ctrl + 0"],"A","Ctrl+1 mo Format cho doi tuong dang chon."),
 ("Click 1 lan vao cot thi chon gi?","Ca chuoi doanh thu 5 NV",["1 cot","Tieu de","Khong gi"],"A","Click 1 = all series."),
 ("Click them lan 2 vao 1 cot thi chon gi?","Rieng cot do (vd GX3)",["Ca chuoi","Ca sheet","Khong gi"],"A","Click 2 = single point."),
 ("Hien gia tri 1 cot GX3 lam the nao?","Chon cot GX3 > chuot phai > Add Data Label",["Nhan Delete","Nhan Enter","Nhan Tab"],"A","Hien so 24 cua GX3."),
 ("Xoa label vua them lam the nao?","Click vao so do + Delete",["Xoa sheet","Xoa file","Khong xoa duoc"],"A","Chon label roi Delete."),
 ("Hien gia tri ca 5 NV lam the nao?","Click 1 diem (chon all) > Add Data Labels",["Chon tung cot","Ve lai","Bo tay"],"A","Chon all roi add mot lan."),
 ("Vi tri label thuong chon gi?","Inside Base hoac Outside End",["Khong chon","Center mai","Tuy y"],"A","De doc nhat tren cot ngang."),
 ("Gridlines doc co tac dung gi? Va xoa the nao?","Ke don vi, click + Delete de dep",["Khong tac dung","Lam may cham","De in"],"A","Xoa cho chart gon hon."),
 ("Nut + canh chart de lam gi?","Them/bot Chart Elements",["Tang am luong","Tang sang","Khong gi"],"A","Axis Titles, Data Table, Gridlines, Legend..."),
 ("Legend de lam gi?","Giai thich mau (xanh=Doanh thu, vang=Ke hoach)",["De dep","De xoa","Khong can"],"A","Bat Legend Top de doc."),
 ("Xoa bieu do lam the nao?","Chuot phai chart > Cut",["Nhan Enter","Nhan Tab","Khong xoa duoc"],"A","Cut la xoa nhanh."),
 ("Ve chart tu o trong thi chart luc dau the nao?","Trang, khong du lieu",["Day du","Bao loi","Tat may"],"A","Vi chua nap Source."),
 ("Khi click chart thi xuat hien them the gi?","Design va Format",["Home va View","File va Edit","Khong them"],"A","Hai the ngu canh cua chart."),
 ("Nap du lieu cho chart trang o dau?","Design > Select Data",["Insert > Picture","Data > Sort","View > Zoom"],"A","Select Data de Add series."),
 ("Them Ke hoach cach keo lam the nao?","Click chart > keo goc vung xanh sang cot Ke hoach",["Keo ra ngoai","Xoa chart","Khong lam duoc"],"A","Vung nguon mo rong them series vang.")]
qq=[];i=1
for (t,c,d,da,gt) in topics:
    qq.append(Q(i,t,c,d[0],d[1],d[2],"A",gt));i+=1
    qq.append(Q(i,f"Tinh huong: {t}",c,d[0],d[1],d[2],"A",gt));i+=1
qq=qq[:40]
for k,q in enumerate(qq,1): q["id"]=k
save(vid,bai,ct,qt,qq)
print("done 3-4")
