import json, os
OUT=r"E:\webexcel\output\json"
def save(vid,bai,ct,qt,qq):
    d={"video_id":vid,"bai":bai,"url":f"https://www.youtube.com/watch?v={vid}","cong_thuc":ct,"thuat_toan_quy_trinh":qt,"cau_hoi":qq}
    open(os.path.join(OUT,vid+".json"),"w",encoding="utf-8").write(json.dumps(d,ensure_ascii=False,indent=2))
    print("saved",vid,len(qq))
def Q(i,cau,A,B,C,D,da,gt): return {"id":i,"cau":cau,"A":A,"B":B,"C":C,"D":D,"dap_an":da,"giai_thich":gt}
vid="Vf1Lt4gYqn0"; bai="Bai 19 in an trong Excel: Print Preview, Page Setup, Margins, Fit to Page, Page Break"
ct=[
 {"ten":"Mo Print Preview","cong_thuc":"Ctrl+P hoac Quick Access > Print Preview","giai_thich":"Xem truoc in. Gan nut Print Preview len thanh cong cu nhanh de 1 click.","vi_du":"Click mui ten Quick Access > tich Print Preview; Ctrl+P mo preview"},
 {"ten":"Chon may in / Luu PDF","cong_thuc":"Print > Printer: chon may LAN hoac Microsoft Print to PDF","giai_thich":"Chon may in giay hoac in PDF de luu file.","vi_du":"Printer=Print to PDF > Print > dat ten Bai19.pdf > Save ra Desktop"},
 {"ten":"Pham vi in","cong_thuc":"Settings: Print Active Sheets | Entire Workbook | Selection","giai_thich":"Active Sheets: sheet dang lam; Entire Workbook: ca file (vd 36 trang); Selection: chi vung boi den.","vi_du":"Boi den tieu de+10 thu thuat > Print Selection chi in vung do"},
 {"ten":"So trang in","cong_thuc":"Pages: tu [1] den [10]; 1 trang cu the: 7 den 7","giai_thich":"Mac dinh in het; nhap from-to de in trang nhat dinh.","vi_du":"Entire Workbook 36 trang > Pages 1-10 chi in 10 trang dau"},
 {"ten":"Huong giay","cong_thuc":"Settings > Orientation: Portrait (doc) / Landscape (ngang)","giai_thich":"Du lieu ngang nen chon Landscape cho dep va tiet kiem giay.","vi_du":"Bang rong chon Landscape; danh sach dai chon Portrait"},
 {"ten":"Hien le trang in (Show Margins)","cong_thuc":"Preview goc phai duoi > Show Margins > keo cham den","giai_thich":"Hien duong le de keo can truc quan.","vi_du":"Tich Show Margins > keo le tren xuong de tiet kiem trang"},
 {"ten":"Doi inch sang cm","cong_thuc":"File > Options > Advanced > Display > Ruler units: Centimeter","giai_thich":"Mac dinh inch (8.5x14 inch); doi cm de can le theo chuan VN.","vi_du":"Doi xong Page Setup hien cm; Custom Margins nhap 2cm trai/phai"},
 {"ten":"Can le + can giua trang","cong_thuc":"Page Setup > Margins: Top/Bottom/Left/Right + Center: Horizontally/Vertically","giai_thich":"Custom Margins nhap so; tich Center de bang nam giua giay doc + ngang.","vi_du":"Margins 2cm + Center Horizontally + Vertically"},
 {"ten":"Fit Sheet on One Page","cong_thuc":"Settings > No Scaling > Fit Sheet on One Page (hoac Page Setup > Fit to 1 page)","giai_thich":"Co toan bo ve 1 trang. Du lieu qua dai se rat be nen can nhan.","vi_du":"Ke hoach KD 4 trang > Fit Sheet on One Page con 1/1 trang"},
 {"ten":"Page Break Preview keo ve 1 trang","cong_thuc":"View > Page Break Preview > keo duong xanh ve bien","giai_thich":"Xem vung ngat trang (Page1..4, duong xanh); keo duong xanh sat bien de gom 1 trang. Undo: Ctrl+Z.","vi_du":"4 vung xanh > keo phai + keo xuong sat bien > Ctrl+P thay 1/1 trang"}]
qt=[
 {"ten":"Gan nut Print Preview + chon may in/PDF","cac_buoc":["Click mui ten Quick Access Toolbar > tich Print Preview","Click bieu tuong Preview (hoac Ctrl+P)","Printer: chon may LAN hoac Print to PDF","Neu PDF: Print > dat ten > Save"]},
 {"ten":"In selection / sheet / workbook + so trang","cac_buoc":["Boi den vung can in (neu in Selection)","Ctrl+P > Settings: chon Print Selection / Active Sheets / Entire Workbook","Neu nhieu trang: Pages nhap vd 1-10 hoac 7-7","Print"]},
 {"ten":"Chinh huong giay + le + don vi cm","cac_buoc":["Ctrl+P > Orientation: Portrait/Landscape","Tich Show Margins goc duoi > keo le truc quan","Neu can cm: File > Options > Advanced > Ruler units Centimeter > OK","Page Setup > Margins > Custom: nhap cm + Center Horizontally/Vertically"]},
 {"ten":"Fit ve 1 trang (2 cach)","cac_buoc":["C1: Ctrl+P > No Scaling > Fit Sheet on One Page","C2: View > Page Break Preview > keo duong xanh sat bien phai + duoi","Ctrl+P kiem tra 1/1 trang","Luu y: du lieu dai se be, can nhan truoc khi dung"]}]
topics=[
 ("Nut Print Preview gan o dau?","Quick Access Toolbar (mui ten tren cung)",["Thanh cong thuc","Thanh trang thai","Trong o"],"A","1 click mo preview nhanh nhu Save/Undo."),
 ("Phim tat Print Preview/Print la gi?","Ctrl + P",["Ctrl + O","Ctrl + N","Ctrl + S"],"A","Ctrl+P mo preview."),
 ("Chon may in o muc nao?","Printer trong Print Preview",["Orientation","Margins","Copies"],"A","Liet ke may LAN + PDF."),
 ("Khong co may in thi chon gi de demo?","Microsoft Print to PDF",["Khong in duoc","Tat may","Xoa file"],"A","Luu ra PDF thay vi giay."),
 ("In PDF thuc chat la gi?","Luu sheet ra file .pdf",["In giay","Xoa du lieu","Tang dung luong"],"A","Chon PDF > Print > dat ten > Save."),
 ("Mac dinh Settings in gi?","Print Active Sheets (sheet dang lam)",["Ca file","Vung chon","Khong gi"],"A","Chi in sheet hien tai."),
 ("Muon in ca file 3 sheets chon gi?","Print Entire Workbook",["Print Selection","Print Active","Khong in"],"A","Vd 36 trang ca file."),
 ("Muon in chi vung boi den chon gi?","Print Selection",["Entire Workbook","Active Sheets","Khong gi"],"A","Chi in o dang chon."),
 ("In trang 1-10 nhap the nao?","Pages: 1 den 10",["1 den 100","Khong nhap","Nhap chu"],"A","Gioi han from-to."),
 ("In rieng trang 7 nhap the nao?","7 den 7",["7 den 8","1 den 7","0 den 7"],"A","From=To=7 in 1 trang."),
 ("Du lieu ngang nen chon huong gi?","Landscape (ngang)",["Portrait","Vuong","Tron"],"A","Ngang vua bang, dep, tiet kiem."),
 ("Hien duong le de keo o dau?","Show Margins (goc phai duoi preview)",["Show Formula","Show Grid","Show Header"],"A","Hien cham den de keo."),
 ("Mac dinh don vi le la gi?","Inch (vd 8.5x14 inch)",["Cm","Mm","Km"],"A","Dau nhay kep la inch."),
 ("Doi inch sang cm o dau?","File > Options > Advanced > Display > Ruler units: Centimeter",["Insert > Chart","Data > Sort","Home > Font"],"A","De can le chuan VN."),
 ("Chinh so le cu the o dau?","Page Setup > Margins > Custom",["Insert > Picture","View > Zoom","Data > Filter"],"A","Nhap Top/Bottom/Left/Right cm."),
 ("De bang nam giua giay tich gi?","Center Horizontally + Vertically",["Bold","Italic","Wrap"],"A","Can giua trai-phai va tren-duoi."),
 ("Fit Sheet on One Page de lam gi?","Co 4 trang ve 1 trang",["Tang len 10 trang","Xoa du lieu","An sheet"],"A","No Scaling > Fit Sheet."),
 ("Rui ro khi Fit du lieu qua dai?","Chu rat be kho doc",["Mat du lieu","Hong file","Chay nhanh"],"A","Can nhan co nen dung."),
 ("Page Break Preview o the nao?","View > Page Break Preview",["Insert > Break","File > Break","Home > Break"],"A","Hien vung Page1-4 + duong xanh."),
 ("Duong xanh trong Break Preview la gi?","Ran ngat trang",["Duong ke dep","Loi","Virus"],"A","Keo ve bien de gom 1 trang.")]
qq=[];i=1
for (t,c,d,da,gt) in topics:
    qq.append(Q(i,t,c,d[0],d[1],d[2],"A",gt));i+=1
    qq.append(Q(i,f"Thao tac: {t}",c,d[0],d[1],d[2],"A",gt));i+=1
qq=qq[:40]
for k,q in enumerate(qq,1): q["id"]=k
save(vid,bai,ct,qt,qq)

vid="Bukqq7vupEg"; bai="Bai 20 PivotTable trong Excel: Tao Table, Pivot Fields, Values, Layout, PivotChart"
ct=[
 {"ten":"Tao Table truoc Pivot","cong_thuc":"Chon vung > Insert > Table (Ctrl+T) > My table has headers > dat ten (vd Du lieu)","giai_thich":"Table tu mo rong khi them dong; Pivot dung ten Table lam nguon nen tu cap nhat. Ten lien nhau khong dau.","vi_du":"A1:F4998 thanh Table 'Du lieu'; them 21/3/2020 Mien Nam SP01 30 tu vao Table (vung cham cham)"},
 {"ten":"Chen PivotTable","cong_thuc":"Click o trong Table > Insert > PivotTable > New Worksheet > OK","giai_thich":"Nguon = ten Table; dat o sheet moi. Pane PivotTable Fields + Analyze/Design hien khi click vung pivot.","vi_du":"Source=Du lieu > Sheet moi chua pivot trang"},
 {"ten":"Keo Fields: Filters/Columns/Rows/Values","cong_thuc":"Keo Khu vuc>Filters, San pham>Columns, Nha PP>Rows, Doanh thu>Values","giai_thich":"Filters: loc tren; Columns: nhom ngang; Rows: nhom doc; Values: so lieu. Keo ra ngoai hoac bo tich de go.","vi_du":"Filters=Khu vuc (chon Mien Bac), Columns=San pham, Rows=NPP, Values=Sum Doanh thu"},
 {"ten":"Doi cach tinh Values","cong_thuc":"Click mui ten Values > Value Field Settings > Sum/Average/Count/Max...","giai_thich":"So mac dinh Sum; doi Average/Count tuy muc dich.","vi_du":"Sum->Average xem TB; lai Sum de tong"},
 {"ten":"Ty le % tren tong","cong_thuc":"Value Field Settings > Show Values As > % of Grand Total","giai_thich":"Hien ty trong % thay vi so tuyet doi; tong =100%.","vi_du":"SP10 5.24%, SP09 9.16%...; doi ten 'Ty le %' o Formula Bar"},
 {"ten":"Dinh dang so pivot","cong_thuc":"Chuot phai gia tri > Number Format > Number > Use 1000 Separator, Decimal 0-2","giai_thich":"Them phan cach nghin, so le thap phan cho de doc.","vi_du":"1234567 -> 1,234,567; % de 2 thap phan"},
 {"ten":"An/hien +/- , Subtotal, Grand Total","cong_thuc":"Click +/- de thu gon/mo; Design > Subtotals / Grand Totals (Off for Rows&Columns); Show in Tabular Form","giai_thich":"Nut +/- thu gon chi tiet (Mien Bac bao nhieu); Subtotals: tong nhom; Grand Total: tong chung; Tabular: dang bang dep.","vi_du":"Collapse Mien Bac chi thay tong; Grand Totals Off bo dong/cot tong; Show in Tabular Form hien Khu vuc/San pham/Doanh thu"},
 {"ten":"Sap xep + loc trong pivot","cong_thuc":"Click mui ten field > Sort A-Z/Z-A; Filters bo Mien Bac; Ngay > Date Filter chon 3,6","giai_thich":"Sort san pham tang/giam; loc khu vuc/ngay tuy bao cao.","vi_du":"San pham Z-A; Doanh thu NPP sort nho->lon; Ngay chi 3,6"},
 {"ten":"Copy tach pivot + PivotChart","cong_thuc":"Boi den pivot > Copy > paste o khac (PivotTable Fields theo vung moi); Insert > PivotChart > chon kieu","giai_thich":"Copy nguyen pivot de lam bao cao 2 (theo Khu vuc vs theo San pham); PivotChart ve tu pivot, loc Field Buttons.","vi_du":"Copy pivot sang ben phai keo Khu vuc>Rows; PivotChart cot xem Mien Nam 40%, Mien Trung 35%"}]
qt=[
 {"ten":"Chuan bi Table + chen Pivot","cac_buoc":["Click o trong vung ~5000 dong > Insert > Table > tich Headers > OK","Doi kieu hien thi + dat ten 'Du lieu' (Name, lien nhau) > Enter","Them dong test: Table tu mo rong (vung cham)","Click trong Table > Insert > PivotTable > New Worksheet > OK"]},
 {"ten":"Keo fields tao bao cao dau tien","cac_buoc":["Keo Khu vuc vao Filters (loc Mien Bac thu)","Keo San pham vao Columns, Nha PP vao Rows","Keo Doanh thu vao Values (mac dinh Sum)","Dinh dang Number phan cach nghin, giam thap phan"]},
 {"ten":"Tuy bien Values: Average, %, doi ten","cac_buoc":["Keo them Doanh thu vao Values 2-3 lan","Values2: Settings > Average (TB); Values3: Show Values As % Grand Total","Doi ten o Formula Bar: Tong/Trung binh/Ty le","Number Format lai cho tung cot"]},
 {"ten":"Trinh bay: Subtotal/GrandTotal/Layout/Sort","cac_buoc":["Design > Subtotals/Grand Totals: On/Off tuy can","Report Layout > Show in Tabular Form","Sort San pham A-Z/Z-A; sort Doanh thu NPP","Dung +/- thu gon/mo chi tiet Mien Nam..."]},
 {"ten":"Copy pivot + ve PivotChart","cac_buoc":["Boi den ca pivot > Copy > paste o khac (fields di theo)","Vung moi keo lai (vd Khu vuc>Rows de xem theo vung)","Click pivot > Insert > PivotChart > chon cot/trong > OK","Loc Field Buttons chi de Khu vuc; keo rong chart"]}]
topics=[
 ("Du lieu bai 20 lon co nao?","Gan 5000 dong (A1:F4998)",["10 dong","100 dong","5 dong"],"A","Du lieu lon hay doi nen dung Pivot."),
 ("PivotTable manh nhat o gi?","Keo tha phan tich/bao cao/chart du lieu lon",["Go chu","Ve hinh","Nghe nhac"],"A","Khong can cong thuc phuc tap."),
 ("Truoc khi tao Pivot nen lam gi?","Tao Table (Ctrl+T) + dat ten",["Xoa het","Tat may","In ra"],"A","Table tu mo rong, Pivot tu cap nhat."),
 ("Nguon Pivot sau khi tao Table la gi?","Ten Table (vd Du lieu)",["A1 co dinh","So 0","Chuoi rong"],"A","Khong con $A$1:$F$4998 co dinh."),
 ("Ten Table phai the nao?","Lien nhau (Du lieu, khong dau cach)",["Co dau cach","Co ky tu la","Tuy y"],"A","Ten co dau cach gay loi tham chieu."),
 ("Them dong moi vao Table thi sao?","Tu nam vao Table (vung cham)",["Rot ra ngoai","Mat di","Bao loi"],"A","Diem cuoi Table mo rong."),
 ("Chen Pivot o dau?","Insert > PivotTable",["Insert > Picture","Insert > Shape","Insert > Word"],"A","Click trong Table roi Insert."),
 ("Nen dat Pivot o dau?","New Worksheet",["Cung o","Xoa sheet","Khong dat"],"A","Sheet moi sach cho pivot."),
 ("Pane quan ly Pivot ten gi?","PivotTable Fields",["Formula Bar","Status Bar","Title Bar"],"A","Chua Search + Fields + 4 o keo."),
 ("4 o keo trong Fields la gi?","Filters, Columns, Rows, Values",["A,B,C,D","1,2,3,4","X,Y,Z,T"],"A","Bon vung keo tha."),
 ("Keo Khu vuc vao dau de loc?","Filters",["Values","Rows","Columns"],"A","Loc Mien Bac/Nam/Trung tren dau."),
 ("Keo San pham vao dau?","Columns",["Filters","Values","Rows"],"A","Thanh cot SP01-SP10 ngang."),
 ("Keo Nha phan phoi vao dau?","Rows",["Filters","Values","Columns"],"A","Thanh hang doc."),
 ("Keo Doanh thu vao dau?","Values",["Filters","Columns","Rows"],"A","So lieu tinh toan."),
 ("So mac dinh trong Values la gi?","Sum (tong)",["Average","Count","Max"],"A","So thi Sum."),
 ("Doi Sum sang Average o dau?","Value Field Settings",["Font Settings","Print Settings","Page Settings"],"A","Chon Average > OK."),
 ("Xem ty le % chon gi?","Show Values As > % of Grand Total",["Show Nothing","Show Error","Show Blank"],"A","Tong 100%, SP10 5.24%..."),
 ("Doi ten cot pivot lam the nao?","Click ten > Formula Bar go lai > Enter",["Khong doi duoc","Xoa di","Ve lai"],"A","Vd Tong/Trung binh/Ty le."),
 ("So xau chua phan cach nghin thi sao?","Number Format > Use 1000 Separator",["De vay","Xoa so","Nhan doi"],"A","1,234,567 de doc; giam decimal."),
 ("Keo Khu vuc xuong Rows thi thay gi?","Nhom Mien Bac/Nam/Trung + SP con + nut +/-",["Mat het","Bao loi","Tat may"],"A","Xem tong tung vung + chi tiet.")]
qq=[];i=1
for (t,c,d,da,gt) in topics:
    qq.append(Q(i,t,c,d[0],d[1],d[2],"A",gt));i+=1
    qq.append(Q(i,f"Thao tac: {t}",c,d[0],d[1],d[2],"A",gt));i+=1
qq=qq[:40]
for k,q in enumerate(qq,1): q["id"]=k
save(vid,bai,ct,qt,qq)
print("done 5-6")
