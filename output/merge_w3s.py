# -*- coding: utf-8 -*-
import json, os
base = 'E:/webexcel/output'
full = json.load(open(os.path.join(base, 'full_cong_thuc.json'), encoding='utf-8'))
ds = full['cong_thuc'] if isinstance(full, dict) and 'cong_thuc' in full else full
up = {}
for c in ds:
    t = (c.get('ten') or c.get('ham') or '').strip().upper()
    if t and t not in up:
        up[t] = c
W = 'https://www.w3schools.com/excel/'
w3s = [
 dict(ten='AND', nhom='Logic', cu_phap='=AND([logical1],[logical2],...)',
      giai_thich='W3Schools: tra ve TRUE/FALSE dua tren 2+ dieu kien; thuong dung kem IF.',
      vi_du='=AND(B2="Fire",C2>70) | Ket hop IF: =IF(AND(B2="Fire",C2>70),"Yes","No")', url=W+'excel_and.php'),
 dict(ten='AVERAGEIF', nhom='Thong ke co dieu kien', cu_phap='=AVERAGEIF(range,criteria,[average_range])',
      giai_thich='W3Schools: tinh trung binh cua vung theo 1 dieu kien dung/sai; [average_range] tuy chon.',
      vi_du='=AVERAGEIF(B2:B10,E3,C2:C10) — VD: trung binh Speed cua Type "Grass"', url=W+'excel_averageif.php'),
 dict(ten='AVERAGEIFS', nhom='Thong ke co dieu kien', cu_phap='=AVERAGEIFS(average_range,criteria_range1,criteria1,...)',
      giai_thich='W3Schools: trung binh theo 1+ dieu kien (VD Type="Grass" VA Generation=1).',
      vi_du='=AVERAGEIFS(C2:C13,B2:B13,"Grass",D2:D13,1) — trung binh Defense Grass gen 1', url=W+'excel_averageifs.php'),
 dict(ten='COUNTBLANK', nhom='Dem', cu_phap='=COUNTBLANK(range)',
      giai_thich='W3Schools: dem o trong trong vung; huu ich tim o thieu du lieu.',
      vi_du='=COUNTBLANK(C2:C21) — VD dem 8 o trong cot Type 2', url=W+'excel_countblank.php'),
 dict(ten='COUNTIF', nhom='Dem', cu_phap='=COUNTIF(range,criteria)',
      giai_thich='W3Schools: dem o theo 1 dieu kien (so hoac chu). Dung $ khoa vung khi fill.',
      vi_du='=COUNTIF($B$2:$B$21,F6) — VD dem 6 Pokemon Water; 1 Grass', url=W+'excel_countif.php'),
 dict(ten='COUNTIFS', nhom='Dem', cu_phap='=COUNTIFS(criteria_range1,criteria1,...)',
      giai_thich='W3Schools (suy tu COUNTIF): dem theo nhieu dieu kien AND; xem them FILTER/COUNTIFS.',
      vi_du='=COUNTIFS(C2:C21,"Bug",E2:E21,">300") — mo rong tu VD COUNTIF W3Schools', url=W+'excel_countifs.php'),
 dict(ten='DATEDIF', nhom='Ngay thang', cu_phap='=DATEDIF(start_date,end_date,unit)',
      giai_thich='W3Schools: khoang cach 2 ngay theo unit "Y","M","D","MD","YM","YD"; start phai <= end (sai ra #NUM!).',
      vi_du='=DATEDIF(B2,DATE(2026,9,28),"Y") & " years, " & DATEDIF(B2,DATE(2026,9,28),"YM") & " months"', url=W+'excel_datedif.php'),
 dict(ten='FILTER', nhom='Mang dong (dynamic array)', cu_phap='=FILTER(array,include,[if_empty])',
      giai_thich='W3Schools: tra ve cac dong thoa DK (365/2021+); nhieu DK AND dung *, OR dung +; trong -> #CALC!.',
      vi_du='=FILTER(A2:E21,C2:C21="Fire") | AND: =FILTER(A2:E21,(C2:C21="Bug")*(E2:E21>300)) | OR dung +', url=W+'excel_filter_function.php'),
 dict(ten='IFERROR', nhom='Xu ly loi', cu_phap='=IFERROR(value,value_if_error)',
      giai_thich='W3Schools: tra ve value neu khong loi, nguoc lai value_if_error; bat #N/A,#DIV/0!,#VALUE!,#REF!,#NAME?,#NUM!,#NULL!.',
      vi_du='=IFERROR(VLOOKUP(H3,B2:E21,4,FALSE),"Not found") | =IFERROR(C2/B2,"No battles") | IFNA chi bat #N/A', url=W+'excel_iferror.php'),
 dict(ten='IFS', nhom='Logic', cu_phap='=IFS(logical_test1,value_if_true1,[logical_test2,value_if_true2],...)',
      giai_thich='W3Schools: tra ve gia tri cua DK dung dau tien; nhieu DK dung -> lay DK dung dau.',
      vi_du='=IFS(C2>90,"Fast",C2>50,"Normal",C2<=50,"Slow") — xep loai Speed Pokemon', url=W+'excel_ifs.php'),
 dict(ten='MEDIAN', nhom='Thong ke', cu_phap='=MEDIAN(number1,...)',
      giai_thich='W3Schools: gia tri giua cua du lieu (tu sap xep); kieu trung binh mo ta trung tam.',
      vi_du='=MEDIAN(A2:G2) — tra ve gia tri o giua day so', url=W+'excel_median.php'),
 dict(ten='MODE', nhom='Thong ke', cu_phap='=MODE.SNGL(number1,...)',
      giai_thich='W3Schools: so xuat hien nhieu nhat (1 so); VD tra ve 1 (xuat hien 7 lan).',
      vi_du='=MODE.SNGL(B2:E7) — dem so bong xuat hien nhieu nhat', url=W+'excel_mode.php'),
 dict(ten='NPV', nhom='Tai chinh', cu_phap='=NPV(rate,value1,value2,...)',
      giai_thich='W3Schools: gia tri hien tai rong; rate la ty le chiet khau (VD 10%).',
      vi_du='=NPV(10%,B2:K2) — VD ra 377,87 cho dong tien 10 nam', url=W+'excel_npv.php'),
]
trung = 0
moi = []
trung_ten = []
for w in w3s:
    k = w['ten'].upper()
    if k in up:
        up[k]['vi_du_bo_sung'] = w['vi_du'] + ' (nguon W3Schools: ' + w['url'] + ')'
        up[k]['cu_phap_w3s'] = w['cu_phap']
        trung += 1
        trung_ten.append(w['ten'])
    else:
        moi.append(dict(ten=w['ten'], cu_phap_chuan=w['cu_phap'], tham_so=[],
                        giai_thich=w['giai_thich'], vi_du_chay_duoc=w['vi_du'],
                        ket_qua_vi_du='', nguon='W3Schools', url=w['url'], nhom=w['nhom']))
merged = list(ds) + moi
out = dict(tong_cong_thuc=len(merged), trung_lap=trung, moi_tu_web=len(moi), cong_thuc=merged)
json.dump(out, open(os.path.join(base, 'full_cong_thuc_merged.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
nhom = {}
for w in w3s:
    e = dict(ten=w['ten'], cu_phap=w['cu_phap'], giai_thich=w['giai_thich'], vi_du=w['vi_du'], url=w['url'])
    nhom.setdefault(w['nhom'], []).append(e)
json.dump(dict(nguon='W3Schools https://www.w3schools.com/excel/', tong_ham=len(w3s), nhom=nhom),
          open(os.path.join(base, 'web_nhom_W3S.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('TRUNG:', trung, trung_ten, 'MOI:', len(moi), 'TONG:', len(merged))
print('MOI:', ', '.join(m['ten'] for m in moi))
