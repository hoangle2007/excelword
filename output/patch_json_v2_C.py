# -*- coding: utf-8 -*-
import json, os, glob
from collections import Counter
OUT = r"E:\webexcel\output\json_v2"

GLOSS = {
 "NpFOpCK8SPk": [
   {"thuat_ngu":"Bieu do cot","tieng_anh":"Column Chart","giai_thich":"Bieu do cot dung so sanh gia tri theo hang muc."},
   {"thuat_ngu":"Truc phu","tieng_anh":"Secondary Axis","giai_thich":"Truc doc phai cho series ti le nho (%)."},
   {"thuat_ngu":"Bieu do ket hop","tieng_anh":"Combo Chart","giai_thich":"Tron Column + Line trong mot chart."},
   {"thuat_ngu":"Nhan du lieu","tieng_anh":"Data Label","giai_thich":"So hien tren moi cot/diem."},
   {"thuat_ngu":"Do rong khe","tieng_anh":"Gap Width","giai_thich":"Khoang cach giua cac nhom cot."},
   {"thuat_ngu":"Dau danh dau","tieng_anh":"Marker","giai_thich":"Diem danh dau tren Line (gach ngang/size)."}],
 "Vf1Lt4gYqn0": [
   {"thuat_ngu":"Xem truoc in","tieng_anh":"Print Preview","giai_thich":"Xem phan trang truoc khi in (Ctrl+P)."},
   {"thuat_ngu":"Can le","tieng_anh":"Margins","giai_thich":"Le tren/duoi/trai/phai + can giua trang."},
   {"thuat_ngu":"Vua mot trang","tieng_anh":"Fit Sheet on One Page","giai_thich":"Co toan bo ve dung 1 trang."},
   {"thuat_ngu":"Ngat trang","tieng_anh":"Page Break","giai_thich":"Ran xanh tach trang in (Insert/Preview)."},
   {"thuat_ngu":"Lap tieu de","tieng_anh":"Print Titles","giai_thich":"Rows to repeat at top moi trang."},
   {"thuat_ngu":"Vung in","tieng_anh":"Print Area","giai_thich":"Gioi han vung duoc in (Set/Clear)."}],
 "Bukqq7vupEg": [
   {"thuat_ngu":"Bang xoay","tieng_anh":"PivotTable","giai_thich":"Keo tha phan tich du lieu lon."},
   {"thuat_ngu":"O gia tri","tieng_anh":"Values","giai_thich":"So lieu tinh Sum/Average/Count/%."},
   {"thuat_ngu":"Nut loc","tieng_anh":"Slicer","giai_thich":"Nut loc truc quan cho Pivot."},
   {"thuat_ngu":"Gom nhom","tieng_anh":"Group","giai_thich":"Gom Ngay ve Thang/Quy/Nam."},
   {"thuat_ngu":"Lam moi","tieng_anh":"Refresh","giai_thich":"Nap du lieu moi tu Table vao Pivot."},
   {"thuat_ngu":"Bieu do xoay","tieng_anh":"PivotChart","giai_thich":"Chart gan Pivot, nhay theo loc."}],
 "HLBpUHYU6CI": [
   {"thuat_ngu":"Bao cao dong","tieng_anh":"Dashboard","giai_thich":"Mot man hinh tong hop nhieu chart + slicer."},
   {"thuat_ngu":"Nut loc","tieng_anh":"Slicer","giai_thich":"Nut loc Nam/Thang/Khu vuc/NPP."},
   {"thuat_ngu":"Ket noi bao cao","tieng_anh":"Report Connections","giai_thich":"Noi 1 slicer toi moi PivotTable."},
   {"thuat_ngu":"Loc Top","tieng_anh":"Value Filter Top 10","giai_thich":"Loc Top/Bottom N theo tong gia tri."},
   {"thuat_ngu":"Nhan du lieu","tieng_anh":"Data Labels","giai_thich":"Ten + so tien tren chart."},
   {"thuat_ngu":"Lam moi","tieng_anh":"Refresh","giai_thich":"Nap thang moi tu Table vao Pivot."}],
}

for vid in ["NpFOpCK8SPk","Vf1Lt4gYqn0","Bukqq7vupEg","HLBpUHYU6CI"]:
    p = os.path.join(OUT, vid + ".json")
    d = json.load(open(p, encoding="utf-8"))
    # quy_trinh don gian tu thuat_toan_quy_trinh (tuong thich A/B)
    simple = []
    for q in d.get("thuat_toan_quy_trinh", []):
        steps = []
        for s in q.get("cac_buoc", []):
            t = s.get("thao_tac","")
            if s.get("phim_tat"): t += f" ({s['phim_tat']})"
            steps.append(t)
        simple.append({"ten": q.get("ten",""), "cac_buoc": steps})
    d["quy_trinh"] = simple
    d["luong"] = "C_truc_quan_bao_cao"
    c = Counter(q.get("do_kho","trung_binh") for q in d["cau_hoi"])
    d["thong_ke"] = {"tong_cau_hoi": len(d["cau_hoi"]), "de": c.get("de",0), "trung_binh": c.get("trung_binh",0), "kho": c.get("kho",0)}
    d["thuat_ngu_chuan"] = GLOSS[vid]
    src = {"NpFOpCK8SPk":"Bai 18","Vf1Lt4gYqn0":"Bai 19","Bukqq7vupEg":"Bai 20","HLBpUHYU6CI":"Bai 22"}[vid]
    d["kiem_chung"] = {"nguon": [f"E:\\webexcel\\output\\json\\{vid}.json", f"E:\\webexcel\\output\\txt\\{vid}.vi.txt"],
        "ghi_chu": f"Deep-verify {src}: doi chieu transcript + JSON goc; chuan hoa cong_thuc co vi_du + loi_hay_gap; sua dap_an sai/cau trung lap; bo sung chu_de/do_kho/timestamp; JSON hop le."}
    # sap xep khoa cho gon
    order = ["video_id","bai","url","luong","cong_thuc","thuat_toan_quy_trinh","quy_trinh","cau_hoi","thong_ke","thuat_ngu_chuan","kiem_chung"]
    d = {k: d[k] for k in order if k in d}
    open(p, "w", encoding="utf-8").write(json.dumps(d, ensure_ascii=False, indent=2))
    print("patched", vid, d["thong_ke"])
print("DONE PATCH")
