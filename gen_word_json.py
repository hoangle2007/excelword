# -*- coding: utf-8 -*-
import json, os
OUT = r"E:\webexcel\output_word\json"
os.makedirs(OUT, exist_ok=True)

def Q(cau, A, B, C, D, da, md, gt):
    return {"cau": cau, "lua_chon": {"A": A, "B": B, "C": C, "D": D}, "dap_an": da, "muc_do": md, "giai_thich": gt}

videos = {}

# 1. cJURPNHyiw0 - can le margins
videos["cJURPNHyiw0"] = {
 "bai": "Cách căn lề chuẩn trong Word (Top-Bottom-Left-Right)",
 "ky_nang": [
  {"ten": "Mở Page Setup căn lề", "thao_tac": "Vào Layout > Margins > Custom Margins để mở hộp thoại Page Setup", "duong_dan_menu": "Layout > Margins > Custom Margins", "phim_tat": "", "giai_thich": "Mọi thao tác căn lề chuẩn đều thực hiện trong hộp thoại Page Setup.", "vi_du": "Mở Custom Margins để nhập Top/Bottom/Left/Right"},
  {"ten": "Nhập kích thước lề chuẩn", "thao_tac": "Nhập Top 2cm, Bottom 2cm, Left 3cm, Right 1.5cm", "duong_dan_menu": "Layout > Margins > Custom Margins > Margins", "phim_tat": "", "giai_thich": "Đây là chuẩn lề văn bản hành chính Việt Nam phổ biến: trên 20-25mm, dưới 20-25mm, trái 30-35mm, phải 10-20mm.", "vi_du": "Top 2cm, Bottom 2cm, Left 3cm, Right 1.5cm"},
  {"ten": "Đổi đơn vị inch sang cm", "thao_tac": "File > Options > Advanced > Display > Show measurements in units of: Centimeters", "duong_dan_menu": "File > Options > Advanced > Display", "phim_tat": "", "giai_thich": "Nếu Word hiển thị inch thì đổi sang cm để nhập lề cho dễ.", "vi_du": "Chuyển inches thành centimeters"},
  {"ten": "Đặt lề mặc định cho mọi tài liệu", "thao_tac": "Trong Page Setup nhấn Set As Default > Yes", "duong_dan_menu": "Layout > Margins > Custom Margins > Set As Default", "phim_tat": "", "giai_thich": "Sau khi đặt lề chuẩn, nhấn Set As Default để mọi file Word mới đều dùng lề này.", "vi_du": "Set As Default > Yes"},
  {"ten": "Tạo tài liệu mới kiểm tra lề", "thao_tac": "Nhấn Ctrl+N mở tài liệu mới và quan sát thước kẻ", "duong_dan_menu": "File > New", "phim_tat": "Ctrl+N", "giai_thich": "Mở file mới để xác nhận lề mặc định đã áp dụng.", "vi_du": "Ctrl+N rồi xem thước đo lề"}
 ],
 "cau_hoi": [
  Q("Lề trên chuẩn theo video là bao nhiêu?", "1cm", "2cm", "3cm", "1.5cm", "B", "dễ", "Video đặt Top = 2cm (cách mép trên 20-25mm)."),
  Q("Lề dưới chuẩn theo video là bao nhiêu?", "2cm", "3cm", "1cm", "2.5cm", "A", "dễ", "Video đặt Bottom = 2cm."),
  Q("Lề trái chuẩn theo video là bao nhiêu?", "2cm", "1.5cm", "3cm", "2.5cm", "C", "dễ", "Video đặt Left = 3cm (30-35mm)."),
  Q("Lề phải chuẩn theo video là bao nhiêu?", "3cm", "2cm", "1cm", "1.5cm", "D", "dễ", "Video đặt Right = 1.5cm (10-20mm)."),
  Q("Muốn căn lề vào menu nào?", "Insert", "Layout", "View", "Review", "B", "dễ", "Vào Layout > Margins."),
  Q("Lệnh mở hộp thoại căn lề tùy chỉnh là gì?", "Custom Margins", "Normal Margins", "Orientation", "Size", "A", "dễ", "Layout > Margins > Custom Margins."),
  Q("Đơn vị nào nên dùng khi căn lề ở Việt Nam?", "Inch", "Point", "Centimet (cm)", "Pica", "C", "dễ", "Đổi sang cm cho dễ nhập 2cm/3cm."),
  Q("Phím tắt tạo tài liệu Word mới là gì?", "Ctrl+O", "Ctrl+S", "Ctrl+N", "Ctrl+P", "C", "dễ", "Ctrl+N tạo file mới để kiểm tra lề mặc định."),
  Q("Nút nào lưu lề thành mặc định cho mọi file sau này?", "OK", "Set As Default", "Cancel", "Print", "B", "dễ", "Nhấn Set As Default > Yes."),
  Q("Video dùng phiên bản Word nào làm mẫu?", "Word 2010", "Word 2021", "Word 2007", "Word 365 web", "B", "dễ", "Video nói đang dùng Word 2021, các bản khác thao tác giống nhau."),
  Q("Nếu Word hiện đơn vị inch, cần làm gì trước khi nhập lề cm?", "Cứ nhập số cm bình thường", "Đổi đơn vị đo sang Centimeters trong Options", "Nhân số inch với 10", "Đổi khổ giấy sang A3", "B", "trung bình", "File > Options > Advanced > Display > Centimeters."),
  Q("Đường vào hộp thoại Page Setup đầy đủ là?", "Layout > Margins > Custom Margins", "Home > Paragraph > Margins", "Insert > Page Setup", "File > Margins", "A", "trung bình", "Đúng đường dẫn trong video."),
  Q("Sau khi nhập kích thước lề xong nhấn gì để áp dụng?", "OK", "Cancel", "Set As Default rồi OK", "Apply rồi Cancel", "C", "trung bình", "Video hướng dẫn Set As Default để lưu mặc định rồi OK."),
  Q("Tại sao lề trái thường lớn nhất (3cm)?", "Để trang trí", "Để chừa chỗ đóng gáy/đóng tập văn bản", "Do lỗi Word", "Để tiết kiệm giấy", "B", "trung bình", "Lề trái 30-35mm để đóng gáy không che chữ."),
  Q("Quy định lề trên/dưới 20-25mm tương ứng bao nhiêu cm?", "0.2-0.25cm", "2-2.5cm", "20-25cm", "1-1.5cm", "B", "trung bình", "20mm = 2cm."),
  Q("Quy định lề phải 10-20mm tương ứng?", "1-2cm, video chọn 1.5cm", "10-20cm", "0.1-0.2cm", "3cm", "A", "trung bình", "10mm=1cm nên 10-20mm là 1-2cm."),
  Q("Kiểm tra lề đã đúng bằng cách nào nhanh nhất?", "Nhìn thước kẻ (ruler) phía trên và bên trái", "Đếm số trang", "Xem tên file", "In ra rồi đo bằng mắt", "A", "trung bình", "Thước ruler hiển thị vị trí lề Top/Left/Right."),
  Q("Set As Default có tác dụng gì?", "Chỉ áp dụng cho file hiện tại", "Áp dụng lề này cho mọi tài liệu mới dựa trên mẫu Normal", "Xóa hết văn bản", "Đổi khổ giấy", "B", "trung bình", "Video nhấn Set As Default > Yes để mọi file sau đều dùng lề chuẩn."),
  Q("Hộp thoại xác nhận sau khi nhấn Set As Default yêu cầu gì?", "Nhấn Yes để xác nhận đổi mẫu mặc định", "Nhập mật khẩu", "Chọn máy in", "Khởi động lại máy", "A", "trung bình", "Word hỏi xác nhận thay đổi Normal template, chọn Yes."),
  Q("Nếu chỉ nhấn OK mà không nhấn Set As Default thì?", "Lề chỉ áp dụng cho tài liệu hiện tại", "Word bị lỗi", "Máy tính tắt", "Mọi file sau cũng đổi", "A", "trung bình", "OK chỉ áp dụng file hiện tại; muốn mặc định phải Set As Default."),
  Q("Lề trái 3cm, phải 1.5cm phản ánh nguyên tắc nào?", "Đối xứng", "Ưu tiên gáy bên trái khi đóng quyển", "Ngẫu nhiên", "Tiết kiệm mực", "B", "trung bình", "Văn bản Việt Nam chừa gáy trái lớn."),
  Q("Muốn kiểm tra lề của file mới cần làm gì?", "Ctrl+N mở file mới và xem Page Setup vẫn giữ 2-2-3-1.5", "Mở file cũ", "Xóa Word cài lại", "Đo bằng thước thật", "A", "trung bình", "Video mở trang Word mới để kiểm tra."),
  Q("Top 2cm + Bottom 2cm trên khổ A4 (29.7cm) còn lại chiều cao vùng in?", "29.7 - 4 = 25.7cm", "29.7cm", "2cm", "33.7cm", "A", "khó", "29.7 - 2 - 2 = 25.7cm."),
  Q("Left 3cm + Right 1.5cm trên khổ A4 rộng 21cm còn lại?", "21 - 4.5 = 16.5cm", "21cm", "4.5cm", "25.5cm", "A", "khó", "21 - 3 - 1.5 = 16.5cm."),
  Q("Nếu đơn vị là inch, 2cm tương đương bao nhiêu inch (1 inch=2.54cm)?", "Khoảng 0.79 inch", "2 inch", "5.08 inch", "0.2 inch", "A", "khó", "2/2.54 ≈ 0.79 inch, vì vậy nên đổi sang cm."),
  Q("Vì sao video khuyên đổi inch sang cm?", "Vì nhập số lẻ inch khó và dễ sai; cm phù hợp chuẩn Việt Nam", "Vì cm đẹp hơn", "Vì inch bị cấm", "Vì máy chỉ cho dùng cm", "A", "khó", "Nhập 0.79 inch dễ nhầm hơn nhập 2cm."),
  Q("Trường hợp nào cần Custom Margins thay vì chọn mẫu có sẵn?", "Khi cần đúng chuẩn 2-2-3-1.5 không có trong mẫu nhanh", "Khi muốn in nhanh", "Khi muốn xóa lề", "Không bao giờ", "A", "khó", "Mẫu nhanh không có đúng số này nên phải tùy chỉnh."),
  Q("Sau Set As Default, file Word mới dựa trên mẫu nào?", "Normal template (Normal.dotm)", "Mẫu web", "Mẫu nhà sản xuất giấy", "Không dựa trên mẫu nào", "A", "khó", "Word lưu lề mặc định vào Normal template."),
  Q("Nếu mở Word trên máy khác, lề mặc định 2-2-3-1.5 có còn không?", "Không, vì lưu trong Normal template của máy đã cài", "Còn, vì lưu trong file", "Còn mãi trên mọi máy", "Mất Word", "A", "khó", "Set As Default chỉ lưu trên máy đó; máy khác phải cài lại."),
  Q("Tổng kiểm tra: bộ lề chuẩn trong video là?", "Top 2 - Bottom 2 - Left 3 - Right 1.5 (cm)", "Đều 1cm", "Đều 2.54cm", "Top 3 - Bottom 3 - Left 2 - Right 2", "A", "khó", "Đúng bộ số video đã đặt và kiểm tra trên file mới."),
 ]
}

# 2. OCKU3op8RXk - font mac dinh
videos["OCKU3op8RXk"] = {
 "bai": "Cài font chữ mặc định trong Word",
 "ky_nang": [
  {"ten": "Mở hộp thoại Font", "thao_tac": "Nhấn Ctrl+Shift+F hoặc Ctrl+D; hoặc Home > nhóm Font > mũi tên mở rộng", "duong_dan_menu": "Home > Font > mũi tên mở rộng", "phim_tat": "Ctrl+Shift+F hoặc Ctrl+D", "giai_thich": "Hộp thoại Font là nơi đặt font mặc định.", "vi_du": "Ctrl+D mở hộp thoại Font"},
  {"ten": "Chọn font mặc định", "thao_tac": "Chọn Font: Arial (hoặc Times New Roman theo yêu cầu)", "duong_dan_menu": "Home > Font > Font", "phim_tat": "", "giai_thich": "Chọn họ chữ dùng cho mọi văn bản mới.", "vi_du": "Font Arial"},
  {"ten": "Chọn kiểu và cỡ chữ", "thao_tac": "Font style: Regular; Size: 11 (hoặc 13/14)", "duong_dan_menu": "Home > Font > Style/Size", "phim_tat": "", "giai_thich": "Video chọn Regular, cỡ 11.", "vi_du": "Regular, 11"},
  {"ten": "Đặt làm mặc định", "thao_tac": "Nhấn Set As Default > chọn All documents based on the Normal template > OK", "duong_dan_menu": "Font dialog > Set As Default", "phim_tat": "", "giai_thich": "Lưu lựa chọn thành mặc định cho mọi file mới.", "vi_du": "Set As Default > All documents"}
 ],
 "cau_hoi": [
  Q("Phím tắt mở hộp thoại Font trong video là gì?", "Ctrl+F", "Ctrl+Shift+F", "Ctrl+G", "Ctrl+L", "B", "dễ", "Video nhấn Ctrl+Shift+F; Ctrl+D cũng mở được."),
  Q("Ngoài phím tắt, mở hộp thoại Font bằng cách nào?", "Mũi tên nhỏ góc nhóm Font thẻ Home", "Nút Save", "Nút Print", "Thanh cuộn", "A", "dễ", "Click mũi tên ở góc nhóm Font."),
  Q("Font mặc định video chọn là gì?", "Calibri", "Arial", "Verdana", "Tahoma", "B", "dễ", "Video chọn Arial."),
  Q("Kiểu chữ video chọn là gì?", "Bold", "Italic", "Regular", "Bold Italic", "C", "dễ", "Kiểu thông thường Regular."),
  Q("Cỡ chữ video chọn là bao nhiêu?", "14", "12", "11", "26", "C", "dễ", "Video chọn cỡ 11."),
  Q("Nút nào lưu font thành mặc định?", "Set As Default", "OK luôn", "Cancel", "Print", "A", "dễ", "Phải nhấn Set As Default."),
  Q("Sau Set As Default phải chọn tùy chọn nào?", "This document only", "All documents based on the Normal template", "No documents", "Web only", "B", "dễ", "Chọn All documents để áp dụng mọi file mới."),
  Q("Mở file Word mới để kiểm tra bằng phím nào?", "Ctrl+N", "Ctrl+O", "Ctrl+W", "Ctrl+Q", "A", "dễ", "Ctrl+N mở file mới thấy Arial 11."),
  Q("Thẻ nào chứa nhóm Font?", "Insert", "Home", "Layout", "View", "B", "dễ", "Nhóm Font nằm thẻ Home."),
  Q("Ctrl+D trong Word có tác dụng gì?", "Mở hộp thoại Font", "Xóa văn bản", "Lưu file", "In file", "A", "dễ", "Ctrl+D mở Font dialog."),
  Q("Vì sao phải nhấn Set As Default thay vì chỉ OK?", "OK chỉ đổi chỗ đang chọn; Set As Default mới lưu cho file mới sau này", "OK làm hỏng máy", "Set As Default để trang trí", "Không khác gì", "A", "trung bình", "Video nhấn mạnh điểm này."),
  Q("Font style Regular nghĩa là gì?", "Chữ thường, không đậm không nghiêng", "Chữ đậm", "Chữ nghiêng", "Chữ gạch chân", "A", "trung bình", "Regular là kiểu thông thường."),
  Q("Size trong hộp thoại Font là gì?", "Cỡ chữ (pt)", "Màu chữ", "Khoảng cách dòng", "Lề giấy", "A", "trung bình", "Size là cỡ chữ."),
  Q("Nếu muốn cỡ 13 cho luận văn thì làm sao?", "Vẫn các bước trên nhưng chọn Size 13 rồi Set As Default", "Không đổi được", "Phải cài lại Word", "Đổi cỡ màn hình", "A", "trung bình", "Chỉ thay số cỡ rồi lưu mặc định."),
  Q("Tùy chọn All documents based on Normal có nghĩa gì?", "Áp dụng cho mọi văn bản mới dùng mẫu Normal", "Chỉ file hiện tại", "Chỉ máy in", "Chỉ tiêu đề", "A", "trung bình", "Normal là mẫu mặc định của Word."),
  Q("Dấu hiệu font mặc định đã đổi thành công?", "Mở file mới thấy Arial 11 sẵn", "Máy kêu", "Màn hình đổi màu", "Mất hết chữ", "A", "trung bình", "Video mở file mới kiểm tra."),
  Q("Muốn mở Font dialog từ thẻ Home thì click đâu?", "Mũi tên chéo góc dưới-bên phải nhóm Font", "Nút B", "Nút U", "Thước kẻ", "A", "trung bình", "Icon launcher góc nhóm Font."),
  Q("Phím Ctrl+Shift+F gồm những phím nào?", "Ctrl + Shift + F", "Ctrl + F", "Shift + F", "Alt + F", "A", "trung bình", "Đúng tổ hợp 3 phím."),
  Q("Nếu chọn This document only thì điều gì xảy ra?", "Chỉ file hiện tại đổi, file mới sau không đổi", "Mọi file đều đổi", "Word lỗi", "Mất font", "A", "trung bình", "Phải chọn All documents mới có tác dụng lâu dài."),
  Q("Font Arial thuộc nhóm nào?", "Font không chân (sans-serif), dễ đọc trên màn hình", "Font chữ thư pháp", "Font biểu tượng", "Font mã vạch", "A", "trung bình", "Arial là sans-serif phổ biến."),
  Q("Cỡ 11pt phù hợp cho?", "Văn bản thông thường, tiết kiệm trang mà vẫn dễ đọc", "Tiêu đề pano", "Chữ ký", "Mã QR", "A", "trung bình", "11pt là cỡ văn bản phổ biến."),
  Q("Sau khi OK ở hộp thoại xác nhận thì nhấn gì nữa?", "OK để đóng hộp thoại Font", "Tắt máy", "Xóa chữ", "In ngay", "A", "trung bình", "Nhấn OK hoàn tất."),
  Q("Mẫu Normal.dotm lưu ở đâu?", "Trong thư mục Templates của user trên máy đó", "Trên mạng", "Trong máy in", "Trong USB", "A", "khó", "Mẫu Normal lưu cục bộ theo user."),
  Q("Đổi font mặc định có ảnh hưởng file cũ không?", "Không, chỉ file mới tạo sau đó", "Có, đổi hết mọi file cũ", "Xóa hết file cũ", "Đổi cả Excel", "A", "khó", "File cũ giữ định dạng đã lưu."),
  Q("Vì sao cùng thao tác dùng được cho Word 2013-2021?", "Vì hộp thoại Font và Normal template không đổi qua các bản", "Vì Microsoft quên cập nhật", "Vì video đoán", "Vì máy giống nhau", "A", "khó", "Các bản Word giữ nguyên cơ chế này."),
  Q("Nếu mở Word thấy vẫn Calibri 11 thì nguyên nhân?", "Chưa nhấn Set As Default hoặc chọn nhầm This document only", "Chưa bật máy", "Chưa có điện", "Do virus chắc chắn", "A", "khó", "Lỗi phổ biến là quên bước lưu mặc định."),
  Q("Regular + 11pt + Arial tạo ra?", "Văn bản thường dễ đọc, chuẩn công sở", "Tiêu đề nghệ thuật", "Chữ đậm nghiêng", "Chữ ẩn", "A", "khó", "Đúng combo video chọn."),
  Q("Muốn font mặc định là Times New Roman 13 thì?", "Làm đúng quy trình, chọn Times New Roman + Regular + 13 + Set As Default All documents", "Không thể", "Phải mua Word mới", "Đổi hệ điều hành", "A", "khó", "Quy trình giống hệt, chỉ đổi giá trị."),
  Q("Thứ tự đúng các bước cài font mặc định?", "Mở Font dialog > chọn Font/Style/Size > Set As Default > All documents > OK > kiểm tra file mới", "OK trước rồi mở dialog", "Tắt máy rồi mở", "In rồi mới cài", "A", "khó", "Đúng trình tự video hướng dẫn."),
  Q("Kết luận video?", "File mới đã sẵn Arial 11 nhờ Set As Default All documents", "Không đổi được", "Phải đổi mỗi lần", "Chỉ đổi được 1 lần", "A", "khó", "Video mở file mới chứng minh thành công."),
 ]
}

# 3. xCdzpYdrVqY - tat gach do
videos["xCdzpYdrVqY"] = {
 "bai": "Tắt gạch chân đỏ kiểm tra chính tả trong Word",
 "ky_nang": [
  {"ten": "Mở Word Options", "thao_tac": "File > Options", "duong_dan_menu": "File > Options", "phim_tat": "", "giai_thich": "Mọi cài đặt kiểm tra chính tả nằm trong Word Options.", "vi_du": "File > Options"},
  {"ten": "Vào mục Proofing", "thao_tac": "Chọn Proofing trong khung Word Options", "duong_dan_menu": "File > Options > Proofing", "phim_tat": "", "giai_thich": "Proofing quản lý spelling/grammar.", "vi_du": "Click Proofing"},
  {"ten": "Bỏ kiểm tra chính tả khi gõ", "thao_tac": "Bỏ tick 'Check spelling as you type' > OK", "duong_dan_menu": "File > Options > Proofing > Check spelling as you type", "phim_tat": "", "giai_thich": "Bỏ tick này thì gạch đỏ biến mất ngay.", "vi_du": "Uncheck Check spelling as you type > OK"}
 ],
 "cau_hoi": [
  Q("Gạch đỏ dưới chữ trong Word là gì?", "Báo lỗi chính tả/ngữ pháp tự động", "Trang trí", "Virus", "Mực in lem", "A", "dễ", "Đó là kiểm tra spelling khi gõ."),
  Q("Muốn tắt gạch đỏ vào menu nào đầu tiên?", "File", "Home", "Insert", "View", "A", "dễ", "File > Options."),
  Q("Trong Word Options chọn mục nào?", "Proofing", "Save", "Language", "General", "A", "dễ", "Proofing quản lý chính tả."),
  Q("Bỏ tick dòng nào để hết gạch đỏ?", "Check spelling as you type", "Save AutoRecover", "Show rulers", "Print hidden text", "A", "dễ", "Bỏ tick dòng này."),
  Q("Sau khi bỏ tick nhấn gì?", "OK", "Cancel", "Delete", "Print", "A", "dễ", "OK để lưu."),
  Q("Video dùng Word bản nào?", "Word 2021", "Word 2003", "Word 97", "WordPad", "A", "dễ", "Video dùng Word 2021."),
  Q("Cách này dùng cho Word 2013/2016/2019 được không?", "Được, giống nhau", "Không", "Chỉ Word 2021", "Chỉ Excel", "A", "dễ", "Video khẳng định các bản giống nhau."),
  Q("Gạch đỏ có in ra giấy không?", "Không, chỉ hiển thị trên màn hình", "Có, in ra luôn", "Tùy máy in", "Có, màu đen", "A", "dễ", "Đây chỉ là gợi ý hiển thị."),
  Q("Muốn bật lại gạch đỏ thì làm gì?", "Tick lại Check spelling as you type", "Cài lại Windows", "Mua máy mới", "Không bật lại được", "A", "dễ", "Tick lại là xong."),
  Q("Số bước chính trong video là mấy?", "2 bước: File>Options>Proofing rồi bỏ tick > OK", "10 bước", "1 bước", "100 bước", "A", "dễ", "Video nói 2 bước."),
  Q("Proofing trong Word Options quản lý gì?", "Spelling và grammar", "Màu nền", "Âm thanh", "Mạng", "A", "trung bình", "Đúng chức năng Proofing."),
  Q("Check spelling as you type nghĩa là?", "Kiểm tra chính tả khi bạn gõ", "Kiểm tra khi in", "Kiểm tra khi lưu", "Kiểm tra virus", "A", "trung bình", "Dịch đúng nghĩa tùy chọn."),
  Q("Vì sao văn bản tiếng Việt hay bị gạch đỏ?", "Vì từ điển Word mặc định tiếng Anh nên từ tiếng Việt bị coi là sai", "Vì gõ sai hết", "Vì máy hỏng", "Vì hết mực", "A", "trung bình", "Nguyên nhân phổ biến nhất."),
  Q("Tắt gạch đỏ có sửa được lỗi chính tả thật không?", "Không, chỉ ẩn gợi ý; lỗi vẫn còn", "Có, sửa hết", "Xóa hết chữ", "Biến thành tiếng Anh", "A", "trung bình", "Tắt hiển thị chứ không sửa lỗi."),
  Q("Ngoài gạch đỏ còn gạch màu nào liên quan?", "Xanh dương (ngữ pháp) và xanh lá tùy bản", "Vàng kim", "Cầu vồng", "Đen", "A", "trung bình", "Word dùng đỏ cho spelling, xanh cho grammar."),
  Q("Muốn ẩn cả gạch xanh ngữ pháp thì bỏ tick nào?", "Check grammar as you type / Mark grammar errors", "Check spelling", "Save", "Zoom", "A", "trung bình", "Tùy chọn ngữ pháp nằm cạnh."),
  Q("File > Options nằm ở đâu?", "Thẻ File góc trái, mục Options cuối danh sách", "Giữa màn hình", "Trên thước kẻ", "Trong văn bản", "A", "trung bình", "Đúng vị trí."),
  Q("Sau khi OK, văn bản trong video thế nào?", "Hết gạch đỏ", "Mất hết chữ", "Đổi màu", "Tăng gấp đôi", "A", "trung bình", "Video cho thấy gạch đỏ biến mất."),
  Q("Cách tắt gạch đỏ nhanh bằng chuột phải?", "Chuột phải vào chữ gạch đỏ > Ignore (chỉ từng chỗ, không triệt để)", "Xóa Word", "Tắt màn hình", "Rút điện", "A", "trung bình", "Ignore chỉ bỏ qua từng từ."),
  Q("Ignore All khác Ignore ở điểm nào?", "Bỏ qua mọi lần xuất hiện của từ đó trong file", "Xóa từ", "Tắt máy", "Không khác", "A", "trung bình", "Ignore All áp dụng toàn file."),
  Q("Add to Dictionary có tác dụng gì?", "Thêm từ vào từ điển để Word không gạch nữa", "Xóa từ điển", "Đổi ngôn ngữ máy", "In từ điển", "A", "trung bình", "Hữu ích cho tên riêng/thuật ngữ."),
  Q("Khi nào NÊN tắt hẳn Check spelling as you type?", "Khi soạn tiếng Việt dài và gạch đỏ gây rối, không cần gợi ý Anh", "Khi thi chính tả", "Khi biên tập sách", "Không bao giờ", "A", "trung bình", "Tắt để đỡ rối mắt khi không cần."),
  Q("Phân biệt tắt hiển thị và tắt kiểm tra?", "Bỏ tick là tắt kiểm tra khi gõ; vẫn có thể F7 kiểm tra thủ công", "Giống nhau", "Không phân biệt được", "Tắt hiển thị là tắt màn hình", "A", "khó", "F7 (Spelling & Grammar) vẫn chạy thủ công."),
  Q("Phím tắt kiểm tra chính tả thủ công là gì?", "F7", "F1", "F12", "F5", "A", "khó", "F7 mở Spelling & Grammar (Review > Spelling & Grammar)."),
  Q("Nếu chỉ muốn ẩn gạch đỏ tạm thời mà vẫn kiểm tra?", "Vẫn có thể dùng Review > Spelling & Grammar (F7) khi cần", "Không thể", "Phải bật lại", "Xóa Word", "A", "khó", "Tắt khi gõ không mất chức năng F7."),
  Q("Vì sao cùng 1 từ lúc bị gạch lúc không?", "Do ngôn ngữ proofing của đoạn đó khác nhau hoặc đã Ignore/Add", "Do ma", "Do mạng", "Do máy in", "A", "khó", "Thuộc tính ngôn ngữ từng đoạn quyết định."),
  Q("Muốn Word bớt gạch sai tiếng Việt mà vẫn giữ kiểm tra tiếng Anh?", "Đặt ngôn ngữ proofing đúng + thêm từ vào từ điển thay vì tắt hẳn", "Xóa hết tiếng Việt", "Viết toàn tiếng Anh", "Tắt máy", "A", "khó", "Giải pháp cân bằng nhất."),
  Q("Ngoại lệ (Exceptions) trong Proofing dùng để?", "Bỏ qua lỗi cho file hiện tại mà vẫn giữ kiểm tra chung", "Xóa file", "Đổi pass", "Tăng tốc mạng", "A", "khó", "Mục Exceptions for khoanh phạm vi."),
  Q("Tóm tắt đúng nhất video?", "File > Options > Proofing > bỏ Check spelling as you type > OK là hết gạch đỏ", "Xóa văn bản là hết", "Tắt màn hình là hết", "Đổi font là hết", "A", "khó", "Đúng quy trình 2 bước."),
  Q("Sau khi tắt, muốn kiểm tra lại 1 lần trước khi nộp bài?", "Nhấn F7 hoặc Review > Spelling & Grammar", "Đoán mò", "In ra đọc", "Nhờ máy in", "A", "khó", "F7 quét toàn văn bản theo yêu cầu."),
 ]
}

# 4. MiiYe2rUBlc - tao bang
videos["MiiYe2rUBlc"] = {
 "bai": "3 cách tạo bảng trong Word",
 "ky_nang": [
  {"ten": "Tạo bảng bằng lưới Insert Table", "thao_tac": "Đặt con trỏ > Insert > Table > quét chọn số cột/dòng (ví dụ 4x3) > click", "duong_dan_menu": "Insert > Table", "phim_tat": "", "giai_thich": "Cách nhanh nhất, kéo lưới chọn cột/dòng.", "vi_du": "Insert > Table > chọn 4 cột 3 dòng"},
  {"ten": "Tạo bảng bằng hộp thoại Insert Table", "thao_tac": "Insert > Table > Insert Table > nhập Number of columns/Rows > OK", "duong_dan_menu": "Insert > Table > Insert Table", "phim_tat": "", "giai_thich": "Nhập chính xác số cột/dòng.", "vi_du": "Columns 4, Rows 3 > OK"},
  {"ten": "Vẽ bảng Draw Table", "thao_tac": "Insert > Table > Draw Table > kéo vẽ khung, hàng, cột, đường chéo", "duong_dan_menu": "Insert > Table > Draw Table", "phim_tat": "", "giai_thich": "Tự vẽ bảng theo kích thước và ô chéo tùy ý.", "vi_du": "Kéo vẽ khung rồi kẻ thêm cột"},
  {"ten": "Di chuyển và thêm dòng bằng Tab", "thao_tac": "Tab chuyển ô tiếp theo; Tab ở ô cuối cùng thêm 1 dòng mới", "duong_dan_menu": "", "phim_tat": "Tab", "giai_thich": "Tab là cách nhập liệu nhanh trong bảng.", "vi_du": "Nhấn Tab để sang ô Họ tên"},
  {"ten": "Định dạng và tô màu bảng", "thao_tac": "Bôi đen > Home định dạng chữ, căn giữa, Shading đổ màu; Table Design > Table Styles", "duong_dan_menu": "Table Design > Table Styles / Home > Shading", "phim_tat": "", "giai_thich": "Làm đẹp tiêu đề và áp style có sẵn.", "vi_du": "Shading màu xanh cho hàng tiêu đề"}
 ],
 "cau_hoi": [
  Q("Muốn tạo bảng vào thẻ nào?", "Home", "Insert", "View", "File", "B", "dễ", "Insert > Table."),
  Q("Cách 1 tạo bảng là gì?", "Quét lưới Table chọn cột/dòng", "Vẽ tay", "Copy Excel", "Chụp ảnh", "A", "dễ", "Cách nhanh nhất trong video."),
  Q("Ví dụ video tạo bảng mấy cột mấy dòng?", "4 cột 3 dòng", "3 cột 4 dòng", "2 cột 2 dòng", "10 cột 10 dòng", "A", "dễ", "Video tạo 4x3."),
  Q("Phím nào chuyển nhanh giữa các ô?", "Enter", "Tab", "Shift", "Alt", "B", "dễ", "Nhấn Tab."),
  Q("Tab ở ô cuối cùng của bảng thì sao?", "Thêm 1 dòng mới", "Xóa bảng", "Tắt Word", "Không gì", "A", "dễ", "Word tự thêm dòng."),
  Q("Muốn đổ màu tiêu đề dùng lệnh nào?", "Shading", "Bold", "Print", "Zoom", "A", "dễ", "Home/Table Design > Shading."),
  Q("Muốn áp mẫu bảng có sẵn vào đâu?", "Table Design > Table Styles", "File > Save", "View > Zoom", "Review > Comment", "A", "dễ", "Table Styles nhiều mẫu đẹp."),
  Q("Muốn đổi độ rộng cột thì làm sao?", "Đưa chuột vào đường biên đến khi thành mũi tên 2 chiều rồi kéo", "Xóa cột", "Tắt máy", "In ra kéo giấy", "A", "dễ", "Kéo biên cột/dòng."),
  Q("Cách 2 tạo bảng là gì?", "Insert Table nhập số cột/dòng > OK", "Vẽ tay", "Gõ chữ", "Chụp màn hình", "A", "dễ", "Hộp thoại Insert Table."),
  Q("Cách 3 tạo bảng là gì?", "Draw Table tự vẽ", "Gõ phím", "Hát", "Vẽ trong Paint", "A", "dễ", "Insert > Table > Draw Table."),
  Q("Trước khi tạo bảng phải làm gì?", "Đặt con trỏ tại vị trí cần tạo bảng", "Tắt máy", "Rút điện", "Xóa hết", "A", "trung bình", "Bảng chèn tại vị trí con trỏ."),
  Q("Trong Insert Table, ô nào nhập số cột?", "Number of columns", "Number of rows", "Table size", "Font size", "A", "trung bình", "Columns là cột."),
  Q("Trong Insert Table, ô nào nhập số dòng?", "Number of rows", "Number of columns", "Page number", "Zoom", "A", "trung bình", "Rows là dòng."),
  Q("Muốn chọn cả bảng nhanh thì click đâu?", "Ô vuông có mũi tên 4 chiều góc trên-trái bảng", "Giữa trang", "Ngoài lề", "Thanh taskbar", "A", "trung bình", "Handle chọn toàn bảng."),
  Q("Muốn căn chữ giữa ô theo chiều dọc dùng nút nào?", "Align Center (căn giữa dọc) trong Layout/Table Tools", "Bold", "Italic", "Underline", "A", "trung bình", "Video click Align Center cho đều trên-dưới."),
  Q("Định dạng chữ trong ô bảng có khác văn bản thường không?", "Không, bôi đen rồi dùng Home như bình thường", "Khác hoàn toàn", "Không định dạng được", "Phải dùng Paint", "A", "trung bình", "Bold/căn giữa dùng như thường."),
  Q("Draw Table phù hợp khi nào?", "Khi cần bảng kích thước tự do, ô chéo, mẫu không chuẩn", "Khi cần bảng chuẩn nhanh", "Khi không cần bảng", "Khi in", "A", "trung bình", "Vẽ tay linh hoạt nhất."),
  Q("Để kẻ thêm 1 cột trong Draw Table thì?", "Click-giữ từ điểm trên xuống điểm dưới rồi thả", "Nhấn Enter", "Nhấn Delete", "Tắt Word", "A", "trung bình", "Kéo 1 đường dọc là thêm cột."),
  Q("Để tạo đường chéo trong ô thì?", "Kéo 1 đường chéo trong ô ở chế độ Draw Table", "Gõ dấu /", "Vẽ bằng bút", "Không thể", "A", "trung bình", "Video vẽ đường chéo mẫu."),
  Q("Muốn thoát chế độ vẽ bảng thì?", "Nhấn Esc hoặc click ra ngoài", "Tắt máy", "Xóa bảng", "Rút điện", "A", "trung bình", "Esc thoát Draw Table."),
  Q("Ưu điểm của Table Styles là?", "Tự tô màu xen kẽ, thêm dòng mới vẫn giữ định dạng", "Làm mất bảng", "Xóa chữ", "Tốn giấy", "A", "trung bình", "Video thêm dòng vẫn giữ style."),
  Q("Nếu không thích style đã áp thì?", "Chọn lại style khác trong Table Styles", "Xóa bảng làm lại", "Cài lại Word", "Bỏ Word", "A", "trung bình", "Đổi mẫu 1 click."),
  Q("Khi nhập liệu nên dùng Tab hay click chuột từng ô?", "Tab nhanh hơn, đỡ nhấc tay khỏi bàn phím", "Click nhanh hơn", "Như nhau", "Không nên nhập", "A", "khó", "Video khuyên dùng Tab."),
  Q("Shift+Tab trong bảng có tác dụng gì?", "Lùi về ô trước đó", "Xóa ô", "Thêm bảng", "Lưu file", "A", "khó", "Ngược với Tab."),
  Q("Muốn hàng tiêu đề lặp lại khi bảng dài nhiều trang?", "Table Properties > Row > Repeat as header row", "Copy tiêu đề mỗi trang", "In riêng", "Không thể", "A", "khó", "Tùy chọn lặp hàng tiêu đề."),
  Q("Muốn gộp 2 ô thành 1 dùng lệnh nào?", "Merge Cells", "Split Cells", "Delete", "Copy", "A", "khó", "Layout > Merge Cells."),
  Q("Muốn tách 1 ô thành nhiều ô dùng lệnh nào?", "Split Cells", "Merge Cells", "Delete Table", "Save", "A", "khó", "Layout > Split Cells."),
  Q("Kéo biên cột mà cả bảng méo theo là vì?", "Đang kéo không giữ phím cố định; muốn đều nên dùng Table Properties đặt Preferred width", "Lỗi Word", "Lỗi chuột", "Lỗi Win", "A", "khó", "Đặt số đo chính xác trong Properties ổn định hơn."),
  Q("So sánh 3 cách tạo bảng?", "Cách 1 nhanh trực quan; cách 2 chính xác số; cách 3 tự do hình dạng", "Cả 3 giống nhau", "Chỉ có 1 cách", "Không cách nào đúng", "A", "khó", "Đúng tổng kết video."),
  Q("Quy trình nhập và làm đẹp bảng hoàn chỉnh?", "Tạo bảng > nhập liệu bằng Tab > bôi đen định dạng Home > Shading/Table Styles", "Shading trước rồi tạo bảng", "In trước rồi nhập", "Xóa rồi tạo", "A", "khó", "Đúng thứ tự video làm."),
 ]
}

# 5. 8AdaaIKlbAo - chia cot columns
videos["8AdaaIKlbAo"] = {
 "bai": "Chia cột văn bản trong Word (Columns)",
 "ky_nang": [
  {"ten": "Chia nhanh 2-3 cột", "thao_tac": "Bôi đen văn bản > Layout > Columns > chọn Two/Three/Left/Right", "duong_dan_menu": "Layout > Columns", "phim_tat": "", "giai_thich": "Chia nhanh thành 2-3 cột hoặc lệch trái/phải.", "vi_du": "Columns > Two"},
  {"ten": "Chia nhiều cột tùy ý", "thao_tac": "Layout > Columns > More Columns > nhập Number of columns > OK", "duong_dan_menu": "Layout > Columns > More Columns", "phim_tat": "", "giai_thich": "Nhập số cột bất kỳ như 4 cột.", "vi_du": "Number of columns = 4 > OK"},
  {"ten": "Kẻ đường giữa các cột", "thao_tac": "Trong Columns dialog tick Line between > OK", "duong_dan_menu": "Layout > Columns > More Columns > Line between", "phim_tat": "", "giai_thich": "Tạo đường kẻ dọc phân cách cột.", "vi_du": "Tick Line between"},
  {"ten": "Chỉnh khoảng cách cột", "thao_tac": "Trong Columns dialog chỉnh Spacing (ví dụ 1cm) > OK", "duong_dan_menu": "Layout > Columns > More Columns > Spacing", "phim_tat": "", "giai_thich": "Spacing càng nhỏ cột càng sát nhau.", "vi_du": "Spacing 1cm"},
  {"ten": "Ngắt sang cột tiếp theo", "thao_tac": "Đặt con trỏ > Layout > Breaks > Column", "duong_dan_menu": "Layout > Breaks > Column", "phim_tat": "", "giai_thich": "Đẩy nội dung từ vị trí con trỏ sang cột bên phải.", "vi_du": "Breaks > Column"}
 ],
 "cau_hoi": [
  Q("Muốn chia cột phải bôi đen văn bản trước đúng không?", "Đúng", "Sai", "Tùy", "Không bao giờ", "A", "dễ", "Video bôi đen nội dung cần chia trước."),
  Q("Lệnh chia cột nằm ở thẻ nào?", "Home", "Layout", "Insert", "View", "B", "dễ", "Layout > Columns."),
  Q("Muốn chia 2 cột chọn gì?", "Two", "Three", "One", "Four", "A", "dễ", "Two là 2 cột."),
  Q("Muốn chia 3 cột chọn gì?", "Two", "One", "Three", "Left", "C", "dễ", "Three là 3 cột."),
  Q("Muốn chia 4 cột thì làm sao?", "More Columns > nhập 4 > OK", "Không chia được", "Nhấn Enter 4 lần", "Copy 4 lần", "A", "dễ", "Nhập Number of columns = 4."),
  Q("Ô nhập số cột tên gì?", "Number of columns", "Spacing", "Line between", "Width", "A", "dễ", "Đúng tên ô."),
  Q("Muốn có đường kẻ giữa các cột tick gì?", "Line between", "Equal width", "Spacing", "OK", "A", "dễ", "Line between tạo kẻ dọc."),
  Q("Muốn chỉnh khoảng cách cột chỉnh ở đâu?", "Spacing", "Line between", "Number", "OK", "A", "dễ", "Width and spacing > Spacing."),
  Q("Muốn đẩy chữ sang cột bên phải dùng gì?", "Breaks > Column", "Enter", "Tab", "Space", "A", "dễ", "Layout > Breaks > Column."),
  Q("Trước khi ngắt cột phải làm gì?", "Đặt con trỏ đúng vị trí muốn ngắt", "Bôi đen cả bài", "Xóa bài", "Tắt máy", "A", "dễ", "Break chèn tại con trỏ."),
  Q("Left trong Columns nghĩa là?", "Cột trái hẹp hơn cột phải", "Cột trái rộng hơn", "2 cột bằng nhau", "1 cột", "A", "trung bình", "Left = narrow left."),
  Q("Right trong Columns nghĩa là?", "Cột phải hẹp hơn cột trái", "Cột phải rộng hơn", "2 cột bằng nhau", "1 cột", "A", "trung bình", "Right = narrow right."),
  Q("More Columns dùng khi nào?", "Khi cần trên 3 cột hoặc chỉnh line/spacing chi tiết", "Khi chia 2 cột", "Khi xóa cột", "Khi in", "A", "trung bình", "Hộp thoại chi tiết."),
  Q("Video giảm spacing xuống bao nhiêu?", "1cm", "5cm", "10cm", "0cm", "A", "trung bình", "Video giảm xuống 1cm cho khít."),
  Q("Spacing lớn thì cột thế nào?", "Xa nhau, mỗi cột hẹp lại", "Sát nhau", "Mất cột", "Tăng số cột", "A", "trung bình", "Khoảng trắng giữa cột tăng."),
  Q("Muốn về 1 cột bình thường chọn gì?", "One", "Two", "Three", "Four", "A", "trung bình", "One gỡ chia cột."),
  Q("Nếu không bôi đen mà chia cột thì?", "Áp dụng cho toàn section/file hiện tại", "Không gì xảy ra", "Máy hỏng", "Mất chữ", "A", "trung bình", "Word áp dụng phạm vi con trỏ/section."),
  Q("Line between là đường nào?", "Đường kẻ dọc giữa các cột", "Đường viền trang", "Gạch chân chữ", "Khung ảnh", "A", "trung bình", "Đúng mô tả."),
  Q("Column break khác Page break ở điểm nào?", "Column nhảy sang cột kế; Page nhảy sang trang mới", "Giống nhau", "Column xóa trang", "Page xóa cột", "A", "trung bình", "Phạm vi nhảy khác nhau."),
  Q("Muốn cột rộng không bằng nhau thì?", "Bỏ tick Equal column width rồi nhập Width từng cột", "Không thể", "Phải vẽ tay", "Đổi khổ giấy", "A", "trung bình", "Bỏ Equal width mới chỉnh riêng."),
  Q("Sau khi chia 2 cột, thêm chữ thì chữ chảy thế nào?", "Đầy cột trái rồi tự chảy sang cột phải", "Tràn ra ngoài", "Mất chữ", "Xuống trang mới", "A", "trung bình", "Luồng cột tự động."),
  Q("Muốn xóa Column break thì?", "Bật dấu ¶ (Show/Hide) rồi xóa ký tự Column Break", "Không xóa được", "Xóa cả bài", "Cài lại Word", "A", "trung bình", "Break là ký tự ẩn, xóa như chữ."),
  Q("Chia cột báo/tạp chí thường dùng mấy cột?", "2-3 cột kèm Line between", "100 cột", "1 cột", "0 cột", "A", "khó", "Bố cục báo phổ biến."),
  Q("Vì sao phải bôi đen đoạn nhân bản cho dài trước khi demo?", "Để thấy rõ hiệu ứng chia cột và chữ chảy", "Để tốn giấy", "Để máy chậm", "Không lý do", "A", "khó", "Đoạn ngắn khó thấy chia cột."),
  Q("Equal column width nghĩa là?", "Các cột rộng bằng nhau", "Các cột khác nhau", "Xóa cột", "Thêm cột", "A", "khó", "Tick mặc định cột đều nhau."),
  Q("Đơn vị Spacing trong video là?", "cm (đã đổi đơn vị đo)", "inch", "pixel", "kg", "A", "khó", "Video nhập 1cm."),
  Q("Muốn áp chia cột cho 1 đoạn giữa trang thì?", "Bôi đen đúng đoạn đó rồi chia; Word tự chèn section break", "Không thể", "Phải chia cả file", "Xóa đoạn khác", "A", "khó", "Word chèn Continuous break bao quanh."),
  Q("Section break liên quan gì đến Columns?", "Mỗi section có thiết lập cột riêng; chia 1 đoạn tạo section mới", "Không liên quan", "Xóa section", "Section là cột", "A", "khó", "Columns là thuộc tính section."),
  Q("Muốn đường kẻ + spacing 1cm + ngắt cột đúng chỗ thì thứ tự?", "Chia cột > More Columns tick Line between + Spacing 1cm > đặt trỏ Breaks Column", "Break trước rồi chia", "In trước rồi chia", "Xóa rồi chia", "A", "khó", "Đúng quy trình video."),
  Q("Tổng kết video gồm mấy nội dung?", "Cách chia 2/nhiều cột + line between + spacing + column break", "Chỉ chia 2 cột", "Chỉ in", "Chỉ gõ chữ", "A", "khó", "Đủ 2 phần video nêu."),
 ]
}

# 6. tFoYWXq7oWA - so mu
videos["tFoYWXq7oWA"] = {
 "bai": "Viết số mũ và ký hiệu độ trong Word",
 "ky_nang": [
  {"ten": "Bật/tắt số mũ bằng phím tắt", "thao_tac": "Nhấn Ctrl+Shift+= để bật superscript, gõ mũ, nhấn lại để về thường", "duong_dan_menu": "", "phim_tat": "Ctrl+Shift+=", "giai_thich": "Cách nhanh nhất viết 2 mũ x-1.", "vi_du": "Gõ 2, Ctrl+Shift+=, gõ x-1, Ctrl+Shift+="},
  {"ten": "Bật số mũ bằng nút x²", "thao_tac": "Home > nút Superscript (x²), gõ mũ, nhấn lại để tắt", "duong_dan_menu": "Home > Superscript (x²)", "phim_tat": "", "giai_thich": "Nút x² trên ribbon.", "vi_du": "Click x² rồi gõ mũ"},
  {"ten": "Số mũ bằng hộp thoại Font", "thao_tac": "Ctrl+D > tick Superscript > OK; bỏ tick để về thường", "duong_dan_menu": "Home > Font dialog > Superscript", "phim_tat": "Ctrl+D", "giai_thich": "Cách 3, nhiều thao tác hơn.", "vi_du": "Tick Superscript > OK"},
  {"ten": "Chèn ký hiệu độ ° bằng mẹo", "thao_tac": "Tìm 'độ C' trên Google/copy ký tự ° rồi dán vào Word", "duong_dan_menu": "", "phim_tat": "", "giai_thich": "Không cần nhớ mã, copy ° dán vào.", "vi_du": "Copy ° dán thành 25°C"}
 ],
 "cau_hoi": [
  Q("Phím tắt viết số mũ trong video là gì?", "Ctrl+Shift+=", "Ctrl+B", "Ctrl+I", "Ctrl+Z", "A", "dễ", "Tổ hợp Ctrl+Shift+dấu bằng."),
  Q("Ví dụ video viết là gì?", "2 mũ x-1", "2+2", "ax+b", "100 độ", "A", "dễ", "Video viết 2^(x-1)."),
  Q("Sau khi gõ phần mũ muốn về chữ thường thì?", "Nhấn lại Ctrl+Shift+=", "Nhấn Enter", "Tắt máy", "Xóa đi", "A", "dễ", "Nhấn lại để tắt superscript."),
  Q("Nút số mũ trên ribbon ký hiệu gì?", "x²", "x2", "2x", "X", "A", "dễ", "Nút Superscript hình x²."),
  Q("Nút x² nằm ở thẻ nào?", "Home nhóm Font", "Insert", "Layout", "View", "A", "dễ", "Cạnh Bold/Italic."),
  Q("Mở hộp thoại Font bằng phím nào?", "Ctrl+D", "Ctrl+F", "Ctrl+G", "Ctrl+H", "A", "dễ", "Ctrl+D hoặc Ctrl+Shift+F."),
  Q("Trong Font dialog tick ô nào để thành số mũ?", "Superscript", "Subscript", "Bold", "Italic", "A", "dễ", "Superscript là chỉ số trên."),
  Q("Subscript là gì?", "Chỉ số dưới (ví dụ H2O)", "Chỉ số trên", "Chữ đậm", "Chữ nghiêng", "A", "dễ", "Ngược với superscript."),
  Q("Ký hiệu độ C viết thế nào?", "Số + ° + C (25°C)", "Số + C", "Số + o + C", "Số + * + C", "A", "dễ", "Dùng ký tự °."),
  Q("Mẹo chèn ° trong video là gì?", "Google 'độ C' rồi copy ° dán vào", "Vẽ tay", "Chụp ảnh", "Đoán", "A", "dễ", "Cách mẹo nhanh không cần nhớ mã."),
  Q("Khi bật superscript, dấu nháy chuột thế nào?", "Nhảy lên phía trên, cỡ nhỏ lại", "To ra", "Biến mất", "Xuống dưới", "A", "trung bình", "Dấu hiệu superscript đang bật."),
  Q("Vì sao phải tắt superscript sau khi gõ mũ?", "Không tắt thì chữ tiếp theo cũng thành mũ", "Máy hỏng", "Hết pin", "Mất file", "A", "trung bình", "Phải nhấn lại để về thường."),
  Q("Ctrl+Shift+= nhấn lần 2 có tác dụng gì?", "Tắt superscript về chữ thường", "Xóa chữ", "Lưu file", "In file", "A", "trung bình", "Bật/tắt luân phiên."),
  Q("Cách nút x² và phím tắt khác nhau ở điểm nào?", "Cùng kết quả; nút dùng chuột, phím tắt nhanh hơn", "Khác kết quả", "Nút chỉ trang trí", "Phím tắt xóa chữ", "A", "trung bình", "Cả hai đều bật superscript."),
  Q("Cách Font dialog phức tạp hơn vì sao?", "Nhiều click: mở dialog > tick > OK, rồi lặp lại để tắt", "Phải gõ code", "Phải khởi động lại", "Phải trả phí", "A", "trung bình", "Video nhận xét vậy."),
  Q("Muốn sửa mũ đã gõ thì?", "Bôi đen phần mũ rồi tắt/bật superscript hoặc gõ lại", "Không sửa được", "Xóa cả dòng", "Gõ đè ngoài", "A", "trung bình", "Bôi đen rồi đổi định dạng."),
  Q("Viết H2O (số 2 dưới) dùng gì?", "Subscript", "Superscript", "Bold", "Underline", "A", "trung bình", "Chỉ số dưới là subscript."),
  Q("Phím tắt subscript là gì?", "Ctrl+= (Ctrl + dấu bằng)", "Ctrl+Shift+=", "Ctrl+B", "Ctrl+S", "A", "trung bình", "Ctrl+= bật subscript."),
  Q("Copy ° từ Google rồi làm gì?", "Paste (Ctrl+V) vào Word rồi chỉnh cỡ cho hợp", "In ra dán giấy", "Chụp ảnh", "Ghi nhớ", "A", "trung bình", "Ctrl+V dán."),
  Q("Ngoài °, mẹo copy còn dùng cho ký tự nào?", "Mọi ký tự đặc biệt khó nhớ (π, ±, √...)", "Chỉ °", "Không dùng được", "Chỉ số", "A", "trung bình", "Mẹo chung cho ký tự đặc biệt."),
  Q(" superscript thực chất là gì?", "Định dạng đẩy chữ lên trên và thu nhỏ, không phải chèn ảnh", "Chèn ảnh", "Đổi font", "Xoay chữ", "A", "trung bình", "Nên copy/bỏ định dạng dễ."),
  Q("Bôi đen chữ mũ rồi nhấn Clear Formatting thì?", "Mất superscript về chữ thường", "Mất chữ", "Đậm hơn", "To hơn", "A", "trung bình", "Clear xóa định dạng trên."),
  Q("Viết 2^(x-1) đúng thứ tự phím?", "Gõ 2 > Ctrl+Shift+= > gõ x-1 > Ctrl+Shift+= > gõ tiếp", "Gõ mũ trước", "Nhấn phím sau khi tắt máy", "Gõ tất cả rồi tô màu", "A", "khó", "Đúng quy trình video."),
  Q("Nếu quên tắt superscript và gõ tiếp 1 đoạn dài?", "Cả đoạn thành mũ; khắc phục: bôi đen > tắt superscript", "Mất đoạn đó", "Máy treo", "Không sao", "A", "khó", "Bôi đen rồi bỏ Superscript."),
  Q("Trong Font dialog, Superscript nằm nhóm nào?", "Effects", "Font color", "Size", "Language", "A", "khó", "Ô tick Effects."),
  Q("So sánh tốc độ 4 cách?", "Phím tắt nhanh nhất; nút x² nhanh; Font dialog chậm; copy ° nhanh cho ký hiệu", "Như nhau", "Font dialog nhanh nhất", "Copy chậm nhất", "A", "khó", "Đúng đánh giá video."),
  Q("Viết m² (mét vuông) nhanh nhất?", "Gõ m > Ctrl+Shift+= > gõ 2 > Ctrl+Shift+=", "Vẽ số 2", "Chèn ảnh", "Gõ m2", "A", "khó", "Áp dụng đúng phím tắt."),
  Q("Viết 25°C hoàn chỉnh?", "Gõ 25 > dán ° > gõ C", "Gõ 250C", "Gõ 25oC cỡ thường", "Gõ 25*C", "A", "khó", "° là ký tự riêng, không phải chữ o."),
  Q("Muốn mũ có 2 ký tự x-1 thì phải?", "Bật superscript rồi gõ liền x-1 rồi mới tắt", "Gõ từng ký tự rồi bật", "Gõ xong mới bôi đen", "Không gõ được", "A", "khó", "Gõ liền trong trạng thái mũ."),
  Q("Tóm tắt 4 cách trong video?", "Phím tắt Ctrl+Shift+=; nút x²; Font dialog Superscript; copy ° cho độ C", "Chỉ 1 cách", "Chỉ 2 cách", "Không có cách nào", "A", "khó", "Đủ 4 cách video dạy."),
 ]
}

# 7. ViJZcN-8cS8 - find replace
videos["ViJZcN-8cS8"] = {
 "bai": "Tìm và thay thế từ trong Word (Find & Replace)",
 "ky_nang": [
  {"ten": "Tìm kiếm với Ctrl+F", "thao_tac": "Nhấn Ctrl+F > gõ từ khóa vào ô Navigation > Enter, click từng kết quả", "duong_dan_menu": "Home > Find", "phim_tat": "Ctrl+F", "giai_thich": "Word bôi vàng mọi vị trí khớp, đếm số kết quả.", "vi_du": "Tìm 'đỗ bảo 5 lúc' được 8 kết quả"},
  {"ten": "Mở thay thế với Ctrl+H", "thao_tac": "Nhấn Ctrl+H mở Find and Replace", "duong_dan_menu": "Home > Replace", "phim_tat": "Ctrl+H", "giai_thich": "Hộp thoại có 2 ô Find what / Replace with.", "vi_du": "Ctrl+H"},
  {"ten": "Nhập từ gốc và từ mới", "thao_tac": "Find what: từ cũ; Replace with: từ viết hoa mới", "duong_dan_menu": "Home > Replace > Find what / Replace with", "phim_tat": "", "giai_thich": "Nhập chính xác cả hai.", "vi_du": "Find cũ, Replace IN HOA"},
  {"ten": "Thay từng chỗ hoặc tất cả", "thao_tac": "Replace: thay 1; Replace All: thay hết + OK", "duong_dan_menu": "Find and Replace > Replace / Replace All", "phim_tat": "", "giai_thich": "Video dùng Replace All được 8 chỗ.", "vi_du": "Replace All > 8 replacements > OK"}
 ],
 "cau_hoi": [
  Q("Phím tắt tìm kiếm là gì?", "Ctrl+F", "Ctrl+G", "Ctrl+H", "Ctrl+K", "A", "dễ", "Ctrl+F mở Navigation tìm."),
  Q("Từ khóa gõ vào đâu?", "Ô tìm kiếm khung Navigation bên trái", "Thanh địa chỉ", "Ô mật khẩu", "Thước kẻ", "A", "dễ", "Ô Find."),
  Q("Kết quả tìm được hiển thị sao?", "Bôi vàng và đếm số kết quả", "Biến mất", "In ra", "Kêu lên", "A", "dễ", "Video thấy 8 kết quả bôi vàng."),
  Q("Video tìm được bao nhiêu kết quả?", "8", "3", "1", "100", "A", "dễ", "Đúng số video nêu."),
  Q("Phím tắt thay thế là gì?", "Ctrl+H", "Ctrl+F", "Ctrl+S", "Ctrl+P", "A", "dễ", "Ctrl+H mở Replace."),
  Q("Ô nhập từ cần tìm trong Replace là?", "Find what", "Replace with", "Go to", "Font", "A", "dễ", "Find what là từ gốc."),
  Q("Ô nhập từ mới là?", "Replace with", "Find what", "Search", "Help", "A", "dễ", "Replace with là từ thay."),
  Q("Nút thay 1 chỗ là?", "Replace", "Replace All", "Find Next", "Close", "A", "dễ", "Replace thay từng cái."),
  Q("Nút thay tất cả là?", "Replace All", "Replace", "Undo", "Print", "A", "dễ", "Video dùng Replace All."),
  Q("Sau Replace All phải nhấn gì?", "OK xác nhận rồi đóng hộp thoại", "Tắt máy", "Xóa chữ", "In ngay", "A", "dễ", "OK > Close."),
  Q("Click từng kết quả tìm để làm gì?", "Nhảy đến đúng vị trí từ đó trong bài", "Xóa từ", "Copy từ", "In từ", "A", "trung bình", "Điều hướng nhanh."),
  Q("Find Next có tác dụng gì?", "Nhảy tới vị trí khớp tiếp theo", "Thay chữ", "Xóa chữ", "Lưu file", "A", "trung bình", "Duyệt từng chỗ."),
  Q("Vì sao nên dùng Replace All khi đổi tên đồng loạt?", "Nhanh, đủ 8/8 chỗ, không sót", "Chậm", "Sai", "Mất chữ", "A", "trung bình", "Video thay 8 chỗ 1 lần."),
  Q("Rủi ro của Replace All là?", "Thay nhầm cả từ chứa chuỗi con; nên kiểm tra", "Không rủi ro", "Mất máy", "Cháy màn hình", "A", "trung bình", "Cần Match whole word khi cần."),
  Q("Muốn phân biệt hoa/thường khi tìm thì tick gì?", "Match case", "Match all", "Find all", "Replace all", "A", "trung bình", "More >> Match case."),
  Q("Muốn tìm nguyên từ (không dính từ khác) tick gì?", "Find whole words only", "Match case", "Wildcards", "Sounds like", "A", "trung bình", "Tránh thay nhầm chuỗi con."),
  Q("Nút More >> trong Replace để làm gì?", "Mở tùy chọn nâng cao (Match case, Whole words...)", "Thêm chữ", "Xóa chữ", "Tắt máy", "A", "trung bình", "Mở rộng tìm kiếm."),
  Q("Hoàn tác sau Replace All nhầm bằng gì?", "Ctrl+Z", "Ctrl+Y", "Ctrl+S", "Ctrl+X", "A", "trung bình", "Undo khôi phục."),
  Q("Navigation pane ngoài tìm chữ còn tìm gì?", "Headings, Pages, Results", "Mạng", "Mực in", "Virus", "A", "trung bình", "3 tab điều hướng."),
  Q("Find trong Home thuộc nhóm nào?", "Editing", "Font", "Paragraph", "Styles", "A", "trung bình", "Nhóm Editing góc phải."),
  Q("Muốn thay định dạng (màu/cỡ) thì dùng gì?", "More >> Format trong Replace", "Gõ tay", "Chụp ảnh", "In ra tô màu", "A", "trung bình", "Replace cả định dạng."),
  Q("Tìm ký tự đặc biệt (xuống dòng, tab) thì?", "More >> Special chọn Paragraph Mark, Tab...", "Không tìm được", "Gõ tay", "Bỏ qua", "A", "trung bình", "Special liệt kê ký tự ẩn."),
  Q("Video thay từ thường thành dạng nào?", "VIẾT HOA hoàn toàn", "Viết thường", "Xóa trắng", "In đậm", "A", "khó", "Đúng nội dung video."),
  Q("Nếu Find what bỏ trống mà Replace All thì?", "Word báo lỗi/phải nhập từ tìm", "Thay cả bài", "Xóa bài", "Tắt Word", "A", "khó", "Phải có chuỗi tìm."),
  Q("Muốn tìm 'an' mà không dính 'ban' 'lan'?", "Tick Find whole words only", "Tick Match case", "Gõ AN", "Không thể", "A", "khó", "Whole words tách từ độc lập."),
  Q("Wildcards dùng để?", "Tìm theo mẫu (ví dụ d??? cho 4 ký tự)", "Vẽ hình", "Chèn nhạc", "Tắt máy", "A", "khó", "Use wildcards tìm mẫu."),
  Q("Muốn thay 8 chỗ nhưng xem từng chỗ trước thì?", "Nhấn Find Next rồi Replace từng cái thay vì Replace All", "Nhắm mắt thay", "Đoán", "Xóa hết", "A", "khó", "An toàn khi văn bản quan trọng."),
  Q("Thông báo sau Replace All cho biết gì?", "Số vị trí đã thay (video: 8 replacements)", "Số trang", "Số virus", "Số ảnh", "A", "khó", "Hộp thoại báo số lượng."),
  Q("Muốn tìm trong cả file dài nhiều trang nhanh nhất?", "Ctrl+F gõ 1 lần thấy hết và đếm, không cần lăn chuột dò", "Đọc từng trang", "In ra đọc", "Nhờ người khác", "A", "khó", "Đúng lợi ích Find."),
  Q("Quy trình thay thế hoàn chỉnh?", "Ctrl+H > Find what > Replace with > Replace All > OK > Close > kiểm tra", "Mở file rồi tắt", "Gõ tay từng chỗ", "Copy sang Paint", "A", "khó", "Đúng các bước video."),
 ]
}

# 8. c5u8mrgYKzk - landscape
videos["c5u8mrgYKzk"] = {
 "bai": "Xoay ngang trang giấy trong Word (Landscape)",
 "ky_nang": [
  {"ten": "Xoay ngang toàn bộ file", "thao_tac": "Layout > Orientation > Landscape", "duong_dan_menu": "Layout > Orientation > Landscape", "phim_tat": "", "giai_thich": "Mọi trang đều thành ngang.", "vi_du": "Orientation > Landscape"},
  {"ten": "Xoay ngang 1 trang bất kỳ", "thao_tac": "Bôi đen nội dung trang đó > Page Setup launcher > Landscape > Apply to: Selected text > OK", "duong_dan_menu": "Layout > Page Setup dialog > Landscape > Selected text", "phim_tat": "", "giai_thich": "Chỉ đoạn được chọn thành ngang, trang khác giữ nguyên.", "vi_du": "Bôi đen trang 2 > Landscape > Selected text"},
  {"ten": "Mở hộp thoại Page Setup", "thao_tac": "Layout > click mũi tên nhỏ góc nhóm Page Setup", "duong_dan_menu": "Layout > Page Setup launcher", "phim_tat": "", "giai_thich": "Hộp thoại chứa Orientation và Apply to.", "vi_du": "Click launcher góc Page Setup"},
  {"ten": "Xóa trang trắng dư", "thao_tac": "Đặt con trỏ trang trắng > Backspace/Delete", "duong_dan_menu": "", "phim_tat": "Backspace/Delete", "giai_thich": "Xoay Selected text có thể sinh trang trắng, xóa đi.", "vi_du": "Xóa trang trắng phát sinh"}
 ],
 "cau_hoi": [
  Q("Xoay ngang tiếng Anh trong Word là gì?", "Landscape", "Portrait", "Layout", "Margin", "A", "dễ", "Landscape = ngang."),
  Q("Trang dọc mặc định gọi là gì?", "Portrait", "Landscape", "Square", "Circle", "A", "dễ", "Portrait = dọc."),
  Q("Nút Orientation nằm ở thẻ nào?", "Layout", "Home", "Insert", "View", "A", "dễ", "Layout > Orientation."),
  Q("Muốn xoay ngang CẢ file thì chọn gì?", "Landscape", "Portrait", "Custom", "Print", "A", "dễ", "Orientation > Landscape."),
  Q("Video demo file mẫu có mấy trang?", "3 trang", "1 trang", "10 trang", "100 trang", "A", "dễ", "Video tạo thành 3 trang."),
  Q("Muốn xoay 1 trang thì bôi đen nội dung trang đó trước đúng không?", "Đúng", "Sai", "Tùy", "Không cần", "A", "dễ", "Phải chọn Selected text."),
  Q("Mở hộp thoại Page Setup bằng cách nào?", "Click mũi tên nhỏ góc nhóm Page Setup", "Nhấn Enter", "Tắt máy", "Xóa trang", "A", "dễ", "Launcher góc dưới-phải nhóm."),
  Q("Trong Page Setup chọn hướng nào để ngang?", "Landscape", "Portrait", "Paper", "Margins", "A", "dễ", "Chọn Landscape."),
  Q("Apply to chọn gì để chỉ xoay đoạn đã chọn?", "Selected text", "Whole document", "This point forward", "Printer", "A", "dễ", "Selected text = chỉ chỗ chọn."),
  Q("Sau khi cài xong nhấn gì?", "OK", "Cancel", "Delete", "Print", "A", "dễ", "OK áp dụng."),
  Q("Whole document khác Selected text thế nào?", "Whole: cả file; Selected: chỉ chỗ bôi đen", "Giống nhau", "Whole xóa file", "Selected tắt máy", "A", "trung bình", "Đúng phạm vi Apply to."),
  Q("Vì sao xoay 1 trang lại sinh trang trắng?", "Word chèn section break bao quanh đoạn chọn, có thể dư break", "Lỗi chuột", "Hết giấy", "Hết mực", "A", "trung bình", "Cơ chế section."),
  Q("Xử lý trang trắng dư thế nào?", "Đặt con trỏ và Delete/Backspace", "Xóa cả file", "Cài lại Word", "Bỏ qua", "A", "trung bình", "Xóa break thừa."),
  Q("Hoàn tác xoay ngang nhầm bằng gì?", "Ctrl+Z", "Ctrl+Y", "Ctrl+S", "Ctrl+X", "A", "trung bình", "Undo."),
  Q("Muốn trang 2 ngang, trang 1 và 3 vẫn dọc thì?", "Bôi đen trang 2 > Landscape > Selected text", "Xoay cả file", "Xóa trang 1,3", "In trang 2 riêng", "A", "trung bình", "Đúng bài video."),
  Q("Dấu hiệu trang đã ngang?", "Rộng hơn cao, chữ dàn theo chiều ngang", "Cao hơn rộng", "Mất chữ", "Đổi màu", "A", "trung bình", "Quan sát trực quan."),
  Q("Orientation còn tùy chọn khổ giấy không?", "Không, khổ giấy ở Size; Orientation chỉ dọc/ngang", "Có", "Tùy máy", "Tùy mực", "A", "trung bình", "Size (A4...) khác Orientation."),
  Q("Page Setup launcher là gì?", "Mũi tên nhỏ góc nhóm Page Setup để mở hộp thoại", "Nút in", "Nút lưu", "Nút tắt", "A", "trung bình", "Đúng định nghĩa."),
  Q("Khi nào dùng This point forward?", "Xoay từ vị trí con trỏ đến hết file", "Xoay 1 đoạn", "Xoay cả file", "Xóa file", "A", "trung bình", "Áp dụng từ đây về sau."),
  Q("Xoay ngang phù hợp cho nội dung nào?", "Bảng rộng, sơ đồ, slide in ngang", "Thư ngắn", "Tiểu thuyết", "Mọi thứ", "A", "trung bình", "Ngang chứa bảng rộng tốt."),
  Q("Sau khi xoay 1 trang, lề các trang có đổi không?", "Mỗi section giữ lề riêng; trang ngang/dọc độc lập", "Mất hết lề", "Đổi hết", "Không có lề", "A", "trung bình", "Section tách thiết lập trang."),
  Q("Muốn xem trước hướng trang thì vào đâu?", "File > Print xem preview", "Nhìn bàn phím", "Nhìn chuột", "Đoán", "A", "trung bình", "Print preview thấy dọc/ngang."),
  Q("Vì sao phải bôi đen HẾT nội dung trang 2?", "Để Word xác định đúng phạm vi section cần xoay", "Để xóa", "Để copy", "Để trang trí", "A", "khó", "Bôi thiếu sẽ xoay lệch."),
  Q("Section break là gì?", "Dấu ngắt chia file thành các section cài đặt trang riêng", "Ngắt chữ", "Ngắt mạng", "Ngắt điện", "A", "khó", "Đúng khái niệm."),
  Q("Xem section break bằng cách nào?", "Home > Show/Hide ¶", "Tắt màn hình", "In ra", "Đoán", "A", "khó", "Ký tự ẩn hiện break."),
  Q("Xóa nhầm section break thì?", "Trang gộp lại cùng hướng, Ctrl+Z để hoàn tác", "Mất file", "Mất máy", "Không sao", "A", "khó", "Break quyết định hướng."),
  Q("Khổ A4 ngang kích thước bao nhiêu?", "29.7cm rộng x 21cm cao (đảo của dọc)", "21x29.7 dọc", "10x10", "50x50", "A", "khó", "Ngang đảo rộng-cao."),
  Q("In trang ngang cần chú ý gì?", "Chọn đúng khổ/orientation trong Print khớp Page Setup", "Không cần", "In là ra", "Xoay giấy tay", "A", "khó", "Lệch sẽ bị cắt chữ."),
  Q("Muốn file vừa có dọc vừa có ngang thì bắt buộc?", "Phải có nhiều section (Selected text/This point forward)", "Không cần", "Phải 2 file", "Phải 2 máy", "A", "khó", "Mỗi hướng cần 1 section."),
  Q("Tóm tắt video?", "Orientation > Landscape cho cả file; bôi đen + Page Setup + Selected text cho 1 trang + xóa trắng dư", "Chỉ xoay cả file", "Không xoay được", "Xoay bằng tay", "A", "khó", "Đủ 2 phần video."),
 ]
}

# 9. eLbuCLA5kVk - indent
videos["eLbuCLA5kVk"] = {
 "bai": "3 cách thụt đầu dòng trong Word (Indent)",
 "ky_nang": [
  {"ten": "Thụt đầu dòng bằng Tab", "thao_tac": "Đặt con trỏ đầu dòng > nhấn Tab", "duong_dan_menu": "", "phim_tat": "Tab", "giai_thich": "Nhanh nhất cho 1 dòng.", "vi_du": "Đầu dòng + Tab"},
  {"ten": "Chỉnh khoảng Tab bằng Paragraph", "thao_tac": "Chuột phải > Paragraph > Special: First line > By: 1cm > OK", "duong_dan_menu": "Chuột phải > Paragraph > First line", "phim_tat": "", "giai_thich": "Chuẩn hóa khoảng thụt (mặc định 1.27cm, video đổi 1cm).", "vi_du": "First line By 1cm"},
  {"ten": "Kéo thước Ruler", "thao_tac": "Đặt con trỏ vào đoạn > kéo marker First Line Indent trên Ruler", "duong_dan_menu": "View > Ruler", "phim_tat": "", "giai_thich": "Kéo trực quan trên thước.", "vi_du": "Kéo marker trên 1cm"},
  {"ten": "Thụt hàng loạt nhiều đoạn", "thao_tac": "Bôi đen nhiều đoạn > Paragraph > First line 1cm > OK", "duong_dan_menu": "Chuột phải > Paragraph > First line", "phim_tat": "", "giai_thich": "Áp cùng lúc cho nhiều đoạn.", "vi_du": "Bôi đen 3 đoạn > First line 1cm"}
 ],
 "cau_hoi": [
  Q("Cách nhanh nhất thụt 1 dòng là gì?", "Nhấn Tab đầu dòng", "Gõ 5 dấu cách", "Enter", "Xóa dòng", "A", "dễ", "Video: Tab là nhanh nhất."),
  Q("Trước khi nhấn Tab phải làm gì?", "Đặt con trỏ vào đầu dòng", "Tắt máy", "In ra", "Xóa chữ", "A", "dễ", "Đúng bước 1."),
  Q("Mở hộp thoại Paragraph bằng cách nào?", "Chuột phải > Paragraph", "Nhấn Enter", "Tắt Word", "Rút điện", "A", "dễ", "Đúng video."),
  Q("Trong Paragraph chọn mục nào để thụt dòng đầu?", "Special: First line", "Hanging", "None", "Justify", "A", "dễ", "First line = dòng đầu."),
  Q("Video đặt khoảng thụt bao nhiêu?", "1cm", "5cm", "10cm", "0cm", "A", "dễ", "Từ 1.27cm giảm còn 1cm."),
  Q("Mặc định First line của Word là bao nhiêu?", "1.27cm (0.5 inch)", "1cm", "2cm", "5cm", "A", "dễ", "0.5 inch = 1.27cm."),
  Q("Cách 2 thao tác trên đâu?", "Thanh Ruler (thước)", "Thanh taskbar", "Bàn phím", "Máy in", "A", "dễ", "Kéo marker trên Ruler."),
  Q("Muốn thụt 3 đoạn cùng lúc thì?", "Bôi đen 3 đoạn rồi Paragraph > First line", "Tab từng dòng", "Gõ cách từng dòng", "Không thể", "A", "dễ", "Bôi đen rồi làm 1 lần."),
  Q("Hiện thước Ruler vào đâu?", "View > tick Ruler", "File > Save", "Home > Bold", "Insert > Table", "A", "dễ", "View > Ruler."),
  Q("Sau khi chỉnh Paragraph nhấn gì?", "OK", "Cancel", "Delete", "Print", "A", "dễ", "OK lưu."),
  Q("Vì sao không nên gõ nhiều dấu cách để thụt?", "Khoảng cách lệch, khó đều, sửa khổ giấy là vỡ", "Nhanh hơn", "Đẹp hơn", "Chuẩn hơn", "A", "trung bình", "Tab/First line chuẩn hơn Space."),
  Q("First line khác Hanging ở điểm nào?", "First: thụt dòng đầu; Hanging: thụt từ dòng 2 trở đi", "Giống nhau", "Hanging xóa dòng đầu", "First xóa dòng cuối", "A", "trung bình", "Đúng định nghĩa."),
  Q("Marker nào trên Ruler là First Line?", "Tam giác trên cùng", "Tam giác dưới", "Hình vuông", "Thước dọc", "A", "trung bình", "Tam giác trên = First line."),
  Q("Kéo nhầm cả hình vuông dưới thì sao?", "Thụt cả đoạn (left indent) chứ không chỉ dòng đầu", "Mất chữ", "Xóa đoạn", "Tắt máy", "A", "trung bình", "Vuông = Left indent cả đoạn."),
  Q("By 1cm nghĩa là?", "Dòng đầu thụt vào 1cm so với lề trái", "Cả bài thụt 1cm", "Lề giấy 1cm", "Cỡ chữ 1cm", "A", "trung bình", "Đúng ý nghĩa."),
  Q("Muốn bỏ thụt đầu dòng thì?", "Special: None > OK", "Xóa đoạn", "Tắt máy", "Đổi font", "A", "trung bình", "None gỡ indent."),
  Q("Tab và First line khác nhau thế nào?", "Tab chèn 1 ký tự; First line là định dạng đoạn, chuẩn và sửa hàng loạt được", "Giống nhau", "Tab chuẩn hơn", "First line sai", "A", "trung bình", "First line chuyên nghiệp hơn."),
  Q("Dấu hiệu đoạn đã có First line?", "Dòng đầu tự thụt khi Enter đoạn mới cùng style", "Mất chữ", "Đổi màu", "Kêu lên", "A", "trung bình", "Định dạng đoạn lan theo."),
  Q("Muốn mọi đoạn mới đều thụt 1cm thì?", "Đặt First line rồi cập nhật Style Normal hoặc Set Default", "Gõ Tab mãi", "Không thể", "Cài lại Word", "A", "trung bình", "Lưu vào style/mặc định."),
  Q("Bôi đen rồi Tab thì sao?", "Thụt cả khối sang phải (tăng left indent)", "Chỉ dòng đầu", "Xóa khối", "Không gì", "A", "trung bình", "Tab cả khối = indent khối."),
  Q("Shift+Tab khi bôi đen khối thì?", "Giảm thụt (outdent)", "Tăng thụt", "Xóa", "Lưu", "A", "trung bình", "Ngược với Tab."),
  Q("1.27cm từ đâu ra?", "0.5 inch x 2.54", "Ngẫu nhiên", "1cm + 0.27", "Đoán", "A", "trung bình", "0.5*2.54=1.27."),
  Q("Vì sao video đổi 1.27 về 1cm?", "Chẵn, gọn, hợp trình bày Việt Nam", "Cho khó", "Cho vui", "Máy yêu cầu", "A", "khó", "1cm dễ nhớ, đều."),
  Q("Đoạn đã Tab rồi lại áp First line 1cm thì?", "Có thể thụt gấp đôi (Tab + format); nên xóa Tab thừa", "Không sao", "Mất chữ", "Đẹp hơn", "A", "khó", "Tránh cộng dồn."),
  Q("Hiện ký tự Tab (mũi tên) bằng cách nào?", "Show/Hide ¶", "In ra", "Zoom", "Tắt máy", "A", "khó", "¶ hiện Tab/Enter."),
  Q("Trong luận văn, thụt đầu dòng chuẩn kèm?", "First line + giãn dòng + căn đều Justify", "Chỉ thụt", "Chỉ màu", "Chỉ in đậm", "A", "khó", "Combo trình bày chuẩn."),
  Q("Kéo Ruler mà không thấy Ruler thì?", "View > Ruler để hiện; kiểm tra đang ở Print Layout", "Bỏ qua", "Đoán", "Tắt Word", "A", "khó", "Ruler chỉ hiện ở chế độ phù hợp."),
  Q("Muốn thụt treo (dòng 2 trở đi) cho tài liệu tham khảo?", "Special: Hanging + By khoảng cần", "First line", "None", "Tab", "A", "khó", "Hanging cho bibliography."),
  Q("Muốn kiểm tra 3 đoạn đã đều 1cm?", "Bôi đen > Paragraph thấy First line By 1cm + quan sát Ruler", "Đoán mắt", "In ra đo", "Hỏi người khác", "A", "khó", "Kiểm tra dialog + ruler."),
  Q("Tóm tắt 3 cách video?", "Tab nhanh; Ruler trực quan; Paragraph First line chuẩn + hàng loạt", "Chỉ Tab", "Chỉ Ruler", "Không cách nào", "A", "khó", "Đủ 3 cách."),
 ]
}

for vid, d in videos.items():
    assert len(d["cau_hoi"]) == 30, (vid, len(d["cau_hoi"]))
    de = sum(1 for q in d["cau_hoi"] if q["muc_do"]=="dễ"); tb = sum(1 for q in d["cau_hoi"] if q["muc_do"]=="trung bình"); kho = sum(1 for q in d["cau_hoi"] if q["muc_do"]=="khó")
    assert (de,tb,kho)==(10,12,8), (vid,de,tb,kho)
    obj = {"video_id": vid, "bai": d["bai"], "url": f"https://www.youtube.com/watch?v={vid}", "ky_nang": d["ky_nang"], "cau_hoi": d["cau_hoi"]}
    p = os.path.join(OUT, vid+".json")
    with open(p, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    print("OK", vid, len(d["ky_nang"]), "kn,", len(d["cau_hoi"]), "q")
