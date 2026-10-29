# Lập hồ sơ kết nạp đảng viên

Nhập thông tin quần chúng một lần, tool tự điền và xuất 5 mẫu Word:
Mẫu 1-KNĐ (Đơn xin vào Đảng), Mẫu 3 (Giấy giới thiệu), Mẫu 5 (Tổng hợp ý kiến đoàn thể và nơi cư trú),
Mẫu 6 (Nghị quyết chi bộ), Mẫu 8 (Nghị quyết Đảng ủy).

## Cài đặt

Cần Python 3 (Windows):

```
pip install -r requirements.txt
```

## Sử dụng

1. Bấm đúp `Lap ho so ket nap.bat`.
2. Tab **Đơn vị** chỉ nhập lần đầu (được lưu vào `tool/don_vi.json`).
3. Nhập các tab còn lại, bấm **XUẤT 5 MẪU WORD**. Kết quả nằm trong `HoSo/<Họ tên>/`.

Số liệu, tỷ lệ % tự tính; ô bỏ trống giữ dấu "……" để viết tay.

### Nhập bằng file Excel

Điền thông tin vào `Mau nhap thong tin.xlsx` (cột "Nội dung nhập"; các ô nhận xét có danh sách chọn mẫu),
rồi bấm **Nhập từ Excel** trên tool hoặc **Nhập file Excel** trên bản online. File mẫu cũng tải được ngay trên trang online.
Tạo lại file mẫu bằng `python tool/build_excel.py` (sau đó chạy `python tool/build_web.py`).

## Bản online

Dùng online: **https://thanhluongst.github.io/ho-so-ket-nap-dang-vien/** (GitHub Pages, từ thư mục `docs/`).

`web/index.html` là mã nguồn trang. Sau khi sửa trang, mẫu trong `tool/mau/` hoặc `tool/mau_nhan_xet.json`,
chạy `python tool/build_web.py` để tạo lại `web/` và `docs/`, rồi commit và push.
Thông tin nhập trên trang chỉ nằm trong trình duyệt, không gửi đi đâu.

## Cấu trúc

- `*.doc` – mẫu gốc.
- `tool/goc/` – mẫu gốc chuyển sang .docx.
- `tool/mau/` – mẫu có trường điền `{{ ... }}`, sinh bởi `tool/build_templates.py`.
- `tool/ket_nap.py` – giao diện nhập liệu và xuất hồ sơ.
- `tool/build_web.py`, `web/` – bản online.

Thư mục `HoSo/` và `tool/don_vi.json` chứa thông tin cá nhân nên không được đưa lên GitHub.
