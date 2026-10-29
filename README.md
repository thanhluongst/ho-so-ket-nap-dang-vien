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

## Bản online

`web/index.html` là bản chạy trên trình duyệt (đã đăng tại https://claude.ai/artifact/75KXCyc32X96DLvaM24sFt, chế độ riêng tư).
Sau khi sửa mẫu trong `tool/mau/`, chạy `python tool/build_web.py` để tạo lại `web/mau_data.js`.

## Cấu trúc

- `*.doc` – mẫu gốc.
- `tool/goc/` – mẫu gốc chuyển sang .docx.
- `tool/mau/` – mẫu có trường điền `{{ ... }}`, sinh bởi `tool/build_templates.py`.
- `tool/ket_nap.py` – giao diện nhập liệu và xuất hồ sơ.
- `tool/build_web.py`, `web/` – bản online.

Thư mục `HoSo/` và `tool/don_vi.json` chứa thông tin cá nhân nên không được đưa lên GitHub.
