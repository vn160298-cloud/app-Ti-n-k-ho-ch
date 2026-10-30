# Tiền kế hoạch

App web cài được lên điện thoại (PWA): tổng hợp **forecast WMS + tồn kho thành phẩm + tồn NVL - phụ gia + cân chỉnh kênh** thành **bảng tiền kế hoạch**, rồi sinh **kế hoạch sản xuất**.

## Cấu trúc file

| File | Vai trò |
|---|---|
| `index.html` | Khung app, khai báo chạy độc lập (standalone) cho iPhone/Android |
| `app.css` | Giao diện (tự đổi sáng/tối theo máy) |
| `app.js` | Toàn bộ logic: nhập Excel, tổng hợp, kế hoạch, đồng bộ |
| `manifest.webmanifest` | Tên app, icon, chế độ toàn màn hình |
| `sw.js` | Service worker: mở app khi mất mạng. **Tăng `VERSION` mỗi lần sửa app** |
| `icons/` | Icon app |

## Công thức

- **Nhu cầu** = Forecast (tổng các kênh) + Cân chỉnh kênh (±)
- **Tồn an toàn** = Nhu cầu × số ngày tồn an toàn / 30
- **Cần SX** = max(0, Nhu cầu + Tồn an toàn − Tồn TP)
- **Đề xuất SX** = Cần SX làm tròn lên theo lô sản xuất (số lượng/mẻ)
- **Nhu cầu NVL** = Đề xuất SX × định mức BOM, so với Tồn + Đang về
- **Số ngày chạy** = Số lượng / Công suất line (bỏ Chủ nhật)

## Chạy thử trên máy

```bash
python -m http.server 5178
```

Mở http://localhost:5178

## Đưa lên mạng để cài vào điện thoại

"Thêm vào màn hình chính" cần trang chạy **HTTPS**. Cách nhanh nhất, miễn phí:

1. **Netlify Drop**: vào app.netlify.com/drop, kéo cả thư mục này vào → có ngay link `https://….netlify.app`.
2. Hoặc **Cloudflare Pages** / **GitHub Pages** / **Vercel** (cũng miễn phí).

Trên điện thoại mở link đó:
- **iPhone (Safari)**: nút Chia sẻ → *Thêm vào MH chính*.
- **Android (Chrome)**: menu ⋮ → *Cài đặt ứng dụng*.

## Lưu trữ dữ liệu đám mây (đề xuất)

| Phương án | Khi nào chọn |
|---|---|
| **Supabase** (khuyến nghị) | PostgreSQL, dữ liệu dạng bảng, có đăng nhập & phân quyền, miễn phí 500MB |
| Firebase Firestore | Cần đồng bộ thời gian thực nhiều máy |
| SharePoint / OneDrive | Công ty đã dùng Microsoft 365 |
| Google Sheets + Apps Script | Quy mô nhỏ, muốn xem trên Sheet |

App đã tích hợp sẵn Supabase (màn **Cài đặt**): tạo project, chạy đoạn SQL hiển thị trong app, dán *Project URL* + *anon key*, chọn mã nhà máy rồi bấm *Đẩy lên* / *Tải về*.

> Lưu ý bảo mật: chính sách mẫu cho phép ai có anon key đều đọc/ghi được. Khi dùng thật nên bật Supabase Auth (đăng nhập email công ty) và đổi policy sang `to authenticated`.

## Việc cần bổ sung

- Danh mục **mã sản phẩm** và **mã NVL - phụ gia** chuẩn (Danh mục → nhập Excel theo file mẫu).
- **Định mức BOM** cho từng sản phẩm.
- **Line sản xuất** và công suất thực tế.
- Gửi 1 file forecast WMS thật để kiểm tra tên cột (app tự nhận diện, có thể chọn lại bằng danh sách).
