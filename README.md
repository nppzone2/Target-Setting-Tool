# Target Setting Tool · Bottom-up từ Outlet

Chia Target tháng M+1 bottom-up: Sales Out L3M của từng Outlet → Outlet → NPP → Area. Tool chạy hoàn toàn trên trình duyệt, không cần Python, dữ liệu không gửi lên máy chủ.

Developed by Ta Thi Thu Minh

## File đầu vào

| File | Bắt buộc | Cột chính |
|---|---|---|
| SI-SO ([DIP – Power BI](https://app.powerbi.com/groups/me/apps/0ddfc8d6-a19c-4f0b-90f2-4357e8be8678/reports/7be06dc7-1a2c-4bc0-9a74-32321b2d4ebd/09b8d87a94136c04100e?ctid=66e853de-ece3-44dd-9d66-ee6bdf4159d4&experience=power-bi)) | Có | Region, Area, Distributor_Name, ShortCode, calendar_month, Target (Case), Sales Out (Case) |
| SO Detail by Outlet | Có (nhiều file) | Area_Name, Seller_Name, Outlet_ID, Outlet_Name, Outlet_Segment, Calendar_Month, ShortCode, UOM (Case) |
| Target ([Target DIP](https://hstsvn140236.hosting.heiway.net/Main/TargetByArea/Target_ByArea.aspx)) | Có | Target by Month (Area, ShortCode, cột tháng, Htls) hoặc Target M+1 theo Case (Item_Code × Area) |
| Item Detail | Không | BrandGroup, BrandFamily, ShortCode 2026, Hl (đã có danh mục mặc định) |
| CP theo NPP | Không | Distributor_Name, ShortCode, cột số CP (Case) |

## Logic chia

```
g (SKU)          = Target Region M+1 ÷ Target Region M − 1
Đề xuất Outlet   = SO bình quân L3M × (1 + g) + Phần chênh × Contribution
Contribution     = SO L3M của Outlet ÷ SO L3M của Region (theo SKU)
```

1. Tháng M là tháng gần nhất trong SI-SO. Target tháng M lấy từ file Target by Month (Htls → Case theo Hl/Case); không có thì lấy Target (Case) trong SI-SO.
2. CP được tách khỏi Target Region trước khi chia, sau đó cộng thẳng vào đúng NPP.
3. Gốc Outlet = Sales Out bình quân 3 tháng đủ trước M, theo Outlet × SKU; số âm tính bằng 0. NPP đã chuyển (tham số `P468=P449`) được gộp sang NPP nhận.
4. Outlet có mua trước đó nhưng Sales Out L3M = 0 nhận tối thiểu 1 case cho mỗi SKU chính (HS2, HSC, TS25, TSS, TC25, TCS), trừ ra trước.
5. Áp g cho gốc từng Outlet; phần chênh còn lại so với Target Region chia theo contribution.
6. Cộng dồn Outlet → NPP → Area, làm tròn theo phần dư lớn nhất để tổng từng cấp luôn khớp Target Region theo SKU.
7. SKU chưa Outlet nào có gốc: giữ số Area theo file Target, nhập tay xuống NPP ở thẻ NPP (Area một NPP thì tự giao).
8. Area adjust theo Brand family: nhập % tăng trưởng mới so với Target M; chênh so với đề xuất trong giới hạn −5% đến +10%. Các Area còn lại tự cân (RAS), tổng Region theo Brand family và theo từng SKU giữ nguyên. Area tăng trên 10% so với Target M được cảnh báo.

## Kết quả

Thẻ **Tổng quan**, **Area adjust**, **NPP**, **Outlet** và file Excel gồm: Tóm tắt, Region SKU, Area Brand, NPP SKU (Target chốt gửi NPP, cột Final), Outlet SKU, Chờ adjust, Kiểm tra.

## Kho dữ liệu GitHub

Lưu file, tham số và Adjust lên repo dưới dạng mã hóa AES-256 (thư mục `data/bottomup`). Cần token fine-grained quyền **Contents: Read and write** cho repo này và mật khẩu mã hóa tối thiểu 8 ký tự. Thư mục `data/` không kích hoạt deploy lại.

## Phiên bản trước

Tool cũ (Region → Area → NPP theo Target M và run-rate, cân YoY) giữ nguyên tại `legacy/` (link: `/legacy/` trên GitHub Pages).
