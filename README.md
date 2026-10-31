# Chia Target M+1

Tool chia Target tháng M+1 theo 3 cấp Region → Area → Distributor (NPP); các Area tăng trưởng đều, mỗi cấp có cột Adjust, dựa trên file SI-SO và file Target by Month mới nhất. Tool chạy hoàn toàn trên trình duyệt: không cần cài Python, dữ liệu không gửi lên máy chủ nào.

Developed by Ta Thi Thu Minh

## Đưa tool lên GitHub Pages (làm một lần)

1. Đăng nhập GitHub, bấm **New repository**, đặt tên (ví dụ `chia-target`), chọn **Public**, bấm **Create repository**.
2. Trong repo vừa tạo, bấm **Add file → Upload files**, kéo thả `index.html` và `README.md`, bấm **Commit changes**.
3. Vào **Settings → Pages**. Ở mục *Build and deployment*, chọn **Source: Deploy from a branch**, **Branch: main**, thư mục **/(root)**, bấm **Save**.
4. Sau 1–2 phút, link tool hiển thị ở đầu trang Pages, dạng `https://<tên-tài-khoản>.github.io/chia-target/`.

Repo Public chỉ chứa code của tool, không chứa dữ liệu. File SI-SO và Target chỉ được đọc trên máy người dùng khi mở tool.

Khi cần cập nhật tool, upload đè file `index.html` mới vào repo, GitHub Pages tự cập nhật.

## Cách lấy file

Nguồn dữ liệu (cũng có link ngay trong từng ô chọn file trên tool):
- SI-SO năm nay: [DIP – Power BI](https://app.powerbi.com/groups/me/apps/0ddfc8d6-a19c-4f0b-90f2-4357e8be8678/reports/7be06dc7-1a2c-4bc0-9a74-32321b2d4ebd/09b8d87a94136c04100e?ctid=66e853de-ece3-44dd-9d66-ee6bdf4159d4&experience=power-bi)
- Target by Month: [Target DIP](https://hstsvn140236.hosting.heiway.net/Main/TargetByArea/Target_ByArea.aspx)
- SI-SO năm trước: [Distributors Sales_Slicer1 – Power BI](https://app.powerbi.com/groups/me/apps/0ddfc8d6-a19c-4f0b-90f2-4357e8be8678/reports/02faacf9-f840-4f23-9bdb-b293e3d39e7d/ReportSection203158e276f7c10a0501?ctid=66e853de-ece3-44dd-9d66-ee6bdf4159d4&experience=power-bi)

**File SI-SO**
- Xuất dạng bảng từ báo cáo SI-SO ở cấp Region, Area, Distributor_Name, ShortCode, calendar_month.
- Giữ các bộ lọc như file mẫu: Customer_Name is not blank, Sales_Target is not blank, Geography = Greater HCM.
- Cột bắt buộc: `Region`, `Area`, `Distributor_Name`, `ShortCode`, `calendar_month` (YYYYMM), `Target (Case)`, `Target (Htls)`, `Sales In (Case)`, `Sales Out (Case)`.
- Lấy ít nhất 4 tháng gần nhất, gồm tháng hiện tại (M). Dòng "Applied filters" ở đầu file không ảnh hưởng.

**File Target by Month**
- Mỗi dòng là một Area × ShortCode, có cột `Area`, `ShortCode`, `SKU`.
- Các cột tháng là ngày đầu tháng (01/01/2026 … 01/12/2026), đơn vị **Htls**.
- Có thể chọn nhiều file cùng lúc (ví dụ mỗi Zone một file).

**File Target tháng M+1 theo Case**
- Ô File Target nhận 2 dạng: Target by Month (như trên) hoặc Target tháng M+1 theo Case (cột `Item_Code` và mỗi Area một cột, đơn vị Case). Chọn một dạng hoặc cả hai.
- Chỉ có file Case: tool dùng danh mục SKU nhúng sẵn để lấy tên SKU, Brand family, Group AA/BB; phần bù YoY dồn vào tháng M+1 vì không có số các tháng sau.
- Có cả hai: số tháng M+1 theo file Case, các thông tin còn lại và các tháng sau theo Target by Month (khuyến nghị).

**File SI-SO năm trước (không bắt buộc)**
- Xuất từ báo cáo SI-SO với Calendar_Year là năm trước, cùng Region. Cột: `Zone_Code`, `Customer_Name`, `TypeName`, `Calendar_Month`, `Search_Text`, `ValueFormat` (dòng `TypeName = SaleOut` được dùng). Tool cũng nhận file cùng định dạng SI-SO năm nay.
- Cần đủ tháng 1 tới tháng 12 năm trước, cùng phạm vi kênh/khách hàng với file năm nay.
- Có file này, tool cân tổng Area theo tăng trưởng YoY lũy kế (xem Logic chia, bước 2b).
- **Ghép mã** năm trước sang năm nay và **bỏ qua mã** chỉnh trên giao diện. Mặc định: `TCC=TSS, TCC25=TS25, HZM=H0M, HSC25=HS2, TC=TCS, HCS=HS, HS=HSB`; bỏ qua `BV12, SB1w, SBC, TJC, TJP, TPC20`. Mã năm trước không có ở năm nay được liệt kê ở phần Kiểm tra.

**File CP theo NPP (không bắt buộc)**
- Số Case CT mua hàng CP đã phân bổ cho từng NPP, cột `Distributor_Name`, `ShortCode`, cột số CP (Case).
- Trên tool, chọn file SI-SO và Target rồi bấm **Tải file mẫu** để có sẵn danh sách NPP × SKU, chỉ cần điền số CP.
- Tool trừ CP khỏi Target Region trước khi cân/chia (trần/sàn chỉ áp cho phần ngoài CP), sau đó cộng CP thẳng vào đúng NPP.

## Các bước mỗi tháng

1. Mở link tool, chọn file SI-SO và file Target mới nhất.
2. Kiểm tra tham số, bấm **Chia số**.
3. Đọc phần Kiểm tra và danh sách SKU chờ adjust.
4. Bấm **Tải file Excel**. Adjust lần lượt ở sheet *1. Region*, *2. Area*, *3. NPP* nếu cần.
5. Mở sheet *Cân đối Region* và *Cân đối Area*, đảm bảo mọi dòng **Khớp**. Target chốt là cột **Final (Case)** ở sheet *3. NPP*.

## Logic chia

1. **Tổng Region:** cộng Target tháng M+1 của các Area trong file Target theo Region × ShortCode (Htls), quy đổi sang Case bằng tỷ lệ Case/Htls lấy từ SI-SO. SKU không có trong SI-SO quy đổi theo quy cách trong tên SKU, ví dụ `12s (330)` = 12 × 0,33 lít.
2. **Region → Area, tăng trưởng đều:** tỷ lệ tăng trưởng chung = Target M+1 Region ÷ Target tháng M Region. Tổng Target mỗi Area = Target tháng M của Area × (1 + tỷ lệ chung), nên mọi Area tăng cùng tỷ lệ. Cơ cấu SKU trong Area đi từ cơ cấu tháng M, cân đối 2 chiều (RAS) để khớp cả tổng từng SKU của Region lẫn tổng từng Area. Có thể chọn gốc là Sales Out run-rate thay cho Target tháng M.
2b. **Cân YoY lũy kế (khi có file năm trước):** tìm tỷ lệ YoY chung để tới tháng cuối của file Target (thường là tháng 12) mọi Area có cùng tăng trưởng lũy kế so với năm trước. Phần chênh được bù dần cho các tháng còn lại. YoY chỉ tính trên SKU có ở cả 2 năm (sau ghép mã); SKU chỉ có năm nay chia Area theo trend Sales Out năm nay. Giới hạn mặc định: Area chạy nhanh giảm tối đa 5% so với chia đều; Area chạy chậm tăng tối đa 10% so với chia đều và không vượt +10% so với Target tháng M. Sheet *So sánh Area* và *So sánh SKU* đối chiếu số tự cân với số Area trong file Target.
2c. **Cân theo nhóm Brand family × Group (AA/BB/Khác):** Group lấy từ cột Group Repertoire của file Target (Larue, Bivina, TSD thuộc Khác; Strongbow thuộc BB). Mỗi nhóm giữ nguyên tổng Region. Nhóm chiếm từ 5% Region trở lên được cân YoY lũy kế giữa các Area, trần/sàn áp cho từng Area × nhóm; nhóm nhỏ chia tăng trưởng đều tháng. Xem sheet *So sánh nhóm*.
3. **Area → NPP:** `Tỷ trọng NPP = W × Tỷ trọng Target tháng M + (1 − W) × Tỷ trọng Sales Out bình quân N tháng`, mặc định W = 0,5, N = 3. NPP chưa có lịch sử cho SKU: đề xuất 0, trạng thái *Chờ adjust*.
4. **Làm tròn:** mỗi dòng = ROUND(Final cấp trên × tỷ trọng, 0), phần dư dồn vào dòng tỷ trọng lớn nhất (cột *Nhận phần dư* = 1). Công thức nằm sẵn trong Excel.
5. **Adjust 3 cấp:** sheet *1. Region* → *2. Area* → *3. NPP*, mỗi sheet có cột Adjust (ô vàng). Final cấp trên tự chia xuống cấp dưới theo đúng tỷ trọng. Kiểm tra sheet *Cân đối Region* và *Cân đối Area* trước khi gửi số.
6. **Run-rate** mặc định bỏ tháng M vì tháng hiện tại thường chưa đủ dữ liệu.

## Kho dữ liệu GitHub (lưu file để dùng trên mọi thiết bị)

Dữ liệu được **mã hóa ngay trên trình duyệt bằng mật khẩu (AES-256)** trước khi đẩy lên GitHub, nên lưu được cả vào repo Public: trên GitHub chỉ có bản mã hóa (`data/latest/*.enc`, `manifest.json`), người không có mật khẩu không mở được.

**Chuẩn bị một lần**
1. Tạo token: GitHub → Settings → Developer settings → Personal access tokens → **Fine-grained tokens** → Generate new token. Repository access: *Only select repositories* → chọn `Target-Setting-Tool`. Permissions → Repository permissions → **Contents: Read and write**.
2. Mở tool, bấm **Cài đặt** trong khung Kho dữ liệu GitHub: repo mặc định `nppzone2/Target-Setting-Tool`, thư mục `data`; dán token và đặt **mật khẩu mã hóa** (tối thiểu 8 ký tự, dùng chung trên mọi máy).

**Sử dụng**
- **Lưu lên GitHub:** mã hóa và đẩy các file đang chọn, tham số và Adjust vào `data/latest/`.
- **Tải dữ liệu đã lưu:** trên máy khác, nhập token và mật khẩu một lần rồi bấm nút này; tool giải mã, tải lại toàn bộ và tự chạy Chia số.
- Quên mật khẩu thì không khôi phục được dữ liệu đã lưu (chỉ cần lưu lại bằng mật khẩu mới). Token và mật khẩu chỉ nằm trên trình duyệt của từng máy; bỏ chọn *Ghi nhớ* nếu dùng máy chung.
- Lưu dữ liệu vào thư mục `data/` không kích hoạt deploy lại trang (workflow bỏ qua thư mục này).

## Thẻ Chi tiết

Liệt kê kết quả chia theo **Area × SKU** và **NPP × SKU**, có cột **Adjust** nhập trực tiếp trên web. Dòng chưa cân (Area × SKU còn số chưa chia xuống NPP, hoặc NPP nhập vượt) được tô đỏ. Có bộ lọc Area, tìm SKU/NPP và chỉ xem dòng chờ adjust. Sau khi nhập, Final, Cân đối, Tổng quan cập nhật ngay; Adjust được ghi vào file Excel khi tải.

## Thẻ Tổng quan

Sau khi chia, thẻ **Tổng quan** (và sheet *Tổng quan* trong Excel) cho biết Target tháng M+1 tăng trưởng bao nhiêu so với Sales Out cùng tháng năm trước, và lũy kế T1 – M+1 so với cùng kỳ, ở cấp Region, Area, NPP theo Brand family (Heineken, Tiger, Bia Viet, Khác). Cần file SI-SO năm trước. Lũy kế năm nay = Sales Out các tháng đã đóng + tháng M (số lớn hơn giữa SO MTD và Target M nếu chưa đóng) + Target M+1 Final.

Thẻ Tổng quan mở bằng 4 bước để giải thích con số Target, dùng khi trao đổi với sale team: **nền tảng** (Target tháng M và lũy kế đã có), **mùa vụ** (Sales Out cùng tháng năm trước), **đà tăng trưởng** (lũy kế so với năm trước), rồi **Target M+1**. Phần "Vì sao" tóm tắt bằng lời, kèm biểu đồ đà tăng từng Area trước và sau khi cân, danh sách NPP tăng mạnh và NPP lệch run-rate cần xem lại.

## Target M+1 so với thực tế tháng M

Thẻ **Tổng quan** so Target M+1 Final với **Sale In** hoặc **Sales Out** tháng M, dạng cây Region → Area → Brand family → Group AA/BB → SKU, hoặc Area → Distributor → Brand family → Group → SKU. Brand family gồm Heineken, Tiger, Bia Viet và Khác. Chọn đơn vị **Case** hoặc **Htls**. Bấm ▸ để mở từng cấp. Tỷ lệ trên 100% nghĩa là Target tháng sau cao hơn thực tế tháng M.

## Lũy kế đến tháng M+1

Thẻ **Chia số** có bảng **Lũy kế đến tháng M+1** cho từng Region, Area và Distributor: Lũy kế T1–M (thực tế), Target M+1 Final, Lũy kế T1–M+1 và cùng kỳ năm trước. **YoY lũy kế** cho biết cả năm tính đến M+1 tăng bao nhiêu so với năm trước, nên đây là chỉ số đánh giá tổng thể, không chỉ riêng tháng đang chia.
