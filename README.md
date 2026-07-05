# Excel Intellisense XLL - IntelliSense Tooltip & VBA Function Bridge

Add-in Excel (.xll, 64-bit) viết bằng Delphi, giúp các hàm VBA tự viết trong workbook có được **tooltip gợi ý tham số** giống hàm gốc của Excel (SUM, VLOOKUP...), và tùy chọn cho phép gọi hàm VBA **như một hàm Excel thật** (xuất hiện trong AutoComplete/Insert Function, có tooltip đầy đủ) mà không cần build lại add-in mỗi khi thêm hàm mới.

> Lưu ý: đây là bản phát hành dạng **binary (.xll đã build sẵn)**. Repo này **không chia sẻ mã nguồn Delphi** của add-in - chỉ chia sẻ file .xll đã build, file demo, và phần cấu hình XML/VBA mà người dùng có thể tự chỉnh sửa để mở rộng hàm của riêng mình.

---

## 1. Nội dung repo

| File | Mô tả |
|---|---|
| `ExcelIntellisense64.XLL` | Add-in Excel 64-bit đã build sẵn - **không kèm mã nguồn**. |
| `ExcelIntellisense64_Demo.xlsb` | Workbook demo, đã import sẵn `modVBAFunctions.bas`, dùng để test nhanh không cần tự setup. |
| `VBAFunctions.xml` | File khai báo danh sách hàm VBA cần có tooltip / đăng ký làm hàm Excel thật. Người dùng tự sửa file này để thêm hàm mới, **không cần build lại .xll**. |
| `modVBAFunctions.bas` | Module VBA mẫu (4 hàm Cong, Tru, Nhan, Chia + 1 hàm test tham số Range là TinhTongCoDieuKien) minh họa cách viết hàm tương thích với add-in. |

---

## 2. Add-in này làm được gì

### 2.1. Tooltip IntelliSense cho hàm VBA
Khi gõ `=TenHam(` trong 1 ô, add-in hiện 1 tooltip ngay dưới thanh công thức, gồm:
- Tên hàm và danh sách tham số (tham số đang gõ được in đậm/gạch dưới để dễ nhận biết).
- Mô tả của tham số đang gõ.
- Mô tả tổng quát của hàm.

Toàn bộ nội dung tooltip lấy từ file `VBAFunctions.xml` - không cần viết thêm gì trong VBA để có tooltip.

Tooltip cũng tự hiện lại đúng nội dung khi:
- Mở lại file đã lưu rồi bấm vào 1 ô có sẵn công thức để sửa tham số.
- Double-click trực tiếp vào ô có công thức.
- Nhấn F2 để sửa công thức có sẵn.

### 2.2. Cầu nối gọi hàm VBA như hàm Excel thật (tùy chọn)
Nếu muốn hàm VBA xuất hiện trong AutoComplete/Insert Function của Excel như 1 hàm thật (không chỉ có tooltip khi gõ tay đúng tên), add-in đọc thêm `VBAFunctions.xml` để tự đăng ký hàm đó với Excel. Khi được gọi, add-in chuyển tiếp lời gọi sang đúng hàm VBA cùng tên trong workbook đang mở để lấy kết quả - **logic tính toán luôn nằm ở VBA**, add-in chỉ đóng vai trò cầu nối.

Giới hạn của cơ chế cầu nối này:
- Tối đa 20 hàm được đăng ký kiểu này trong 1 lần chạy Excel.
- Tối đa 16 tham số cho mỗi hàm.
- Tham số hỗ trợ cả giá trị thường (số/chuỗi/mảng) và tham chiếu Range thật (đọc được `.Address`, `.Cells`, định dạng...).

---

## 3. Yêu cầu hệ thống

- Windows, Microsoft Excel bản 64-bit (đã test trên Excel 2024 / Windows 11).
- Excel phải cho phép chạy Macro (VBA) và load Add-in .xll.

---

## 4. Cài đặt

### Bước 1 - Chuẩn bị thư mục
Đặt 2 file sau vào **cùng 1 thư mục**:
```
ExcelIntellisense64.XLL
VBAFunctions.xml
```
(có thể đặt bất kỳ đâu, không bắt buộc theo cấu trúc cố định, chỉ cần 2 file này nằm cùng thư mục với nhau)

### Bước 2 - Nạp add-in vào Excel
Cách 1 - Nạp cố định (khuyến nghị, tự load mỗi lần mở Excel):
1. Mở Excel > `File` > `Options` > `Add-ins`.
2. Ở ô `Manage`, chọn `Excel Add-ins` > bấm `Go...`.
3. Bấm `Browse...`, chọn file `ExcelIntellisense64.XLL`.
4. Bấm `OK`, đảm bảo add-in đã được tick chọn trong danh sách.

Cách 2 - Nạp tạm (chỉ dùng cho session hiện tại):
- Double-click trực tiếp vào file `ExcelIntellisense64.XLL` khi Excel đang mở, hoặc kéo-thả file .xll vào Excel.

### Bước 3 - Thêm hàm VBA của bạn
Có 2 cách để test:

**Cách A - Dùng file demo có sẵn:**
Mở `ExcelIntellisense64_Demo.xlsb` (đã import sẵn `modVBAFunctions.bas`), đảm bảo add-in đã nạp (Bước 2), gõ thử `=Cong(` trong 1 ô.

**Cách B - Dùng workbook của riêng bạn:**
1. Mở workbook cần dùng, nhấn `Alt+F11` để mở VBA Editor.
2. `File` > `Import File...` > chọn `modVBAFunctions.bas` (hoặc module VBA của riêng bạn, xem mục 5 để biết quy tắc viết hàm tương thích).
3. Lưu workbook dạng có Macro (`.xlsm` hoặc `.xlsb`).
4. Đóng và mở lại Excel để add-in nạp lại `VBAFunctions.xml`.

> Add-in chỉ đọc `VBAFunctions.xml` **1 lần lúc Excel mở add-in**. Sửa file XML hoặc sửa code VBA xong phải **đóng và mở lại Excel** để nạp lại, không có nút "Reload".

---

## 5. Thêm hàm VBA mới (không cần build lại .xll)

### 5.1. Viết hàm trong VBA
```vb
Private Function TenHam(A As Double, B As Double) As Double
    TenHam = A + B
End Function
```

Lưu ý quan trọng:
- Khai báo hàm là **`Private Function`**, không phải `Public Function`. Nếu để `Public`, Excel sẽ tự động liệt kê hàm này ra làm 1 UDF (User Defined Function) riêng của VBA, gây trùng lặp với hàm mà add-in đăng ký cùng tên - kết quả là gõ `=tenham` sẽ thấy **2 dòng trùng tên** trong gợi ý AutoComplete. Add-in vẫn gọi được hàm `Private` bình thường, chỉ Excel không tự liệt kê nó ra ngoài.
- Hàm phải là `Function` (có giá trị trả về), không dùng `Sub`.
- Tham số có thể là giá trị thường (`Double`, `String`, `Boolean`...) hoặc `Range` (để đọc trực tiếp vùng dữ liệu, định dạng ô...).
- Module VBA chứa hàm phải nằm trong 1 workbook/add-in **đang mở** khi gọi hàm.

### 5.2. Khai báo trong VBAFunctions.xml
```xml
<Function name="TenHam">
  <Description>Mo ta ngan gon chuc nang cua ham</Description>
  <Param name="A">Mo ta tham so A</Param>
  <Param name="B">Mo ta tham so B</Param>
</Function>
```

Quy ước:
- `name` của `<Function>` và `<Param>` không được để trống.
- Tham số tùy chọn: bọc tên trong dấu ngoặc vuông, ví dụ `name="[Dieu_Kien]"`.
- Thứ tự `<Function>`, `<Param>` trong XML là thứ tự hiển thị trong tooltip.
- 1 `<Function>` hoặc `<Param>` khai báo sai (thiếu `name`, XML lỗi cú pháp...) chỉ bị bỏ qua đúng phần đó, không làm mất tooltip của các hàm còn lại trong file.
- Tên trong thuộc tính `name` chính là tên bạn sẽ gõ trong Excel (ví dụ `name="TenHam"` thì gõ `=TenHam(`), và cũng phải khớp đúng tên hàm VBA thật đã viết ở bước 5.1.

### 5.3. Đóng và mở lại Excel
Sửa xong cả VBA và XML, đóng hoàn toàn Excel rồi mở lại để add-in nạp lại danh sách hàm.

---

## 6. Một số lưu ý khi dùng

- Tooltip dựa trên việc bắt từng ký tự gõ trong thanh công thức, nên trong vài trường hợp gõ/sửa công thức phức tạp (dùng phím mũi tên di chuyển giữa các tham số đã gõ, Delete giữa chuỗi...), tooltip có thể tự ẩn đi thay vì hiển thị sai - đây là hành vi có chủ đích (ưu tiên không hiện tooltip sai hơn là hiện sai).
- Nếu thêm quá 20 hàm vào phần "cầu nối gọi như hàm Excel thật" (mục 2.2), các hàm dư sẽ không được đăng ký làm hàm Excel thật (tooltip khi gõ tay vẫn hoạt động bình thường).
- Hàm có quá 16 tham số sẽ không được đăng ký làm hàm Excel thật vì lý do tương tự.

---

## 7. Liên hệ / Đóng góp

Nếu gặp lỗi, có góp ý, hoặc cần hỗ trợ tùy biến thêm cho nhu cầu riêng, liên hệ:

**Kiều Mạnh**
Email: kieumanh366377@gmail.com

---

## 8. Giấy phép

Ghi rõ giấy phép sử dụng (ví dụ MIT, hoặc "chỉ dùng phi thương mại", tùy bạn quyết định) trước khi công khai repo. Phần này để trống, bạn tự điền theo mong muốn.
