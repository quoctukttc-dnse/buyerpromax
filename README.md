# Inbound SAP – Đối chiếu hóa đơn (v11.3 — 6 mẫu riêng + bộ đọc chung + OCR + .xls)

Trang web tĩnh (1 file `index.html`) xử lý **nhiều hóa đơn cùng lúc**: điền số/ngày hóa đơn vào file
inbound SAP và đối chiếu số lượng – đơn giá – thành tiền giữa **hóa đơn – inbound – packing list – PO SCAF-SCAX**.

Toàn bộ xử lý chạy **trong trình duyệt** (ExcelJS + pdf.js nhúng sẵn). Không gửi file lên máy chủ, chạy được offline
(riêng OCR cho PDF scan cần mạng ở lần đầu để tải bộ nhận dạng).

## Cách dùng

1. Mở `index.html` bằng Chrome/Edge (tiêu đề phải có nhãn **v8**).
2. **Kéo cả thư mục** (ví dụ `ATUS-test` gồm nhiều thư mục con, mỗi thư mục một hóa đơn) vào ô lớn,
   kèm file `PO SCAF-SCAX.xlsx`. Có thể kéo thả nhiều lần để bổ sung file.
3. Công cụ tự nhận diện từng file **theo nội dung** (không theo tên) và ghép thành từng bộ
   *hóa đơn + packing list + inbound*, rồi chạy ngay.
4. **Tick chọn hóa đơn** ở cột đầu của bảng — chỉ những hóa đơn được tick mới được xuất.
5. Bấm nút tải tương ứng, hoặc **Tải tất cả**.

### Thêm inbound từng đợt

Khi bạn thả thêm file (ví dụ chỉ một file inbound cho một hóa đơn), công cụ **tự chọn sẵn đúng hóa đơn
vừa được bổ sung file** và bỏ chọn các hóa đơn trước — nên báo cáo chỉ gồm hóa đơn đó, không cộng dồn
các lần trước. Muốn xuất lại nhiều hóa đơn cùng lúc thì tick thêm, hoặc dùng các liên kết nhanh
*tất cả · bỏ chọn · chỉ hóa đơn đã có inbound*. Nếu bạn đổi lựa chọn sau khi đã xuất, sẽ có dòng nhắc
màu cam để bấm chạy lại.

## Chủ hàng vải (v9)

Thả file `.xlsx` chứng từ vải vào cùng chỗ — công cụ tự nhận diện chủ hàng:

| Chủ hàng | File | Đơn vị |
|---|---|---|
| Fujian Techwork | `… INVPKL DASxxxxxxxx….xlsx` (sheet `Invoice` + `Packing List`) | M |
| New Style Vietnam / BLAO | `CHUNG TU - BLAO dd.mm.yyyy.xlsx` (sheet `INV`, `PACKING`, các sheet lô `Y…`) | KG |
| Quanzhou Hengyu | `x.xxHYUxxxxxxx.xlsx` (sheet `invoice`, `packing list`, `码单`) | YD |
| J&H Yubo | `PKL SCAVI ….xlsx` (sheet `PKL`/`PKL Bulk`, mỗi dòng một lô) + (tùy chọn) file PDF hóa đơn GTGT | KG |

Hóa đơn và packing list nằm trong **cùng một file** nên không cần thả thêm gì ngoài file inbound
(`ZMME0032….xlsx` hoặc file SAP xuất ra). Một file inbound dùng được cho **nhiều hóa đơn**.

Khác với trimming: khóa dò là **PO + mã article + màu**; số lượng lấy theo **đơn vị của inbound**;
công cụ **ghi luôn cột `Invoice Quantity`**; lệch thì soi theo **lô / cây vải** thay cho size;
có kiểm **dung sai** (`Over Tolerance Qty`). Chi tiết ở `QUY-TAC-DO.md` phần II.

## Inkava — hóa đơn PDF + packing list Excel theo PO (v10)

Thả `HD xxx.pdf` (hóa đơn GTGT) + các file `DUYxxxxxxx-….xlsx` (một file cho mỗi PO) + file inbound.
Vì packing list Excel có sẵn `Material Code + Size + Spec`, công cụ **điền Invoice Quantity cho
từng dòng inbound** thay vì bắt điền tay, và tự chọn đúng nhóm `Order No` khi packing list gộp
nhiều đơn.

## Bộ đọc chung — chủ hàng chưa có mẫu riêng (v11)

Mọi file Excel/PDF khác (Capital, Carvico, Celeb, Cheung Hing, Derun, DJIC, Dongguan Uwork, Freetex,
Fujian Baikai, Fujian Honggang, Hing Yip, Hoa Nghiêm, Best Pacific, Junye, Luen Hing, Pioneer,
PT Winner, S&M, Seamless, Stretchline, SunPo, Yibei, Brugnoli, AIM High, Chain Guan…) được đọc bằng
bộ đọc chung: tìm bảng hàng bằng từ khoá tiêu đề, lấy số lượng/đơn vị/giá/PO/mã/màu/size/lô, rồi dò
dòng inbound theo thứ tự **Material SAP → PO → mã hàng → màu/spec/lapdip → size**. Packing list rời
(Excel/PDF) tự gắn vào hoá đơn cùng thư mục; hoá đơn ghi gộp theo mã còn packing list ghi theo PO/size
thì được tách theo packing list. Chi tiết luật ở `QUY-TAC-DO.md` phần IV.

File Excel 97-2003 `.xls` được tự chuyển sang `.xlsx` trong trình duyệt (SheetJS nhúng sẵn, chạy offline) rồi đọc
như bình thường — dòng thông báo ghi "đã tự chuyển sang .xlsx". Không đọc được và công cụ sẽ báo: file Word `.doc`.

## PDF dạng ảnh (scan) — OCR (v11.1)

PDF không có lớp chữ được nhận dạng bằng tesseract.js (tải từ CDN jsDelivr ở lần đầu, ~8–9 MB kể cả dữ liệu
tiếng Anh + Việt, sau đó trình duyệt giữ lại). Ô **OCR** cạnh ô kéo–thả cho phép tắt. Kết quả đi qua bộ đọc
chung như PDF thường, được gắn nhãn **OCR** và cảnh báo đỏ trong báo cáo; bản scan trùng với file gốc
cùng thư mục bị bỏ; bản scan không nhận ra bảng hàng sẽ báo để nhập tay. Luôn đối chiếu lại số liệu OCR với
bản gốc trước khi import. Chi tiết ở `QUY-TAC-DO.md` mục G9.

## Hai chế độ

| Nút | Khi nào dùng | Kết quả |
|---|---|---|
| **Xuất danh sách PO** | Chưa có file inbound (cần số PO để tạo inbound trên SAP) | `DANHSACH_PO_<ngày>.xlsx` — sheet `PO` đúng 2 cột `PO No.` \| `PO No ScaX`; sheet `CHI TIET` liệt kê theo từng hóa đơn/item |
| **Đối chiếu & xuất file** | Đã có file inbound | Mỗi hóa đơn một file `INB_<số HĐ>_<ngày>.xlsx` để import SAP + một file `BAOCAO_TONGHOP_<ngày>.xlsx` |

Thiếu packing list thì chỉ đối chiếu hóa đơn ↔ inbound. Thiếu inbound thì đối chiếu hóa đơn ↔ packing list
và ghi rõ *CHƯA CÓ INBOUND*.

## File xuất ra

- `INB_…xlsx` — **file import SAP**: giữ nguyên 100% cột và sheet của template, chỉ điền
  `Invoice Number` (ký hiệu + `#` + số đệm 8 chữ số) và `Invoice date` (`dd.mm.yyyy`, dạng text).
  Mặc định chỉ giữ các dòng thuộc hóa đơn (Invoice Quantity > 0).
- `BAOCAO_TONGHOP_…xlsx` — sheet `TONG HOP` (mỗi hóa đơn một dòng, kết luận + tiền),
  `CHI TIET` (mọi dòng hóa đơn), `LECH SIZE`, `DANH SACH PO`, và một sheet riêng cho từng hóa đơn.
- Tick *"Xuất thêm file báo cáo chi tiết cho từng hóa đơn"* nếu cần bản đầy đủ theo từng dòng inbound
  (kèm cột `Amount` = `=V*M+N`, `Balance` = `=R-(V+S)` và các cột đối chiếu).

## Quy tắc đối chiếu

| Bước | Cách làm |
|---|---|
| Ghép bộ | Chấm điểm mọi cặp rồi gán từ cặp điểm cao nhất xuống: cùng thư mục (6đ) → tên file chứa số hóa đơn (3đ) → cùng *Despatch Note* (4đ, packing list) → tỉ lệ trùng PO (tối đa 3đ, inbound) |
| PO | Mã PO trong dòng mô tả hóa đơn → tra cột **PO No ScaX**; không thấy thì dò cột **PO No.** (PO đã là ScaF) → lọc cột A file inbound |
| Item | Mã hàng giữa `/` và `//` trong dòng mô tả (kể cả khi bị xuống dòng) → sinh biến thể từ chi tiết đến chung (`LB 5731 Main C/509` → `LB5731MAINC/509` · `LB5731C/509` · `LB5731`), thử lần lượt trong **Material Description** + **Specification** (bỏ dấu cách, không phân biệt hoa thường), rồi lọc bằng từ khóa trong mã (`Main`/`Care`/`Angel Pink`…) |
| Hai item trùng mã trong cùng PO | Chấm điểm quyền sở hữu từng dòng inbound: biến thể càng chi tiết càng cao · có từ khóa của mình +2 · mang từ khóa của item khác −6. Dòng thua đi tìm lại trong phần chưa ai chiếm → `LB 5731 Main C/509` lấy dòng "Main label…", `LB 5731 C/509` lấy dòng "LB care label…" |
| Vẫn mơ hồ | Chọn nhóm có tổng số lượng khớp hóa đơn (ghi rõ trong ghi chú); không nhóm nào khớp → **cảnh báo mơ hồ** kèm các nhóm còn lại để kiểm tra tay |
| Mã hàng không có trong inbound | Ghép theo **đơn giá** trong cùng PO (khi cặp PO + đơn giá là duy nhất trong hóa đơn), có ghi chú rõ |
| Size | Nhóm size packing list `<nội bộ>/<quốc tế>`: `XS/XP`→XS, `S-DD/P-DD`→S-DD, `XL/XXL / XG/XXG`→XL/XXL |
| Giá trị | Đơn giá hóa đơn vs **Gross Price**; thành tiền vs `Σ(Invoice Qty × Gross Price + Surcharge)` — lệch thì cảnh báo **đỏ** |

## Các trạng thái

| Trạng thái | Nghĩa |
|---|---|
| `KHỚP` | Số lượng, đơn giá, thành tiền khớp |
| `LỆCH GIÁ TRỊ` | **Đỏ** — đơn giá/thành tiền hóa đơn khác PO, cần kiểm tra lại hóa đơn |
| `LỆCH SL` | Số lượng inbound khác hóa đơn |
| `LỆCH PKL` | Inbound khớp hóa đơn nhưng packing list lệch (có chi tiết theo size) |
| `CHƯA ĐIỀN SL HĐ` | Cột `Invoice Quantity` của nhóm dòng đó đang trống — điền rồi chạy lại, hoặc tick *"Dùng cột Quantity khi Invoice Quantity còn trống"* |
| `THIẾU DÒNG` | Hóa đơn có nhưng inbound không có dòng nào (kèm số liệu lấy từ file PO) |
| `CHƯA CÓ INBOUND` | Chưa tạo inbound trên SAP cho hóa đơn này |
| `KHỚP (giao thiếu)` | *(vải)* Giao ít hơn PO — bình thường, giao từng đợt |
| `KHỚP (trong dung sai)` | *(vải)* Giao vượt `Quantity` nhưng còn trong `Over Tolerance Qty` — hợp lệ |
| `VƯỢT DUNG SAI` | *(vải)* **Đỏ** — vượt `Over Tolerance Qty`, SAP sẽ báo lỗi khi import |
| `SAI ĐƠN VỊ` | *(vải)* Hóa đơn không có con số cùng đơn vị với `Base Unit of Measure` của inbound |
| `CẦN KIỂM TAY` | *(vải)* Nhiều dòng inbound cùng điểm khớp — công cụ không tự điền |
| `KHỚP (chia nhiều PO)` | *(Yubo)* Cùng mã Material có ở nhiều PO — đã chia theo PO cũ trước (FIFO), xem sheet `PHAN BO PO` |
| nhãn `thiếu ký hiệu HĐ` | *(Yubo)* Ô `HD:` chỉ có số — thả kèm file PDF hóa đơn GTGT là điền đủ ký hiệu + số |
| nhãn `lệch hóa đơn GTGT` | Tổng tiền hàng trên PKL khác `Cộng tiền hàng` trên hóa đơn PDF — kiểm tra lại |
| nhãn `lệch tổng SL` | Dòng `Tổng cộng` trên chứng từ khác tổng `Invoice Quantity` đã ghi vào file inbound — kiểm tra lại |
| `LỖI` | Không tra được PO (vải: thử thả thêm file PO SCAF-SCAX) |
| `THIẾU FILE PO` | Mã PO trên hóa đơn chưa đúng dạng SAP và chưa có file PO SCAF-SCAX để tra |

## Đưa lên GitHub Pages

```bash
git init && git add index.html README.md && git commit -m "Inbound SAP tool v6"
git branch -M main
git remote add origin https://github.com/<tài-khoản>/<tên-repo>.git
git push -u origin main
```

Rồi vào **Settings → Pages** → Source `Deploy from a branch` → Branch `main` / `/ (root)` → **Save**.

## Giới hạn

- PDF scan đọc bằng OCR nên có thể sai số/sai chữ — luôn kiểm bản gốc; ảnh mờ hoặc có con dấu đè lên bảng (Vanessa) không đọc được. `.doc` chưa hỗ trợ; `.xls` chuyển tự động nhưng file `.xls` hỏng thì phải mở bằng Excel rồi lưu lại `.xlsx`.
- Bộ đọc chung không biết bố cục trước, nên với chứng từ lạ hãy xem kỹ cột *Cách dò* và ghi chú trong báo cáo;
  dòng không chắc sẽ là CẦN KIỂM TAY / THIẾU DÒNG và **không được điền**.
- Cột trong file inbound tìm theo **tên tiêu đề dòng 1**, nên đổi thứ tự cột vẫn chạy đúng.
- Nếu file nào không nhận diện được, tên file sẽ hiện ở dòng thông báo ngay dưới ô kéo–thả.
