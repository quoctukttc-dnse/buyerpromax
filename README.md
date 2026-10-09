# Inbound SAP – Đối chiếu hóa đơn (v11.11 — 10 mẫu riêng + bộ đọc chung + OCR + .xls + .doc)

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
| Capital Tricot (Thái Lan) | `INVOICE_SCAVI_…_CT-26-xxxT_PO.CAP…pdf` — **PDF**: trang INVOICE + các trang PACKING LIST theo kiện (v11.9) | YD |
| Carvico S.p.A. (Ý) | `INV_28493.pdf` (FATTURA/INVOICE) + `PKL_851853_-_INV_28493.pdf`… (một PDF cho mỗi packing list) (v11.10) | M |
| Suzhou Celeb (Trung Quốc) | `INV.xlsx` (sheet `INVOICE`, dòng `surcharge` riêng) + `PKL_1.xlsx`… (sheet `发货码单`, mỗi file một màu/lô) (v11.11) | YD |

Hóa đơn và packing list nằm trong **cùng một file** nên không cần thả thêm gì ngoài file inbound
(`ZMME0032….xlsx` hoặc file SAP xuất ra). Một file inbound dùng được cho **nhiều hóa đơn**.

**Hàng FOC (v11.5):** dòng không ghi thành tiền (ví dụ Yubo ghi "FOC trong roll 8") vẫn được tính vào
số lượng nhập inbound, nhưng **không tính tiền** khi so thành tiền hóa đơn ↔ PO — trước đây phần FOC bị
nhân đơn giá nên báo nhầm "SAI GIÁ TRỊ". Dòng không có thành tiền mà cũng không ghi chữ FOC thì vẫn được
nhắc để kiểm lại.

Khác với trimming: khóa dò là **PO + mã article + màu**; số lượng lấy theo **đơn vị của inbound**;
công cụ **ghi luôn cột `Invoice Quantity`**; lệch thì soi theo **lô / cây vải** thay cho size;
có kiểm **dung sai** (`Over Tolerance Qty`). Chi tiết ở `QUY-TAC-DO.md` phần II.

## Inkava — hóa đơn PDF + packing list Excel theo PO (v10)

Thả `HD xxx.pdf` (hóa đơn GTGT) + các file `DUYxxxxxxx-….xlsx` (một file cho mỗi PO) + file inbound.
Vì packing list Excel có sẵn `Material Code + Size + Spec`, công cụ **điền Invoice Quantity cho
từng dòng inbound** thay vì bắt điền tay, và tự chọn đúng nhóm `Order No` khi packing list gộp
nhiều đơn.

### Packing list Inkava nhiều sheet / tự soạn lại (v11.4)

- **Nhận diện rộng hơn:** không bắt buộc có ô `PO No:` — chấp nhận số PO ghi trơn ở dòng đầu
  (`DUY0081000`) hoặc lấy từ tên file; tiêu đề cột bị dính số 0 (`0 Size`, `0 Spec.`, `0 Order No`) vẫn đọc được.
  Trước đây file kiểu này bị bộ đọc chung hiểu nhầm thành một "hóa đơn" tên `Material Code`.
- **Đọc mọi sheet:** sheet `Sheet` thường là số theo PO, các sheet `89-45`, `51-38`, `10-15`, `7-10`… là số
  **thực giao**. Với từng dòng hóa đơn, công cụ chọn sheet có tổng khớp hóa đơn (ưu tiên sheet thực giao).
- **Dòng lọc ẩn:** thử cả "mọi dòng" và "chỉ dòng đang hiện" (ô Total dùng `SUBTOTAL` bỏ dòng ẩn) —
  phương án nào khớp hóa đơn thì dùng, ghi rõ trong cột ghi chú.
- **Mã gốc 12 ký tự** (`PSTIPAPR0003`) ghép được với mã SAP đầy đủ có đuôi size (`PSTIPAPR0003011`).
- **Không cộng trùng:** dòng inbound khớp đủ Material + Size + Spec được chia trước; dòng còn lại chỉ lấy
  phần packing list còn dư cùng Material + Size.

## Thiên Gia — hóa đơn GTGT PDF + packing list PDF (v11.6)

Thả `C26TTG-xxxxxxxx-….pdf` (hóa đơn GTGT ký hiệu 1C26TTG) + `PKL dd-mm_….pdf` (phiếu giao hàng PGH-…) + file inbound.
Hóa đơn ghi mã hàng là **kích thước** (`TAG PAPER (L100xW70MM) PO TGB0052700`); công cụ dò inbound theo
**PO + kích thước**, cùng PO trùng kích thước thì tách bằng **đơn giá**; packing list có **mã code** (`71423.01`)
khớp cột `Specification` (`code 71423.01 - W26`) → điền Invoice Quantity cho **từng dòng** (kể cả giao một phần).
Không có packing list mà hóa đơn gộp nhiều mã code → báo `LỆCH SL` và nhắc thả kèm packing list.

## Capital Tricot — hóa đơn + packing list PDF theo kiện (v11.9)

Một file PDF duy nhất: trang đầu là **INVOICE**, các trang sau là **PACKING LIST** ghi theo kiện (bale), mỗi dòng một cây.
Công cụ nhận ra theo tên `CAPITAL TRICOT` trong chứng từ và đọc bằng mẫu riêng (không qua bộ đọc chung nữa):

- **Hóa đơn:** mỗi dòng `PO.CAP0022700 N.295/207/73 Usable width 60"/… (44GSM) MISTY ROSE 131/24I 309 YDS. 2.81 USD/YD USD 868.29`
  → PO · design (`N.295/207/73`, `N.295L W SOFT`, `N.295L`) · màu · mã L/D · số yard · đơn giá · thành tiền. Dòng kế tiếp
  không ghi lại PO thì thuộc PO của dòng trên. Số và ngày hóa đơn lấy ở ô `INVOICE NO. / DATE` (`CT-26-314T · AUGUST 19, 2026`).
- **Dò inbound:** PO + design (khớp `Supplier Mat. No.` / `Material Description`: `N.295/207/73 Solid`) + màu (`Color`);
  mã L/D so với `Lapdip Color` để cộng điểm. Số lượng ghi theo YD; giao vượt PO nhưng trong `Over Tolerance Qty` →
  `KHỚP (trong dung sai)`.
- **Packing list theo kiện:** design / PO / L/D của một kiện có thể ghi ở dòng 2–3 (sau cây số 1), kiện sau không ghi gì
  thì kế thừa kiện trước (kiện 6 tiếp kiện 5; kiện 9 lấy design của kiện 8), kiện có thể tràn sang trang sau. Công cụ gom
  đủ các cây rồi mới chốt thuộc tính kiện, cộng các kiện cùng PO + design + màu + L/D để so với dòng hóa đơn
  (kiện 5 + kiện 6 = 300 + 304 = 604). Sheet `CHI TIET LO` liệt kê từng kiện với số cây.
- Kiểm thử với bộ `CT-26-314T` (6 PO, 8 dòng, 9 kiện, 34 cây, 3.297 YD, 7.353,04 USD): 8/8 dòng khớp, packing list khớp
  từng dòng, file `INB_…` xuất ra **trùng 100 %** với file inbound đã điền tay của buyer.

## Carvico — hóa đơn FATTURA PDF + packing list PDF rời (v11.10)

Hóa đơn **không ghi PO SAP** (chỉ `Order 579`); mỗi nhóm hàng mở đầu bằng `00851853 000825 160070 SYDNEY ECO` (số packing list ·
mã article · tên vải), rồi mỗi màu một dòng `WONDERLAND 1E 03261 MT 632,50 5,45 3.447,13` (màu · mã màu · mét · đơn giá · thành tiền,
số kiểu Ý). Số/ngày hóa đơn ở dòng dưới `DOCUMENT No` (`28493 22/07/26`). Công cụ:

- Chọn file inbound theo **Partner Name** (`CARVICO S.P.A.`) vì chứng từ không có PO; dò dòng inbound theo article
  (`825 SYDNEY ECO` ⊂ *Material Description*) + **tên màu** (`WONDERLAND 03261`, `BLACK #9164`, `MYSTIC BLUE # 6063` — mã màu trên
  inbound lúc bỏ 0 đầu lúc không nên chỉ để cộng điểm). Hai dòng hóa đơn cùng màu ở hai packing list (BLACK 3.386,3 + 1.909,4)
  **cộng vào cùng một dòng inbound** (5.295,7).
- Packing list rời: mỗi màu một mục `872470 000825 160070 SYDNEY ECO 003261 WONDERLAND`, từng cây `074178902 1E 741789 CM016902M 70,20 20,50`
  (741789 = số **lô**, chỉ ghi ở cây đầu lô, giữ cho các cây sau kể cả sang trang). Ghép với hóa đơn bằng khóa **số PKL + article + mã màu**
  nên không lẫn hai dòng BLACK; sheet `CHI TIET LO` liệt kê từng lô với số cây (BLACK 851853 = lô 742045: 25 cây 1.686,3 + lô 905704: 27 cây 1.700).
- Kiểm thử bộ `28493` (1 PO `CAR0000700`, 8 dòng hóa đơn → 7 dòng inbound, 2 packing list, 108 cây, 7.043,8 m): khớp toàn bộ,
  packing list khớp từng dòng, file `INB_…` trùng 100 % với inbound buyer đã điền.

## Suzhou Celeb — hóa đơn Excel + packing list Excel theo lô (v11.11)

Hóa đơn `INVOICE NO : CELEB260807-3`, `DATE: 7th,Agu,2026` (tháng viết sai cũng đọc được), bảng `PO NO. | DESCRIPTION OF GOODS | Color |
QUANTITY (Y) | UNIT PRICE | AMOUNT`; dòng **`surcharge | 1 | 150 | 150`** đứng dưới dòng hàng là phụ phí của dòng đó → cộng vào thành
tiền dòng và so với cột `Surcharge Item` của inbound (tổng 4.603,53 khớp). Dòng `TOTAL` cộng cả "số lượng 1" của mỗi dòng phụ phí nên
công cụ trừ ra trước khi so tổng số lượng. Packing list mỗi file một màu (`Lot No. | Roll No. | PO Number | Color | Name | Q'ty(Y)`),
màu viết lẫn ngoặc toàn góc / `TCX` (`True Red（19-1664 TCX）`) → so bằng khóa rút gọn `PO | RC031 | TRUERED191664`; lô = `Lot No.`.

**PO cho phép giao vượt không giới hạn (`Unltd Overdelivery = X`, v11.11):** SAP không đặt mức dung sai cho PO này
(`Over Tolerance Qty` = `Quantity`), nên giao vượt PO không còn báo `VƯỢT DUNG SAI` mà là `KHỚP (trong dung sai)` kèm ghi chú.
Khi `Delivered Qty` đã **đúng bằng** số lượng hóa đơn (hàng nhập kho trước khi đối chiếu hóa đơn, cột `Invoice Quantity` SAP xuất ra
âm) công cụ ghi chú "hàng đã nhập kho trước, chỉ cần ghi số/ngày hóa đơn" và vẫn điền Invoice Quantity = số lượng hóa đơn.

## Chủ hàng chưa huấn luyện → đề nghị liên hệ (v11.7, cột "Chủ hàng" v11.10)

Ngay sau khi thả file, bảng hóa đơn có cột **Chủ hàng**: tên chủ hàng + `mẫu riêng` / `bộ đọc chung`, hoặc nhãn đỏ
**chưa huấn luyện**; dòng đếm phía trên ghi "… · n chủ hàng chưa huấn luyện". Từ v11.10 Partner Name của inbound chỉ được xét
trên **đúng các PO của hóa đơn** (file SAP xuất chung nhiều chủ hàng không còn làm tắt cảnh báo).

Mỗi bộ chứng từ được so tên người bán trên chứng từ và cột `Partner Name` của file inbound với danh sách
`TRAINED_SUPPLIERS` (đầu `src/app-all.js`). Chủ hàng không có trong danh sách vẫn được bộ đọc chung xử lý, nhưng:
nhãn đỏ **chủ hàng mới — chưa huấn luyện** + khung nhắc trên thẻ hóa đơn (tự mở), dòng cảnh báo trong nhật ký, và
cột Kết luận của `BAOCAO_TONGHOP` ghi "CHỦ HÀNG MỚI … — liên hệ …". File không nhận diện / không đọc được cũng kèm lời
nhắc. Người liên hệ đặt ở hằng `CONTACT`. Huấn luyện xong chủ hàng nào thì thêm một dòng vào `TRAINED_SUPPLIERS`.

## Word 97-2003 `.doc` (v11.7)

Đọc chữ ngay trong trình duyệt (CFB của SheetJS + bảng piece của Word) rồi xử lý như PDF. Chứng từ chữ dàn cột cố định
(Chain Guan `CG-BLAO-5704IV.DOC` / `…PLB.DOC`) được tách cột theo khoảng trắng; `.docx` chưa hỗ trợ.

## Bộ đọc chung — chủ hàng chưa có mẫu riêng (v11)

Mọi file Excel/PDF khác (Cheung Hing, Derun, DJIC, Dongguan Uwork, Freetex,
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

- PDF scan đọc bằng OCR nên có thể sai số/sai chữ — luôn kiểm bản gốc; ảnh mờ hoặc có con dấu đè lên bảng (Vanessa) không đọc được. `.docx` chưa hỗ trợ (`.doc` 97-2003 đọc được từ v11.7); `.xls` chuyển tự động nhưng file `.xls` hỏng thì phải mở bằng Excel rồi lưu lại `.xlsx`.
- Bộ đọc chung không biết bố cục trước, nên với chứng từ lạ hãy xem kỹ cột *Cách dò* và ghi chú trong báo cáo;
  dòng không chắc sẽ là CẦN KIỂM TAY / THIẾU DÒNG và **không được điền**.
- Cột trong file inbound tìm theo **tên tiêu đề dòng 1**, nên đổi thứ tự cột vẫn chạy đúng.
- Nếu file nào không nhận diện được, tên file sẽ hiện ở dòng thông báo ngay dưới ô kéo–thả.
