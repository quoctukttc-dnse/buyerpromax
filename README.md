# Inbound SAP – Đối chiếu hóa đơn (v11.28 — 25 mẫu riêng + bộ đọc chung + OCR + .xls + .doc)

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
| Cheung Hing (Hồng Kông, phụ liệu) | `Invoice_to_Scavi_….xlsx` (hangtag / sticker / label, pc · set · doz) + `PACKING_LIST_to_Scavi_….xlsx` theo thùng (v11.12) | PC / SET / DZ |
| Chuangjie (Trung Quốc) | `… INV&PKL CJSQTxxxxxxxx.xlsx` (sheet `Invoice` + `Packing List` kiểu Techwork, dòng `Surcharge of Additional Qty` riêng) (v11.13) | M |
| Yibei (Trung Quốc, mould cup) | `… YIB0004500 ….xlsx` (sheet `IV` + `PL`, size M/L trong mô tả, thùng = CTNS × Qty/CTN) (v11.13) | PAA |
| Luen Hing (Hồng Kông, phụ liệu) | `I 2609-00159.pdf` + `P 2609-00159.pdf` (mỗi thùng một dòng, gộp theo PO + article + lapdip) (v11.13) | M |
| Hoa Nghiêm Vina | `1C26THN_….pdf` (hóa đơn GTGT VND, mã hàng tách 3 dòng) + `PKL ….pdf` (phiếu đóng hàng) (v11.13) | M |
| Best Pacific (Hồng Kông) | `Invoice_H26715746-P6….pdf` (PDF: Cust.No = Material SAP, số lượng = Usable Length mét; Excel `Invoice_HL….xlsx` vẫn đi bộ đọc chung) (v11.14) | M |

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

## Cheung Hing — phụ liệu (hangtag / sticker / label), hóa đơn Excel + packing list Excel theo thùng (v11.12)

Hóa đơn `NO:20260908`, `Date :8/9/2026` (ngày/tháng/năm), bảng `Item | PO | Description | Style-Mã | Quantity | đơn vị | Unit Price | Total`
chia nhóm HANGTAG / STICKER / LABEL; đơn vị `pc` / `set` / `doz` khớp `PC` / `SET` / `DZ` của inbound. Mã hàng là phần sau dấu `-` của
cột Style (`Panache-PAN3ST` → `PAN3ST`, `Envy-7283` → `7283`, `Sculptresse-SCLPST12/SCLPST_10`) → dò trong `Specification` / `Supplier Ref`.

- **Cùng mã, khác đơn vị** (`Evangeline-11434` vừa là sticker tính cái vừa là nhãn tính tá) → lọc dòng inbound theo đơn vị (luật chung mới).
- **Sticker / nhãn có nhiều dòng inbound theo size:** tổng PO các size = số lượng hóa đơn → `KHỚP (chia theo size)`, điền từng size theo PO
  (`Envy-7285A` 3.217 cái chia vào 69 dòng size). Tổng khác số lượng hóa đơn (giao thiếu vài size, packing list không ghi size) →
  `CẦN KIỂM TAY`; **từ v11.12 các dòng này vẫn được giữ trong file `INB_…`** (đã ghi số/ngày hóa đơn, Invoice Quantity để nguyên)
  để buyer điền tay cho đủ, không bị cắt khỏi file nữa.
- **Phí ngân hàng** (`BANK CHARGE 50`) là phụ phí dòng riêng; SAP ghi vào cột `Surcharge Item` của một dòng PO bất kỳ → công cụ so ở
  mức tổng (dòng "Phụ phí ghi trong inbound" trong sheet hóa đơn), không bắt dòng PO đó phải cộng 50 vào thành tiền.
- Packing list theo thùng (`Box # | PO# | Description | Style | mã | Item | Quantity`); dòng không ghi thùng / PO / mô tả kế thừa dòng
  trên; một thùng chứa nhiều PO. Ghép với hóa đơn bằng PO + mã + đơn vị; sheet `CHI TIET LO` liệt kê theo thùng.
- Kiểm thử bộ `20260908` (18 PO, 28 dòng hóa đơn, 233 dòng inbound): 26/28 dòng khớp (219 dòng inbound tự điền trùng 100 % với file buyer
  đã điền), 2 dòng giao thiếu theo size (`11434` sticker 390/394, nhãn 58/73) → `CẦN KIỂM TAY`, 14 dòng giữ lại để điền tay.

## Inkava đi gia công in — hóa đơn GTGT không có PO SAP (v11.16)

Khác bộ Inkava mua phụ liệu (PO SAP `DUY…` + packing list Excel), hóa đơn gia công ghi `PO16431- LCB-3C8072- F2604578 | PCS | 19.037,00 |
4.850 | 92.329.450`: `PO16431` là mã đơn nội bộ của Inkava, không phải PO SAP; PO gia công trên inbound là số thuần (`4500018565`…),
nhiều PO cùng một Material (`SPRIPSZZ0038`), phân biệt bằng *Supplier Ref* `Print S27: 8072 body`. Công cụ lấy **4 chữ số cuối của mảnh
giữa** (`LCB-3C8072` → `8072`) dò trên mọi PO của inbound theo mã + đơn giá, chia theo **PO cũ trước (FIFO)** → `KHỚP (chia nhiều PO)`.

**Nhiều hóa đơn dùng chung một file inbound (luật chung mới):** các hóa đơn được chạy theo thứ tự ngày rồi số hóa đơn; phần một hóa đơn đã
lấy của dòng PO nào được chuyển cho hóa đơn sau nên FIFO nối tiếp đúng và không cộng đôi vào cùng dòng (ví dụ 599 lấy 10.386 của PO
…565 + 8.651 của …566, thì 626 chỉ còn 9.308 ở …566). Kết hợp ô tick gộp → một file `INB_GOP_…` với các dòng PO giao hai lần được nhân dòng.
Bộ thử `1C26TKV#00000599` + `#00000626`: 7/7 dòng ghép đúng mã; 2 cảnh báo đỏ **thật**: mã `8059` hai hóa đơn cộng 6.386 > PO 6.356
(thừa 30), mã `8065` 13.102 > PO 5.584 (có thể còn PO khác chưa nằm trong file inbound).

## Fujian Unitex — hóa đơn Excel có surcharge theo dòng (v11.17)

Hóa đơn Excel một sheet (`第1页`), bảng `CUSTOMER COLOR P/N | CUSTOMER PO. | QUALITY P/N | COMPONENT | WIDTH&WIGHT | COLOUR | MATERIAL NAME |
SURCHARGE | Quantity | Unit Price | Amount`; số hóa đơn nằm ở ô sau nhãn `No.` (`:77FZ10L024SAV`, ô gộp nên nhãn lặp sang các cột kế), ngày sau
`Date.` (`：2026-10-08`, dấu hai chấm toàn góc). Số lượng ghi kèm đơn vị (`330 M`), đơn giá / thành tiền kèm tiền tệ (`USD 2.958`), cột
`SURCHARGE` ghi `USD50` theo dòng và **thành tiền dòng đã gồm surcharge** (113 × 2,62 + 50 = 346,06). Công cụ (`readUnitex`, `src/fab.js`)
đọc PO = *CUSTOMER PO.*, article = *QUALITY P/N*, màu = *COLOUR*, dò inbound theo PO + article + màu; surcharge dòng so với *Surcharge Item*.
Bộ thử `77FZ10L024SAV` + `UNI0000700`: 6/6 dòng đúng (330 / 348 / 397 / 113 / 110 / 126 m); 3 dòng `1SR0206SN` báo `LỆCH GIÁ TRỊ` **thật**:
hóa đơn 2,62 USD/m, PO SAP 2,70 USD/m (lệch 27,92 USD trên cả hóa đơn) — cần hỏi lại chủ hàng / người đặt PO.

## Avery Dennison — nhãn giấy, HĐ GTGT Viettel + DELIVERY NOTE theo size (v11.25)

Hóa đơn GTGT (Viettel) ghi mỗi dòng: Sales Order `149684907`, Delivery # `1947154515`, Customer PO `AVV0117400/522881/234556` (bị ngắt
thành `AVV0117400/5` ở dòng trên và `22881/234556` ở dòng dưới dòng số), tên hàng `PT25-LE`, mã nội bộ `1-323012-000-02`, SL × đơn giá.
DELIVERY NOTE (PDF chữ) chia từng dòng theo `;COLOR:SIZE:QTY` (`;SOFT CLOUD BLUE:2:26`). Inbound SAP: PO `AVV0117400`, Supplier Ref
`PT25-LE`, Specification `522881-PT25-LE-TYP -2-PO 234556` (đơn Scavi · size · số PO). Công cụ: hóa đơn GTGT là chứng từ chính
(`readAveryVat`), DELIVERY NOTE là packing list (`readAveryDn`, ghép theo Delivery #); **luật chung mới `expandBySizes`**: dòng chứng từ ghi
gộp nhiều size mà packing list có nhóm theo size cùng khóa và tổng đúng bằng → tách thành từng dòng theo size trước khi dò; mỗi dòng dò
trong PO theo token {đơn Scavi, số PO} + size. Bộ `1C26TNT#00247665` (3 dòng, 940 cái) + ZMME 279 dòng: 27/27 dòng size `KHỚP`, mỗi dòng
đúng Specification và size, Σ Invoice Quantity 940, tổng tiền 250.980 đ khớp.

**Khách Haddad (v11.26):** cùng Avery nhưng khách khác in khác: PO khách `AVV0138700/948126-UGW` (số đơn + mã màu, bị ngắt `AVV0138700/9`
+ `48126-UGW`), inbound ghi Specification `HAD-RFID-12 STICKER/8126/P1/UGW` (đơn chỉ còn 4 số cuối) → token số dài khớp token ngắn khi
trùng đuôi (`tokEq`); DELIVERY NOTE có thêm cột SHIP SET (`149883106 1 1.1`), chia size dạng `; 8A/XS : 1011` (không có màu) hoặc không chia
size (chỉ `;Total:`) → không chia thì packing list là một nhóm tổng và dòng hóa đơn được tách theo số PO từng size (`splitSize`). Bộ
`1C26TNT#00247861` (5 size, 6.041) và `#00247894` (2 màu × 5 size, 12.074): 15/15 dòng `KHỚP`, mỗi dòng đúng Specification + size.

**Khách Lacoste (v11.27, chưa có inbound):** PO khách chỉ `AVV0128700` (không có đơn / màu), mã nội bộ `CB592682A`, tên hàng `FG001T`, mô tả
`PL.LACOSTE/DEVANLAY.FG001T.traceability..PSR7501T..`; DELIVERY NOTE liệt kê mã nhãn (`;11024549::0001:1001:FG001T:7309939:57G`) thay vì size
→ một nhóm tổng 49.709 = hóa đơn. Dòng hóa đơn mang nhiều mã ứng viên (`altArticles`: FG001T, CB592682A, PSR7501T) để dò Supplier Ref khi
có inbound. Cần file ZMME của PO AVV0128700 để xác nhận SAP dùng mã nào.

**Quy ước số hóa đơn GTGT (chốt 27.10.2026):** luôn ghi `KÝ HIỆU#` + số bù đủ 8 chữ số (`1C26TKV#00000599`, `1C26TFT#00000878`,
`1C26TVN#00000476`) cho mọi chủ hàng — mọi bộ đọc HĐ GTGT đều dùng `padStart(8, '0')`.

## Mainetti — nhãn / móc treo, một PDF gộp nhiều hóa đơn + phiếu xuất kho (v11.23)

Hóa đơn GTGT VNPT của Mainetti (Việt Nam): bảng `STT | Mã hàng hóa (8 số: 55012610) | Tên hàng hóa (Nhãn-DCL-30x120F-348066-4897943-
UNIQUE-S1-V1, có thể ngắt thành dòng trên + dòng dưới của dòng số) | Đơn vị | Số lượng | Đơn giá | Thành tiền`; Đơn đặt hàng là PO Mainetti
`MN10031500` kèm style `R76201A13378`; Lệnh giao hàng `6000…`. **File `HD NHAN.pdf` 57 trang gộp 18 hóa đơn** (21264–21281, mỗi hóa đơn 1–12
trang): công cụ tách theo số HĐ ở đầu trang thành 18 mục riêng (tên file kèm `[HĐ 21264]`), đọc 440 dòng, tổng số lượng và tổng tiền của
cả 19 hóa đơn (kể cả HĐ móc treo 21263) đúng bằng dòng `Tổng (Sub Total)`. Phiếu xuất kho PDF chữ (`PXK-HANGER HUE1.pdf`) đọc thành packing
list theo OD, ghép với hóa đơn có cùng Lệnh giao hàng (2.640 khớp). Phiếu xuất kho nhãn (`PXK NHAN.pdf`, 97 trang scan) chưa OCR vì quá dài
và hóa đơn đã có đủ mã + số lượng.

**Đối chiếu với inbound (v11.24):** SAP đặt PO Mainetti đúng mã `MN10031500` (Purchasing Document), không có mã hàng 8 số của Mainetti;
thay vào đó *Specification* `SS26 348066 S1 V1`, *Customer SKU Code* `4897943`, *Size* (`36`, `2XL`, `ONE SIZE`, `34->46, 50->52`), *DSM
Code*. Mô tả nhãn trên hóa đơn `Nhãn-DCL-30x200-346856-B-5035974-M,L33-S1-V1` được tách thành model `346856`, SKU `5035974`, size `M`
(bỏ đuôi `,L33` / `L34`; `5-6Y`, `18M`, `UNIQUE` giữ nguyên), loại `S1`, phiên bản `V1` → dò trong PO theo tập token {model, S?, V?}
(mục IP), lọc size, cộng 3 điểm khi trùng SKU, trừ 3 khi khác; nhiều dòng cùng điểm (S2/S3 cùng spec) tách bằng số lượng đúng bằng PO.
Móc treo `Móc treo-WIB-10(ANZ405)-…` ↔ Supplier Ref `HANGER ANZ405#WIB-10NB` theo token. Kết quả bộ 19 hóa đơn + ZMME 1.150 dòng: 424/441
dòng `KHỚP`, 16 dòng `LỖI` vì inbound không có PO MN10037500 và MN10039300 (HĐ 21268, 21269), 1 dòng `VƯỢT DUNG SAI` thật (HĐ 21267: hai
dòng S3 của model 357446 — A 470 + B 80 — nhưng PO chỉ có một dòng S3 470). Σ Invoice Quantity và Σ tiền của từng file INB đúng bằng hóa đơn.

## Vanessa — hóa đơn SCAN có kẻ ô, OCR từng ô (v11.22)

Hóa đơn mould cup của Vanessa International (HK) chỉ có bản scan (`SKM_C227i…pdf`, `H9258.pdf`): OCR cả trang đọc được tiêu đề, số HĐ
(`H9424`), PO (`VIH0007800` / mã cũ `VANESSA-3109`), mã hàng (`B221588-2`), màu, cột Amount, nhưng các ô nhỏ **Size / Qty / Unit Price**
bị đọc thành chữ rác (`LL wo ss`). Trước v11.22 báo "PDF dạng ảnh — OCR đọc được chữ nhưng không nhận ra bảng hàng".
Cách mới (`ocrTableRows` trong `src/app-all.js`): vẽ trang ở scale 4 → OCR cả trang lấy toạ độ chữ → mỗi dòng hàng lấy mốc là ô Amount
(`US$24.00`) → trong dải dòng dò **đường kẻ dọc** giữa cột màu và cột Amount để cắt từng ô → OCR từng ô bằng worker riêng (phóng 3–4 lần,
PSM 7/8/10/13, lọc ký tự cho phép, ô một ký tự đọc kiểu một ký tự); ô nào ra dạng đơn giá (`0.65`) là Price, ô bên trái là Qty, trước nữa
là Size; không cắt được ô thì lấy số từ chữ OCR cả trang. `readVanessa` (`src/fab.js`) ràng buộc **thành tiền = SL × đơn giá** để sửa ô
hỏng (mất dấu chấm `04` → 0,4; thiếu SL → SL = thành tiền / đơn giá); dòng `DYEING / LEFTOVER SURCHARGE` → phụ phí mức tổng ↔ Surcharge
Item. Size OCR sai (`0`, `84`, trống) được sửa theo inbound trước khi dò (`fixOcrSizes`: size nào của PO chưa dòng nào nhận, chọn theo đơn
giá / số PO / size còn lại). Mã hàng mất chữ B (`82217994` → `B221799-4`), ngày `SEP42026` (mất dấu cách) đọc được.
Kết quả: `H9424` 15/15 dòng đúng số lượng và size (40 L, 23 XL, 141 M, …); `H9258` + inbound điền tay: 9/9 dòng trùng 100 % (lô 07.10
lên 20/32 bộ trùng). OCR một hóa đơn mất khoảng 60–90 giây, có hiện tiến độ.

## B&B Apparel, Premier Elastic, Pioneer Elastic — ba bộ có inbound điền tay (v11.21)

- **B&B Apparel Accessories (Hồng Kông)** — COMMERCIAL INVOICE mould cup một dòng `QJ0344-44 BBA0006900 B119TR-N (ST009NAA) BLACK 800
  4,732 PR 0.39 1,845.48` (mẫu riêng `readBnB`); dòng tổng lặp `4732.00 1845.48` bị bỏ. Hóa đơn không ghi size, inbound có 5 dòng size
  cùng mã mà tổng PO đúng bằng hóa đơn → `KHỚP (chia theo size)` L 1.366 + M 1.893 + S 841 + XL 526 + XS 106. Khớp 100 % file điền tay.
- **Premier Elastic Webbing (VN)** — HĐ GTGT song ngữ Bkav, 30 dòng `Thun FKS207/11(BLACK-HS20082501)(PRE0039700)`; bộ đọc chung đã ghép
  đúng 30/30 nhưng đọc nhầm số HĐ thành `1C26TPM#00000001` (lấy "1" của `Ký hiệu (Serial No.) : 1C26TPM`) và không ra ngày (`Ngày (day)
  17 tháng (month) 10 năm (year) 2026`). Sửa regex số HĐ / ngày trong bộ đọc chung → `1C26TPM#00001698`, 17.10.2026; khớp 100 % điền tay.
- **Pioneer Elastic (HK)** — hóa đơn `HET7291268 PEA0024400 4000 L590/4 BD7602 L04218/4 54A2 BLACK 093 54A2 7.42 296.80`, đơn giá
  **USD/100 m**, có khi thêm dòng `PLUS SURCHARGE 130.00` (mẫu riêng `readPioneer`): mã hàng ↔ Supplier Ref (`ELAS 0940516/10#EL02927`,
  `EL 02928=0940516/12`, `STRING EL02828=428/5/XA`), màu tách khỏi mã style / Material SAP (`54A2 BLACK 093 54A2` → BLACK 093, mã 54A2),
  đơn giá/m = giá/100, phụ phí dòng cộng vào thành tiền ↔ Surcharge Item. Packing list PDF riêng (`readPioneerPkl`, cùng mã lô HET + PO)
  ghép theo PO + mã + màu. Bộ `MF2616171` (19 dòng, 9 PO): 19/19 đúng, khớp 100 % điền tay (trước đây 5 dòng CẦN KIỂM TAY). Bộ cũ
  `MF2615049` + PKL: 5/5 đúng, 1 cảnh báo thật: dòng PEA0014500 hóa đơn cộng phụ phí 130 USD nhưng PO không có Surcharge Item.

## Inkava phụ liệu — hóa đơn 709: dòng tính theo SET bị bỏ sót (v11.28)

Hóa đơn `1C26TKV#00000709` (8 dòng, 5 PO `DUY0102300/2400/2500/2600/3100`, 5 packing list Excel): file INB xuất ra chỉ có 33 dòng của 3 PO,
thiếu PO DUY0102300, DUY0103100 và dòng `ALABPRWV0277` của DUY0102400 (22.877 bộ). Nguyên nhân: bộ đọc hóa đơn Inkava chỉ nhận dòng số
có đơn vị `Cái / PCS / Chiếc`, còn ba dòng care label ghi đơn vị **SET** nên bị bỏ ngay từ bước đọc (hóa đơn chỉ còn 5 dòng). Sửa `reQty`
nhận thêm `SET / Bộ / Cuộn / Tấm / Đôi / KG / M`. Sau sửa: 8/8 dòng `KHỚP`, 51 dòng inbound điền đủ, Σ Invoice Quantity 66.085 = Tổng số
lượng hóa đơn, tổng tiền 29.631.878 đ khớp. Packing list có cột `Spec.` thay cho `Color` (DUY0102300, DUY0103100) và nhiều dòng cùng size khác
Reference (`B10813` + `B_B10813` + `B10842`) đều gộp đúng theo Material + Size.

## Inkava phụ liệu — hóa đơn 695: mã vật tư bị in tách, packing list không có size (v11.20)

Bộ `1C26TKV#00000695` (34 dòng, 5 PO `DUY…`, 5 packing list Excel, inbound 69 dòng) trước v11.20 báo 20 dòng `THIẾU DÒNG` + 1 `LỆCH SL`:
hóa đơn in mã vật tư tách đôi `ALABPRW V0548002` nên bộ đọc không bắt được mã → ghép theo đơn giá rồi không tìm thấy dòng. Sửa:
(1) ghép lại mã bị tách (`ALABPRW V0548002` → `ALABPRWV0548002`) trước khi bắt mã; (2) size `UNKNOWN` trên inbound / packing list coi như
không có size nên packing list không có cột Size vẫn chia được theo Material + Spec; (3) tổng các dòng lẻ của packing list (1388,456 + …)
được làm tròn 3 chữ số, hết lệch 1725,9999 vs 1726; (4) packing list ≠ hóa đơn nhưng inbound chỉ có **một** dòng cho mã đó → ghi Invoice
Quantity theo hóa đơn, giữ cảnh báo lệch packing list (`LỆCH PKL`: PO DUY0105800 nhãn size 2.345 trên PKL, hóa đơn 2.342).
Kết quả: 33 `KHỚP` + 1 `LỆCH PKL`, 69 dòng inbound điền đủ, tổng 138.569 cái và 28.020.548 đ đúng bằng hóa đơn.

## Kiyohara Việt Nam — túi PE, hóa đơn GTGT song ngữ, PO hệ cũ (v11.19)

Hóa đơn easyinvoice song ngữ: dòng số `1 Cái 16.800 1.183 19.874.400`, tên hàng ở dòng ngay trên (`Túi nhựa PE STONE SHEET S-SS21, KT:`),
mô tả và `(PO: KIYOHARAVN-0003)` ở các dòng dưới; dòng đầu không ghi PO. Mã hàng = từ đầu trong ngoặc mô tả (`SS-21`) ↔ *Supplier Ref*.
PO trên hóa đơn là **mã hệ cũ** (`KIYOHARAVN-0003`), inbound là `KYH0000300`: không có file PO SCAF-SCAX, công cụ suy PO SAP theo số thứ tự
(`0003` → phần số `0000300`, chữ `K·Y·H` là dãy con của `KIYOHARAVN`) và chỉ nhận khi đúng một PO trong inbound thỏa; dòng không ghi PO
dò theo mã + **tổng** chỗ còn nhận của PO (cộng các dòng cùng mã) ưu tiên PO chưa xuất hiện trên hóa đơn. Một PO có nhiều dòng cùng mã
(dòng chính + "mua bổ sung") → `KHỚP (chia nhiều dòng)`: lấp phần còn lại theo số PO từng dòng trước, còn dư mới dùng dung sai, số nguyên.
Bộ thử `1C26TVN#00000476` (5 dòng, 36.300 cái) + inbound KYH0000300/400/500/600: 5/5 dòng ghép đúng PO, 15 dòng inbound được điền,
tổng từng dòng hóa đơn đúng. Lưu ý: file inbound anh gửi đã điền sẵn số HĐ `1C26TVN#475` (không bù 0) trong khi hóa đơn là 476 — cần
xác nhận lại đây là bộ của hóa đơn nào và quy ước ghi số HĐ.

## Inkava — hóa đơn GTGT ghi thẳng PO SAP số thuần (v11.18)

Dạng thứ ba của Inkava (in / thêu / serigraphy theo PO `45000xxxxx`), khác bộ mua phụ liệu (`PODUY…` + packing list) và bộ gia công
(`PO16431- LCB-3C8072- F…`). Mỗi dòng ghi `PO4500014796- <mã tham chiếu>`; mã vật tư SAP (`SPRIPSZZ0045`) có thể cùng dòng, dòng dưới
hoặc không có; phần trong ngoặc `( T14102U-W_QHH PREMIUM)` là chi tiết thêm; một dòng hàng có thể bị in thành 2–3 dòng (PO ở trên, số
ở giữa, mã vật tư ở dưới) và hóa đơn dài nhiều trang (`tiep theo trang truoc`). Công cụ (`readInkavaPo`, `src/fab.js`) gom dòng theo STT,
lấy PO + mã tham chiếu + mã vật tư, rồi dò inbound trong đúng PO theo **tập token** của mã tham chiếu vì Inkava ghi không giống Supplier
Ref của SAP: `ARN-SS27- 006779-` ↔ `ARN-SS27-006779`, `115427-SS26` ↔ `KNIT-SS26-115427`, `353109-SS26- 8891598` ↔
`KNIT-SS26-353109-8891598` (không nhầm với `…-8891599`), `REF_01 Printing L15xW15MM,` ↔ `REF_01`.
Bộ thử 10 hóa đơn `1C26TKV#00000632 … #00000676` (156 dòng, 54 PO) + một file ZMME chung: 156/156 dòng ghép đúng PO + mã; nhiều dòng
cùng PO + mã (HĐ 632: 3 dòng khác mã trong ngoặc, HĐ 649: 5 dòng) cộng dồn vào một dòng inbound; gộp 10 hóa đơn → một file `INB_GOP`
148 dòng, tổng Invoice Quantity từng hóa đơn đúng bằng Tổng số lượng trên hóa đơn. Cảnh báo **thật** còn lại: 10 dòng `VƯỢT DUNG SAI`
(HĐ 638/639/667, ví dụ PO 4500018344 giao 10.043 > 9.858) và 1 dòng `LỆCH GIÁ TRỊ` (HĐ 639: PO 4500013475 `KNIT-FW26-340988` đơn giá
2.200 ≠ PO 2.860 — cùng mã này trên PO 4500018343 đúng 2.200, có thể hóa đơn ghi nhầm PO).

## Best Pacific — hóa đơn PDF, và gộp nhiều hóa đơn vào một file INB (v11.14)

**Lỗi đã sửa:** hóa đơn PDF của Best Pacific có cả cột *Usable Length* (mét) lẫn *N.W.(KG)*; bộ đọc chung lấy nhầm số kg
(217 / 206,6 / 220,1) trong khi inbound tính bằng mét (516 / 504 / 523). Mẫu riêng mới đọc theo *Usable Length* (cột tính tiền),
lấy *Cust.No* làm mã Material SAP để dò thẳng dòng inbound, gộp các batch cùng PO + Material, và ghép được ô Material / màu bị
đẩy sang dòng trên-dưới (`MFKSTCSL000504` + `6`, `POTENT PURPLE 19-` + `2520TCX`). **Các file INB đã xuất trước v11.14 cho chủ hàng này
cần xuất lại.**

**Gộp hóa đơn (ô tick "Gộp tất cả hóa đơn đang chọn vào một file INB", v11.14–v11.15):** công cụ ra **một** file
`INB_GOP_<các số HĐ>_<ngày>.xlsx` cho mọi hóa đơn đang chọn, kể cả khi mỗi hóa đơn ghép với một file inbound riêng (các file inbound
cùng mẫu cột được nối tiếp nhau; khác mẫu cột thì để riêng và báo). Đổi ô tick sau khi đã có kết quả thì công cụ **tự chạy lại** (trước
đây phải bấm lại "Đối chiếu & xuất file", nên tick muộn vẫn thấy 3 file cũ). Mỗi dòng inbound mang đúng số/ngày hóa đơn của mình; dòng PO được từ hai hóa đơn trở lên giao thì được
**nhân thêm dòng** (mỗi dòng một số hóa đơn + số lượng). Dòng nhật ký ghi tổng *Surcharge Item* của các dòng giữ lại để đối chiếu phụ phí
PO. Lưu ý: mỗi hóa đơn vẫn được đối chiếu độc lập, nên dung sai không cộng dồn giữa các hóa đơn trong cùng file gộp — kiểm cột
*Over Tolerance Qty* khi nhiều hóa đơn cùng giao một dòng PO.

Công cụ **không ghi** cột *Surcharge Item* / *Surcharge Header*: giá trị đó là của SAP xuất ra theo từng dòng PO, công cụ chỉ đọc để
đối chiếu với phụ phí trên hóa đơn và in trong sheet hóa đơn của báo cáo.

## Lô huấn luyện 07.10.2026 — 32 bộ chứng từ của 28 chủ hàng (v11.13)

Buyer gửi một thư mục gồm 32 bộ (hóa đơn + packing list + inbound). Công cụ chạy lần lượt từng bộ, so file `INB_…` xuất ra với inbound
buyer đã điền theo khóa PO + Material + Size + Spec. Kết quả sau v11.13: **20/32 bộ trùng 100 %**; 9 bộ inbound gửi kèm là bản SAP xuất
thô (chưa điền) nên chỉ kiểm tra tính hợp lý; 3 bộ còn vướng ở dữ liệu (xem bảng).

| Chủ hàng | Kết quả | Ghi chú |
|---|---|---|
| Capital, Carvico, Cheung Hing, Celeb* | 100 % | mẫu riêng từ v11.9–v11.12 (*Celeb inbound thô) |
| Chuangjie* | hợp lý | mẫu riêng mới; trước bị nhận nhầm là Techwork; 3 dòng `VƯỢT DUNG SAI` thật (827 > 819…) buyer vẫn điền |
| Derun*, AIM High*, Pioneer*, Fujian Honggang*, S&M*, Vinity*, Seamless* | hợp lý | inbound thô; Derun/AIM hết báo đỏ nhờ luật "hàng đã nhập kho trước" |
| DJIC, Dongguan Uwork, Freetex†, Fujian Baikai, Hing Yip, Best Pacific, Junye, Paddies (OCR), Prestige (OCR), Stretchline ×2, Brugnoli (OCR), Chain Guan, PT Winner ×2‡ | 100 % | bộ đọc chung; †Freetex chỉ khác cách ghi số HĐ (`1C26TFT#878` ↔ `1C26TFT#00000878`); ‡PT Winner 115 file đáp án ghi năm 2025, hóa đơn in 2026 |
| Yibei, Luen Hing, Hoa Nghiêm | 100 % | mẫu riêng mới |
| SunPo | không so được | inbound gửi kèm thuộc hóa đơn CI80007160 (10.09), chứng từ là CI80007172 (17.09) — cần đúng bộ |
| Vanessa | không đọc được | PDF scan mờ, OCR không ra bảng hàng — cần file gốc Excel/PDF có chữ |

**Luật chung bổ sung trong v11.13** (áp dụng mọi chủ hàng):

- **Hàng đã nhập kho trước khi đối chiếu:** `Delivered Qty` đúng bằng số lượng hóa đơn → không cộng lô này lần nữa khi tính dung sai
  (Derun, AIM High, Celeb hết báo `VƯỢT DUNG SAI` sai).
- **Dòng cùng màu không đủ chỗ** mà có đúng một dòng khác cùng mã + size có số PO bằng đúng số lượng hóa đơn → dùng dòng đó, ghi chú
  kiểm lại màu (Dongguan Uwork: ô màu gộp nhiều dòng).
- **PO hệ cũ không tra được** (`PT.WINNER1-0667`) hoặc dòng không ghi PO: dò mã + màu + còn chỗ trên mọi PO của inbound, ưu tiên PO
  không xuất hiện trên hóa đơn; chỉ nhận khi còn đúng một PO.
- Cột **Chủ hàng** lấy `Partner Name` của inbound (tên chuẩn SAP) thay vì dòng địa chỉ đọc từ chứng từ; từ khóa `WINNER` bắt được
  `WINNERSUMBIRI`; số hóa đơn bản scan dạng `E 26162` giữ đủ chữ cái.

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

Mọi file Excel/PDF khác (Derun, DJIC, Dongguan Uwork, Freetex,
Fujian Baikai, Fujian Honggang, Hing Yip, Best Pacific, Junye, Pioneer,
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
| **Đối chiếu & xuất file** | Đã có file inbound | Mỗi hóa đơn một file `INB_<số HĐ>_<ngày>.xlsx` để import SAP (hoặc một file `INB_GOP_…` cho các hóa đơn dùng chung inbound khi bật tùy chọn gộp, v11.14) + một file `BAOCAO_TONGHOP_<ngày>.xlsx` |

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
| `CẦN KIỂM TAY` | *(vải)* Nhiều dòng inbound cùng điểm khớp — công cụ không tự điền; các dòng đó vẫn được giữ trong file `INB_…` (có số/ngày HĐ) để điền tay (v11.12) |
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
