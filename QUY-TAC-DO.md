# Nguyên tắc dò của công cụ Inbound SAP (v11.3)

Tài liệu này mô tả **chính xác** cách công cụ tìm và so khớp dữ liệu, để bạn kiểm tra lại được
mọi con số nó đưa ra. Không có "thông minh" gì cả — chỉ là một chuỗi luật ưu tiên, chạy theo
đúng thứ tự dưới đây, dừng ở luật đầu tiên cho kết quả.

---

## 0. Chuẩn hóa chuỗi — nền tảng của mọi phép so

Trước khi so bất cứ thứ gì, chuỗi được đưa về dạng chuẩn:

```
norm(s) = viết HOA toàn bộ + xóa MỌI khoảng trắng
```

Ví dụ: `"Main label LB 5731 C/509"` → `MAINLABELLB5731C/509`

Nhờ vậy `LB 5731`, `LB5731`, `lb  5731` đều như nhau. Dấu `/`, `-`, `#` **được giữ nguyên**
vì chúng mang thông tin (`C/509`, `C/1`, `54A2#QB4`).

Số tiền/số lượng: `5.249` và `1,034` đều hiểu là dấu phân cách nghìn → `5249`, `1034`.

---

## 1. Nhận diện loại file (theo nội dung, không theo tên)

Mỗi file được chấm điểm bằng cách đếm số cụm từ đặc trưng xuất hiện bên trong:

| Loại | Cụm từ tìm |
|---|---|
| Hóa đơn | `HÓA ĐƠN GIÁ TRỊ GIA TĂNG`, `VAT INVOICE`, `KÝ HIỆU`, `TIỀN THUẾ`, `NGƯỜI BÁN HÀNG`, `ĐVT` |
| Packing list | `PACKING LIST`, `DESPATCH NOTE`, `LABEL REF`, `CTN NO`, `OUR REF`, `DELIVERY METHOD` |
| File inbound | `purchasing document`, `invoice quantity`, `over tolerance qty`, `act. gds mvmnt date`, `external delivery id` |
| PO SCAF-SCAX | `po no scax`, `po follow up`, `order quantity`, `agreed lead-time`, `pr no.` |

Điểm cao hơn thắng. Đó là lý do file tên `IN C26TVN2088 PKL DBVN229768.pdf` (có chữ "PKL")
vẫn được nhận đúng là hóa đơn.

## 2. Ghép hóa đơn ↔ packing list ↔ inbound

Chấm điểm **mọi cặp** rồi gán từ cặp điểm cao nhất xuống (không duyệt lần lượt từng hóa đơn,
vì như thế hóa đơn đứng trước sẽ giành mất file của hóa đơn sau):

| Tín hiệu | Điểm |
|---|---|
| Cùng thư mục | +6 |
| Tên file chứa số hóa đơn (`2088`) | +3 |
| Cùng số *Despatch Note* (`DBVN229768`) — chỉ packing list | +4 |
| Tỉ lệ PO trùng nhau — chỉ inbound | +0 → +3 |

Nếu cả hóa đơn và file đều có thư mục mà **khác thư mục** → loại thẳng, trừ khi tên file
chứa số hóa đơn.

---

## 3. Đọc hóa đơn PDF

PDF không có "ô" như Excel — pdf.js chỉ trả về từng mẩu chữ kèm tọa độ (x, y). Công cụ gom các
mẩu có cùng `y` (sai số ≤ 3pt) thành một dòng, sắp theo `x`, dựng lại bố cục gần giống bản in,
rồi mới bóc dữ liệu.

**Số hóa đơn** = `Ký hiệu` + `#` + `Số` đệm 8 chữ số → `1C26TVN` + `2204` → `1C26TVN#00002204`

**Ngày** = `Ngày (date) 08 tháng (month) 09 năm (year) 2026` → `08.09.2026` (ghi dạng text)

**Mã hàng** = phần nằm giữa dấu `/` **đầu tiên** và `//`:

```
Nhãn vải đã in/ LB 5731 Main C/509 // TRI0007000
                └──── mã hàng ────┘    └─ PO ─┘
```

Hai trường hợp bất thường đã xử lý:
- **Mã bị xuống dòng**: `… thông minh/ LM-` ⏎ `RFIDST22 // TRIMMINGVN-0708` → nối lại thành `LM-RFIDST22`.
- **Mã nằm SAU dòng số lượng** (bố cục 3 dòng của hóa đơn Viettel) → gán ngược lên dòng vừa tạo.

**Số lượng / đơn giá / thành tiền** = ba số liên tiếp sau chữ `Cái`.

---

## 4. Tra PO

1. Lấy **mọi mã** dạng chữ-số trong dòng mô tả.
2. Tra lần lượt trong cột **PO No ScaX** của `PO SCAF-SCAX.xlsx` → ra **PO No.** (`TRIMMINGVN-0725` → `TRI0004600`).
3. Không thấy → tra tiếp trong cột **PO No.** (trường hợp hóa đơn ghi thẳng PO ScaF như `TRI0007000`).
4. Vẫn không thấy → trạng thái `LỖI`, ghi rõ đã dò cả hai cột.

PO tìm được dùng để **lọc cột A** (`Purchasing Document`) của file inbound. Mọi bước sau chỉ
làm việc trong phạm vi các dòng của PO đó.

---

## 5. Tìm đúng dòng inbound cho một item — phần quan trọng nhất

### 5.1. Sinh biến thể mã hàng (chi tiết → chung)

Mã trên hóa đơn và mã trong SAP hầu như không bao giờ viết giống nhau, nên công cụ sinh ra
nhiều cách viết rồi thử lần lượt. Với `LB 5731 Main C/509`:

| # | Biến thể | Cách tạo |
|---|---|---|
| 1 | `LB5731MAINC/509` | nguyên mã, bỏ dấu cách |
| 2 | `LB5731MAINC509` | bỏ luôn dấu `/` |
| 3 | `LB5731MAINC0509` | đổi `/` thành `0` (SAP hay viết `C01` thay cho `C/1`) |
| 4 | `LB05731MAINC/509`… | thêm số 0 vào phần số |
| 5 | `LB5731` | **gốc** = chữ đầu + số, bỏ hết phần đuôi |
| 6 | `LB05731` | gốc có đệm 0 |

Thử từ trên xuống trong `Material Description` **+** `Specification` (đã chuẩn hóa), **dừng ở
biến thể đầu tiên tìm thấy dòng nào**. Càng ở trên càng chắc chắn.

> Đây là cách `LB 07780 C/1` (hóa đơn) khớp được với `LB7780C01` (SAP).

### 5.2. Lọc bằng từ khóa trong mã

Sau khi có tập dòng, lọc tiếp bằng **mọi từ ≥ 3 chữ cái** có trong mã hàng: `MAIN`, `CARE`,
`ANGEL`, `PINK`… Chỉ lọc nếu kết quả còn lại **không rỗng và ít hơn trước** — nghĩa là lọc
không bao giờ làm mất sạch dữ liệu.

Ví dụ thật của hóa đơn 2204:

```
"LB 5731 Main C/509" · PO TRI0007000
  biến thể trúng: LB5731          (vì SAP viết "Main label LB 5731 C/509" — chữ Main đứng TRƯỚC)
  trước lọc:  [LB care label 5731 C/509]  +  [Main label LB 5731 C/509]
  từ khóa lọc: MAIN
  → chọn: Main label LB 5731 C/509 = 1 dòng, 7.071 pcs   ✓ đúng số lượng hóa đơn
```

### 5.3. Phân xử khi hai item cùng PO giành nhau một dòng

Đây là chỗ luật 5.1 + 5.2 chưa đủ: `LB 5731 C/509` (không có chữ Main) vẫn nằm gọn bên trong
`MAINLABELLB5731C/509`, nên nó có thể bắt nhầm dòng Main.

Khi trong cùng một PO có **nhiều dòng hóa đơn cùng giành một dòng inbound**, chấm điểm quyền
sở hữu:

| Tín hiệu | Điểm |
|---|---|
| Biến thể mã khớp càng chi tiết (vị trí càng trên bảng 5.1) | 10 → 2 |
| Dòng inbound có chứa từ khóa của **chính mã này** | +2 mỗi từ |
| Dòng inbound mang từ khóa của **item khác** trong cùng PO | **−6** mỗi từ |

Dòng inbound thuộc về item điểm cao nhất; item thua **đi tìm lại** trong phần chưa ai chiếm.

```
PO TRI0006400 có cả "Main label LB 5731 C/509" (10.628) lẫn "LB care label 5731 C/509" (8.736)

"LB 5731 Main C/509" : biến thể LB5731 (4đ) + MAIN có mặt (+2) = 6
"LB 5731 C/509"      : biến thể LB5731C/509 (10đ) − MAIN của item khác (−6) = 4
→ dòng Main thuộc về item Main; item còn lại tìm lại → nhận dòng care 8.736 ✓
```

Kết quả **không phụ thuộc thứ tự** hai dòng trên hóa đơn.

### 5.4. Vẫn mơ hồ → dùng số lượng, hoặc cảnh báo

Nếu chỉ có **một** dòng hóa đơn (không có item nào để so kè) mà inbound vẫn còn nhóm khác cùng
họ mã:

- Gom các nhóm theo mô tả, tính tổng số lượng từng nhóm.
- **Đúng một nhóm** có tổng bằng số lượng hóa đơn → chọn nhóm đó, ghi rõ *"đã chọn nhóm có tổng
  số lượng khớp hóa đơn"* trong ghi chú.
- Không nhóm nào khớp (hoặc nhiều nhóm cùng khớp) → **không đoán**, gắn cảnh báo
  *"Mã hàng chưa đủ phân biệt"* kèm danh sách nhóm còn lại và số lượng của chúng để bạn kiểm tra tay.

### 5.5. Mã hàng không hề có trong inbound → ghép theo đơn giá

Một số mã nội bộ của ITL không xuất hiện ở đâu trong file inbound (ví dụ `LM-RFIDST22` trong khi
SAP ghi *"MYSTERE Sticker … Le Mystere Rfid Sticker"*). Khi đó, nếu cặp **(PO + đơn giá)** là
duy nhất trong hóa đơn, công cụ lấy toàn bộ dòng của PO có đúng `Gross Price` đó, và ghi chú
*"Ghép theo đơn giá"*. Không duy nhất → không ghép, báo `THIẾU DÒNG`.

---

## 6. Số lượng của hóa đơn

- Số lượng lấy từ cột **Invoice Quantity (V)**.
- Nếu **cả nhóm** dòng đó đều trống → trạng thái `CHƯA ĐIỀN SL HĐ`, **không** tự lấy cột `Quantity`
  (trừ khi bạn tick tùy chọn tương ứng). Dòng có `V = 0` trong nhóm đã điền được coi là **không
  thuộc hóa đơn** và không được ghi số hóa đơn.

---

## 7. Đối chiếu packing list theo size

Packing list ghi size dạng `<size nội bộ>/<size quốc tế>`, có thể xuống dòng giữa chừng:

| Trên packing list | Hiểu là |
|---|---|
| `XS/XP` | `XS` |
| `S-DD/P-DD` | `S-DD` |
| `XL/XXL / XG/XXG` | `XL/XXL` |
| `M-DD/M-` ⏎ `DD` | `M-DD` |

Quy tắc: nếu có dấu `/` **có khoảng trắng hai bên** thì cắt ở đó; nếu không thì cắt ở dấu `/`
đầu tiên. Phần bên trái là size, đem so thẳng với cột `Size` của inbound.

Mã invoice của packing list (`54A2`, `7VWB`, `34Y5`…) lấy từ `Specification` theo mẫu `/<MÃ>#`,
dùng để chỉ so đúng nhóm hàng tương ứng.

---

## 8. Đối chiếu giá trị

| So cái gì | Cách so | Khi lệch |
|---|---|---|
| Đơn giá | Đơn giá hóa đơn ↔ `Gross Price` của các dòng đã ghép | `LỆCH GIÁ TRỊ` — **cảnh báo đỏ** |
| Thành tiền | Thành tiền hóa đơn ↔ `Σ (Invoice Qty × Gross Price + Surcharge)` | chỉ so **khi số lượng đã khớp**, vì lệch SL thì đương nhiên lệch tiền |
| Tổng tiền | `Cộng tiền hàng` trên hóa đơn ↔ tổng inbound của **các dòng đã điền SL** | banner đỏ nếu do sai đơn giá; banner vàng nếu do thiếu dòng / lệch SL |

---

## 9. Thứ tự kết luận trạng thái

Xét lần lượt, dừng ở điều kiện đầu tiên đúng:

1. Không tra được PO → `LỖI`
2. Chưa có file inbound → `CHƯA CÓ INBOUND`
3. Không tìm được dòng nào → `THIẾU DÒNG`
4. Đơn giá sai (hoặc thành tiền sai khi SL đã khớp) → `LỆCH GIÁ TRỊ` 🔴
5. Cả nhóm chưa điền Invoice Quantity → `CHƯA ĐIỀN SL HĐ`
6. SL inbound ≠ SL hóa đơn → `LỆCH SL`
7. Inbound khớp hóa đơn nhưng packing list lệch → `LỆCH PKL`
8. Còn lại → `KHỚP`

---

## 10. Những gì công cụ **không** làm (phần trimming)

- Không sửa file inbound ngoài hai cột `Invoice Number` và `Invoice date`.
- Không tự thêm dòng vào file inbound, kể cả khi file PO có dòng đó.
- Không tự chia số lượng theo size.
- Không đoán khi mã hàng mơ hồ — luôn báo để bạn quyết.
- Không đọc được PDF scan thành ảnh (không có lớp text).


---
---

# PHẦN II — CHỦ HÀNG VẢI (từ v9, 28.09.2026)

Vải đi theo một đường riêng: chứng từ là **Excel**, không phải PDF, và hóa đơn + packing list
nằm **trong cùng một file**. Vì vậy phần này không dùng lại bất kỳ luật nào của mục 3–7 ở trên;
phần trimming hoàn toàn không đổi.

## V0. Chuẩn hóa cho vải

```
AZ(s)    = viết HOA + bỏ MỌI ký tự không phải chữ/số
WORDS(s) = các từ chữ-số dài ≥ 3 ký tự
```

Khác với `norm()` của trimming (giữ `/ - #`), vải **bỏ luôn** dấu `/ - _` vì mỗi bên viết một
kiểu: `IVORY ACH658-110` ↔ `IVORY ACH 658-110`, `BLACK / S0381-A` ↔ `BLACK-S0381-A`.

Hai bẫy của file Excel chứng từ, đã xử lý:

- **Ô rich-text** trả về object chứ không phải chuỗi → phải ghép các đoạn `richText` lại.
- **Ô gộp dọc**: dòng dưới của vùng gộp vẫn trả về giá trị của dòng trên, làm một dòng hàng bị
  đếm hai lần (BLAO). Công cụ **chỉ đọc ô gộp ở đúng dòng chủ**, các dòng còn lại coi như trống.

## V1. Nhận diện chủ hàng (theo nội dung, không theo tên file)

| Chủ hàng | Dấu hiệu |
|---|---|
| `TECHWORK` — Fujian Techwork | có sheet `Packing List` + trong 30 dòng đầu có `QTY/M` và `SURCHARGE` |
| `BLAO` — New Style Vietnam | có `11. PO Number` và `Quantity/Unit` |
| `HYU` — Quanzhou Hengyu | có `Marks & Nos.` và `Inv. No.` |

Không khớp cả ba → file được đưa trở lại quy trình nhận diện của trimming (inbound / PO / hóa đơn PDF).

## V2. Đọc chứng từ

| | TECHWORK | BLAO | HYU |
|---|---|---|---|
| Số & ngày HĐ | nhãn `INVOICE NO.` / `INVOICE DATE`, giá trị ở **dòng ngay dưới** | nhãn `8. No & Date of invoice`, dòng dưới | `Inv. No.:` và `Date:` trên cùng một dòng |
| Bảng hàng | dòng có `PO NO.` … `AMOUNT` | dòng có `11. PO Number` | dòng có `Marks & Nos.` |
| Đơn vị | lấy từ tiêu đề `QTY/M` → **M** | không ghi → **KG** (số thứ nhất), mét là số thứ hai | dòng dưới tiêu đề: `Yards` → **YD** |
| Phụ phí | **một cột** `SURCHARGE($)` | **một dòng riêng** `SURCHARGE:` | không có |
| Thành tiền HĐ | `SL × đơn giá + phụ phí` | `SL × đơn giá` | `SL × đơn giá` |
| Chi tiết cây | sheet `Packing List`: `Lot No.` + `Rolls #` + `Ttl/M` | sheet `PACKING` chỉ tổng theo màu; chi tiết cây nằm ở các sheet lô `Y92296`… | sheet `码单`: `LOT#` + `Roll/No.` + `Yard` |

Ngày tháng nhận cả 4 dạng: serial Excel (`46282` → 17.09.2026), ô Date thật, `dd/mm/yyyy`,
và `Sep 28th` — dạng cuối thiếu năm nên **lấy năm từ serial/ngày khác trong cùng workbook**.

Cột PO, mã article, màu, lô trong packing list được **kéo xuống (forward-fill)** khi ô trống,
đúng như cách người ta trình bày bảng.

## V3. Chuẩn hóa số PO — luật bù 0

```
poSap(x) = chữ cái + phần số, bỏ 0 đầu rồi bù lại cho đủ 7 chữ số
```

- `TEC002400` → `TEC0002400` ✔ (hóa đơn Techwork thiếu một số 0 — lỗi thật, đã gặp)
- `SRV0000400`, `HYU0001700`, `TEC0001500` → giữ nguyên

Nếu sau khi bù 0 **vẫn** không có PO đó trong file inbound, công cụ thử tra qua file
`PO SCAF-SCAX.xlsx` (nếu có). Không có file đó thì báo `LỖI` kèm câu nhắc thả thêm file PO.

## V4. Tìm đúng dòng inbound — khóa **PO + article + màu**

Không thể chỉ dùng article: PO `TEC0002400` có **hai dòng cùng article `DA-DPS0001`**, chỉ khác
màu (SILVER SATIN / SANDSHELL). Thứ tự:

1. **Lọc theo PO** đã chuẩn hóa.
2. **Lọc theo article**: `AZ(article)` phải nằm trong `Material Description`, `Specification`
   hoặc `Supplier Ref`. Article lấy từ đầu phần mô tả, cắt ở `" - "` (khoảng trắng hai bên),
   ở 2+ khoảng trắng, hoặc ở `-FAB`:
   `HY-N403420FDY-FAB  NY/SP` → `HY-N403420FDY` · `SROPSJ11-OGC160 - FAB ORGA` → `SROPSJ11-OGC160`.
   Không dòng nào khớp article → dùng cả PO làm tập ứng viên.
3. **Chấm điểm màu** (`colorScore` × 3), cộng thêm:
   - `Lapdip Color` của inbound có từ đầu tiên (≥6 ký tự) xuất hiện trong mô tả/mã màu hóa đơn: **+2**
   - đơn giá trùng: **+1**
4. Lấy **các dòng điểm cao nhất**. Nhiều hơn một dòng → `CẦN KIỂM TAY`, **không tự điền**.

`colorScore` cho điểm theo thang từ chắc chắn đến mờ nhạt (v10) — so trên chuỗi đã bỏ
mọi ký tự không phải chữ/số:

| Điểm | Tình huống | Ví dụ thật |
|---|---|---|
| 10 | giống hệt | `UGW MARINE matching color w F25-LCST-004` ↔ y hệt |
| 9 | một bên là **tiền tố** của bên kia | `ROSEWATER` ↔ `ROSEWATER - WNLI0857429-E` · `SILVER SATIN` ↔ `SILVER SATIN-0500552` |
| 8 | một bên **nằm trong** bên kia | `BLACK 19-4201TSX` ⊂ `…SOLID BLACK 19-4201 TSX_NLK107341 SHADE C` |
| 7 | tên màu inbound có trong **mô tả** của chứng từ | |
| 6 | trùng **lõi màu** (đã bỏ phần "matching color w F25-LCST-004", ngoặc, "Non organic"…) | `UGW MARINE matching color w F25-LCST-004` ↔ `UGW MARINE matching w F25-LCST-005` |
| 5 | lõi màu chứa nhau | |
| 4 | **mọi từ** của lõi màu inbound có trong chứng từ (kể cả đảo thứ tự) | `POWDERED SUGAR 11-4002TSX` ↔ `…SOLID11-4002 TSX … Powdered Sugar` |
| 3 | quá nửa số từ của lõi màu | `332 CALM MINT 薄荷绿` ↔ `CALM MINT-S4805-A` |
| 0 | còn lại | `SILVER SATIN` ↔ `SANDSHELL-0802609` |

Điểm cao **thắng tuyệt đối**: nhờ vậy `UGW MARINE matching color w F25-LCST-004` (10 điểm)
tách được khỏi `UGW MARINE matching w F25-LCST-005` (6 điểm) — hai màu khác nhau của cùng
một loại vải bo.

Mã màu (`COLOUR CODE` trên hóa đơn) **không** dùng làm điều kiện, chỉ làm điểm phụ — vì dữ liệu
thật cho thấy nó khi trùng khi không (`TLP24050384-H` vs `TLP24050383-OK F`).

## V5. Số lượng — lấy theo đơn vị của **file inbound**

File inbound lấy từ PO nên `Base Unit of Measure` là đơn vị chuẩn. Hóa đơn phải cho đúng con số
cùng đơn vị đó:

- inbound **KG** → lấy số kí (BLAO: `483.7`, không lấy `1638.6` mét)
- inbound **M** → lấy số mét (Techwork)
- inbound **YD** → lấy số yard (Hengyu)

Không có con số đúng đơn vị → `SAI ĐƠN VỊ`, không so bừa.

Công cụ **ghi cột `Invoice Quantity`** theo hóa đơn, vì SAP xuất file inbound với cột này bằng
số **còn lại của PO** chứ không phải số đã giao. Ghi chú luôn nói rõ: *"Đã ghi Invoice Quantity
= 889.6 (file có sẵn 915)"*.

## V6. Dung sai — chỗ khác trimming nhiều nhất

```
còn được nhận = Over Tolerance Qty − Delivered Qty
```

- `SL hóa đơn ≤ còn được nhận` → đạt
- vượt `Quantity` nhưng còn trong dung sai → `KHỚP (trong dung sai)` 🟡 **không phải lỗi**
- vượt dung sai → `VƯỢT DUNG SAI` 🔴 **SAP sẽ báo lỗi khi import**
- ít hơn `Quantity` → `KHỚP (giao thiếu)` 🟡 giao từng đợt, bình thường

Ví dụ thật: PO `TEC0001500` có `Quantity` 242, `Over Tolerance` 266.933, đã giao 5 → còn nhận
261.933; hóa đơn 261.2 → **đạt, sát mép**. Còn PO `HYU0001800` có `Over Tolerance` = 300 = `Quantity`
(0% dung sai) nhưng hóa đơn 302 → **VƯỢT DUNG SAI**, dù hợp đồng bán ghi ±5%.

## V7. Đối chiếu lô / cây vải (thay cho size)

Gộp các nhóm packing list có **cùng PO + màu khớp ≥ 2 điểm**, so tổng với số lượng hóa đơn.
Lệch thì liệt kê từng lô. Báo cáo có thêm sheet **`CHI TIET LO`**: hóa đơn · PO · item · lô ·
số cây · số lượng lô · đơn vị · SL hóa đơn · kết quả.

Một màu có thể nằm trong **nhiều lô** (BLAO: TEMPEST = `Y92296` 422.7 kg + `Y92387` 61 kg =
483.7 kg đúng bằng hóa đơn), và số cây có thể đánh lại từ 1 trong mỗi lô (Hengyu) — đều xử lý được.

## V8. Đối chiếu giá trị

- Đơn giá hóa đơn vs `Gross Price` — lệch là 🔴 `LỆCH GIÁ TRỊ`
- Thành tiền: Techwork so với `SL × Gross Price + Surcharge Item`; BLAO và Hengyu so với
  `SL × Gross Price` (phụ phí là dòng riêng, không nằm trong thành tiền từng dòng)
- Sai số cho phép: 0.02 với USD, 1 đồng với VND

## V9. Thứ tự kết luận trạng thái (vải)

1. Chưa có file inbound → `CHƯA CÓ INBOUND`
2. PO không có trong inbound → `LỖI`
3. Không dòng nào khớp article + màu → `THIẾU DÒNG`
4. Nhiều dòng cùng điểm cao nhất → `CẦN KIỂM TAY`
5. Đơn vị không so được → `SAI ĐƠN VỊ`
6. Đơn giá / thành tiền lệch → `LỆCH GIÁ TRỊ` 🔴
7. Vượt dung sai → `VƯỢT DUNG SAI` 🔴
8. Packing list ≠ hóa đơn → `LỆCH PKL` 🟡
9. Vượt `Quantity` nhưng trong dung sai → `KHỚP (trong dung sai)`
10. Ít hơn `Quantity` → `KHỚP (giao thiếu)`
11. Còn lại → `KHỚP`

Ba trạng thái bắt đầu bằng `KHỚP` đều được tính là **đạt**, không vào cột "dòng cần xem lại".

## V10. Những gì công cụ **không** làm với vải

- Không tự chọn khi có nhiều dòng inbound cùng điểm → `CẦN KIỂM TAY`.
- Không đổi đơn vị (không tự quy mét ↔ kí ↔ yard).
- Không phân bổ phụ phí dòng riêng của BLAO vào từng dòng hàng — phụ phí đã nằm đúng dòng trong
  file inbound do SAP xuất ra.
- Không sửa gì trong file inbound ngoài `Invoice Number`, `Invoice date`, `Invoice Quantity`.


---
---

# PHẦN III — J&H YUBO và INKAVA (từ v10, 30.09.2026)

## Y. J&H Yubo — packing list nội bộ kèm đơn giá

Chứng từ là **một sheet `PKL`**: mỗi dòng là **một lô (Lot)** của một mã vải + màu, kèm đơn
giá và thành tiền. Không có hóa đơn riêng.

**Y1. Nhận diện:** trong 30 dòng đầu có cả `SCAVI CODE` và `PACKING LIST`.

**Y2. Đọc:**

| Trường | Nguồn |
|---|---|
| Số hóa đơn | ô có nhãn `HD:` (thường còn trống) |
| Ngày hóa đơn | ô có nhãn `Ngày xuất HD:` |
| Mã vải | cột `SCAVI CODE` — trùng với **đầu** cột `Material` của inbound (`MFKNSJSL0520` ⊂ `MFKNSJSL0520089`) |
| Màu | cột `Color` |
| Lô | cột `Lot`, số cây ở `Số lượng (ROLL)` |
| Số lượng | cột **`Cân nặng (KG)`** (đơn vị KG) |
| Đơn giá / thành tiền | `Đơn giá (Price)` · `AMOUNT` (VND, `SL × đơn giá`) |
| Tổng | dòng `Tổng cộng:` — cột AMOUNT là tiền hàng, cột kế bên là tiền có VAT |

**Gộp dòng:** nhiều lô cùng (SCAVI CODE + màu) được **cộng lại thành một dòng đối chiếu**,
vì trong file inbound chúng là cùng một dòng PO. Danh sách lô vẫn giữ để in ra sheet `CHI TIET LO`.

**Cột "comment":** cột nào có chữ `comment` trong tiêu đề (vd `Mỹ Ngọc comment 29Sep`) **không**
được dùng làm số lượng; nếu khác cột `Cân nặng` thì chỉ **ghi chú nhắc** trong báo cáo.
(Dữ liệu thật: lô `YB260808028A` ghi 104,7 ở cột Cân nặng nhưng 76,7 ở cột comment.)

**Y3. Ô PO ghi nhiều mã** — `J&H0009000 / J&H0008500 / J&H0009100 / J&H0009400 / J&H0009700`.
Công cụ tách ra tất cả và dò trong cả 5 PO.

**Y4. Chia PO theo FIFO — luật quan trọng nhất của Yubo.**
Cùng một mã Material (vd `MFKNSJSL0520089`) có thể nằm trong 2–4 PO. Dữ liệu **không** nói được
PO nào, nên:

> Sắp các PO ứng viên theo số PO tăng dần, lấp đầy `Over Tolerance Qty − Delivered Qty` của PO
> nhỏ nhất trước, thừa thì tràn sang PO kế tiếp. Trạng thái `KHỚP (chia nhiều PO)`, ghi rõ chia
> bao nhiêu cho PO nào, và có sheet **`PHAN BO PO`** để kiểm/sửa.
>
> **Từ v10.2:** dòng PO nào đã **giao đủ** (`Delivered ≥ Quantity`) thì xếp xuống cuối — chỉ dùng
> phần dung sai còn lại của nó khi các dòng còn mở không đủ chỗ. Trước đó một dòng 44,8 kg bị
> cắt thành 5,848 kg (ghi vào PO `J&H0008500` đã giao 354,4/343,095) + 38,952 kg, nay vào thẳng
> một dòng của `J&H0009000`.

Phần đã chia của một dòng inbound được **trừ vào chỗ còn lại** của dòng đó, nên hai dòng hóa đơn
khác nhau không cùng giành một chỗ.

**Y5. Số hóa đơn — ba tình huống** (v10.1):

| Tình huống | Công cụ làm gì |
|---|---|
| Ô `HD:` trống | vẫn xuất, để trống cột `Invoice Number`, nhãn vàng *"chưa có số HĐ"* |
| Ô `HD:` chỉ có số (vd `1728`) | điền luôn `1728`, nhãn vàng *"thiếu ký hiệu HĐ"* — vì SAP cần cả ký hiệu |
| Thả kèm **file PDF hóa đơn GTGT** | ghép với file PKL và điền đủ `1C26TYY#00001728` + đúng ngày trên hóa đơn |

**Ghép hóa đơn PDF với file PKL** (sửa ở v10.2): các file PKL của cùng chủ hàng đều ghi **hết** số
PO giống nhau, nên chỉ dựa vào số PO là ghép sai. Thứ tự ưu tiên:

| Dấu hiệu | Điểm |
|---|---|
| số trong ô `HD:` trùng số hóa đơn trên PDF (`1728` ⊂ `00001728`) | **+5** |
| ngày (`Ngày xuất HD:`) trùng ngày hóa đơn | **+3** |
| mỗi số PO của PKL xuất hiện trong nội dung PDF | +1 |
| cùng thư mục | +0,5 |

**Phải có ít nhất một trong hai dấu hiệu đầu**, nếu không thì không ghép. Chấm điểm toàn cục rồi gán
từ cặp điểm cao nhất xuống, nên hai file PKL không giành nhau một hóa đơn. File PDF đã dùng cho
chứng từ vải thì **không** hiện thành một hóa đơn riêng nữa.

**Đối chiếu với hóa đơn GTGT:** so `Cộng tiền hàng` trên PDF với tổng thành tiền của các dòng PKL.
Khớp thì ghi "khớp" trong báo cáo; lệch quá dung sai thì **cảnh báo đỏ**.
(Dữ liệu thật: PKL `Sep 26.9_4 Bulk HD1728` có 10.944,1 KG / 1.890.286.944 đ — khớp đúng hóa đơn
`1C26TYY#00001728`.)

**Y6. Chọn bản xuất inbound.** Tên file SAP có dấu thời gian (`ZMME0032_20260926042733`). Thả nhiều
bản thì chọn bản có **ngày gần ngày hóa đơn nhất** (≤2 ngày +2,5 điểm · ≤7 ngày +1), sau tỉ lệ trùng PO.

**Y7. Phân xử màu "matching".** Một mã vải có thể có nhiều màu chỉ khác nhau ở phần đuôi:
`GH0 ARGENT CHINE HEATHER Bros B12B` · `… matching w F25-LCST-004` · `… matching w F25-LCST-005`.
Lõi màu của cả ba giống nhau, nên khi lõi trùng thì so tiếp **phần sau chữ "matching"**:
khớp **+1**, khác nhau **−2**, một bên có một bên không **−1**. Nhờ vậy dòng đúng thắng rõ rệt
(6 điểm so với 4 và 3) thay vì cả ba bằng 5 điểm rồi phải `CẦN KIỂM TAY`.

**Y8. Nhóm lô ghép bằng khóa chính xác.** Vì nhóm packing list của Yubo do chính công cụ sinh ra từ
các dòng hàng, việc ghép dùng **khóa (SCAVI CODE + màu)** chứ không dò mờ theo màu — nếu dò mờ thì
`GH0 … Bros B12B` và `GH0 … matching w F25-LCST-004` sẽ gộp lô của nhau và báo `LỆCH PKL` oan.

## I. Inkava — hóa đơn PDF + packing list Excel theo PO

Đây là chủ hàng **trimming** (nhãn), không phải vải: hóa đơn GTGT dạng PDF như ITL, nhưng packing
list là file Excel riêng **cho từng PO** và có sẵn mã SAP.

**I1. Hóa đơn — dạng mô tả khác.** Trên hóa đơn ITL mã hàng nằm giữa `/` và `//`. Inkava thì:

```
PODUY0095800- Label L56xW20MM-          ← PO ở dòng TRƯỚC
1 PCS 16.988,00 100 1.698.800           ← dòng số lượng
ALABPRWV0228                            ← mã hàng ở dòng SAU
```

> Nhận ra bằng: bảng hàng (từ dòng "Tên hàng hóa" đến dòng "Cộng tiền hàng") **không có** dấu
> `//` nhưng **có** dòng khớp `PO<chữ><số>`. Khi đó PO lấy từ ngữ cảnh trước dòng số lượng, mã
> hàng lấy ở dòng sau (token dạng `[A-Z]{3,}\d{3,}` cuối cùng, bỏ token chứa chính mã PO).

Cũng bổ sung hai cách đọc đầu hóa đơn: `Số: 00000614` (không có chữ "No.") và
`Ngày29tháng09năm2026` (không có chữ "date").

**I2. Packing list Excel theo PO.** Nhận ra bằng: trong 10 dòng đầu có `PO No:` **và** cột
`Material Code`. Đọc: `Material Code | Description | Supp. Ref. | Order No | Reference | [Story] |
Size | Spec. | Quantity`. Một hóa đơn nhiều PO → thả nhiều file, công cụ gắn theo số PO.

**I3. Điền Invoice Quantity cho TỪNG dòng inbound.** Vì packing list có đủ
`Material Code + Size + Spec`, công cụ ghép đúng từng dòng inbound và điền số lượng — không phải
điền tay 30–90 dòng. Khóa ghép: `Material + Size + Specification`, dự phòng `Material + Size`.
Dòng inbound không có trong packing list thì không được điền và bị loại khỏi file import
(nếu tick "chỉ giữ các dòng thuộc hóa đơn").

**I4. Packing list gộp nhiều Order No.** Dữ liệu thật: `DUY0095800 revise 2` có 83 dòng, tổng
29.742 trong khi hóa đơn ghi 16.988 (và dòng `Total:` trong file cũng ghi 16.988 — tổng cũ chưa
cập nhật).

> Luật: nếu tổng packing list ≠ số lượng hóa đơn mà có **đúng một** nhóm `Order No` cộng lại bằng
> số lượng hóa đơn → lấy nhóm đó và nói rõ trong ghi chú (ở đây là `112568 P27` = 16.988).
> Không nhóm nào khớp → báo `LỆCH SL` kèm bảng tổng theo từng Order No.

**I5. PO lấy từ file inbound.** Nếu mã PO không tra được qua `PO SCAF-SCAX.xlsx` nhưng chính cột
A của file inbound đã có mã đó → dùng luôn (loại PO ghi là `inbound`).

## Thay đổi khác của v10

- Dò mã hàng nay tìm trong cả cột **`Material`** của inbound, không chỉ `Material Description` +
  `Specification` (cần cho Inkava và Yubo).
- Số PO chuẩn hóa nhận cả ký tự `&` trong tiền tố (`J&H0009000`).
- Nhóm packing list phải khớp **cả mã article**, không chỉ màu — vì một màu (`UGW MARINE`) dùng
  cho nhiều loại vải.
- Trạng thái mới: `KHỚP (chia nhiều PO)`. `CẦN KIỂM TAY` nay chỉ dùng khi các dòng ứng viên
  **khác mã Material** (mơ hồ thật), không dùng cho trường hợp cùng mã ở nhiều PO.
- Fujian Techwork bán cho cả SCAVI và B'LAO SPORT — cùng một mẫu chứng từ, không cần thêm gì.

---

## Thay đổi của v10.2 (05.10.2026)

- **Bỏ dòng rác của file SAP.** File `ZMME0032_*.xlsx` đôi khi còn một dòng chỉ có một ô lẻ
  (vd dòng 424 chỉ có `Over Tolerance Qty = 100,001`, không có số PO). Trước đây dòng này
  **sống sót vào file import** vì bộ đọc bỏ qua các dòng không có PO nên nó không nằm trong danh
  sách cần xóa. Nay khi tick *"chỉ giữ các dòng thuộc hóa đơn"*, mọi dòng không được ghi số lượng
  đều bị xóa, kể cả dòng rác.
- **Đối chiếu tổng số lượng độc lập.** Trước đây công cụ ghi số lượng của chứng từ vào inbound rồi
  so chính con số đó → luôn "khớp". Nay so dòng **`Tổng cộng`** trên chứng từ với **tổng
  `Invoice Quantity` thực sự đã ghi vào file inbound**; lệch là cảnh báo đỏ
  *"LỆCH TỔNG SỐ LƯỢNG"*, và sheet `TONG HOP` có ba cột mới: `Tổng SL chứng từ` ·
  `Tổng SL ghi inbound` · `Lệch SL`.
- Chia FIFO bỏ qua dòng PO đã giao đủ (xem Y4).
- Ghép hóa đơn PDF ↔ PKL theo số `HD:` và ngày (xem Y5).

---

## Thay đổi của v10.3 (07.10.2026) — hóa đơn ITL đổi cách ghi PO

Từ 10.2026 hóa đơn của ITL (ký hiệu `1C26TVN`, bản Viettel) **ghi thẳng PO ScaF** sau dấu `//`
thay cho mã ScaX kiểu cũ:

```
cũ :  … / LB 5873 // TRIMMINGVN-0725      → phải tra bảng PO SCAF-SCAX
mới:  … / VCH101F // TRI0012500           → đã là PO SAP
      … / LB 6620 // TRI0010400
```

Vì luật dò PO cũ chỉ nhận dạng `TRIMMINGVN-xxxx` (regex `[A-Z]{4,}-\d{3,5}`) nên với hóa đơn mới
mã PO **không đọc được**, mọi dòng ra `THIẾU FILE PO` và **"Xuất danh sách PO" ra 0 PO**.

**Sửa:** thêm một bước nữa vào `resolvePo`, chạy sau hai bước tra bảng:

> Nếu trong mô tả có token đúng dạng **2–6 chữ cái + đúng 7 chữ số** (`TRI0012500`, `TEC0002400`,
> `DUY0094000`) thì dùng luôn làm PO SAP, **không cần file PO SCAF-SCAX**. Loại PO ghi là
> `ScaF (sẵn trên hóa đơn)`.

Hai chi tiết phải cẩn thận:

- `PODUY0095800` (Inkava dán liền chữ PO) cũng đúng dạng 5 chữ + 7 số → phải **bỏ hai chữ `PO`
  đầu** khi phần còn lại vẫn đúng dạng, nếu không sẽ tra ra một mã PO không tồn tại.
- Mã PO tra được nhưng **không có trong cột A của file inbound** → tự dò lại theo cột A
  (trước đây chỉ dò khi hoàn toàn không tra được).

Ngoài ra: đọc thêm mã PO ngay sau `//` vào trường `po` của dòng hàng để báo cáo hiện đúng mã, và
**một `Despatch Note` chia thành nhiều hóa đơn thì các hóa đơn đó dùng chung một packing list**
(bộ `DBVN230276` được chia cho hai hóa đơn `2564` và `2565`; trước đây hóa đơn thứ hai bị mất
packing list vì mỗi file chỉ được gán cho một hóa đơn).

---

## Thay đổi của v10.4 (07.10.2026) — mất nhóm size đầu tiên của packing list

**Triệu chứng:** hóa đơn `1C26TVN#00002566`, item `LB 6742` — hóa đơn 500, inbound 500, nhưng
**SL packing chỉ 400** → báo `LỆCH PKL`; bảng size cho thấy `XS: inbound 100 / packing 0`.

**Nguyên nhân.** Packing list `DBVN230292` không ghi mã invoice trên từng dòng size (size ghi trần
`XS - 100 PCS`), nên công cụ dùng luật dự phòng. Sau khi nối các dòng pdf lại, đoạn text là:

```
… 5100747450 LB 6742 XS - 100 PCS S - 100 PCS M - 100 PCS …
```

Luật dự phòng bắt được cụm `LB 6742 XS - 100 PCS`, thấy phần "size" chứa mã nhãn `LB 6742` nên
**loại cả cụm** — và vì bộ dò tiếp tục từ sau cụm đã loại, **100 cái size XS bị mất luôn**. Bốn
size còn lại (`S`, `M`, `L`, `XL`) đứng sau một size khác nên không dính mã nhãn, vẫn đọc được →
ra đúng 400.

**Sửa:** thay vì loại cả cụm, **cắt bỏ phần tới hết mã nhãn** rồi lấy phần còn lại làm size
(`LB 6742 XS` → `XS`). Chỉ bỏ qua khi sau khi cắt không còn gì.

**Ảnh hưởng tốt ngoài dự kiến:** bộ ATUS cũ cũng bớt một cảnh báo sai —
`1C26TVN#00002095 / LB 6620 / TRI0005700` trước đây báo `LỆCH PKL` (hóa đơn 853 / packing 794);
nay packing đọc đủ **853**, đúng bằng hóa đơn. Chênh 59 chính là nhóm size đầu tiên bị mất.

---

# PHẦN IV — BỘ ĐỌC CHUNG cho chủ hàng chưa có mẫu riêng (từ v11, 07.10.2026)

Từ v11, mọi file Excel/PDF **không** thuộc 6 mẫu riêng (ITL, Inkava, Techwork, BLAO, HYU, Yubo)
đều được thử bằng **bộ đọc chung**. Bộ đọc này không biết trước bố cục của chủ hàng; nó tìm bảng
hàng bằng từ khoá, rồi dùng **chính dữ liệu của file inbound** (mã Material, PO, Supplier Ref,
màu, lapdip, spec, size) để nhận ra từng dòng chứng từ. Kết quả vẫn đi qua đúng bộ đối chiếu
của vải (PHẦN II) nên trạng thái, dung sai, báo cáo, file INB giống hệt.

Đã kiểm thử với 31 bộ chứng từ (danh sách 07.10.2026): Capital Tricot, Carvico, Suzhou Celeb,
Cheung Hing, Chuangjie, Derun, DJIC, Dongguan Uwork, Freetex VN, Fujian Baikai, Fujian Honggang,
Hing Yip, Hoa Nghiêm Vina, Best Pacific, China Junye, Luen Hing, Pioneer Elastic, PT Winnersumbiri,
S&M International, Seamless Solutions, Stretchline, SunPo, Yibei, Brugnoli, AIM High, Chain Guan
(KNE) — tất cả đọc được và điền đúng theo số người đã điền tay trong inbound (trừ các dòng người
điền từ nguồn khác). PDF dạng ảnh (scan) từ v11.1 được đọc bằng **OCR** (G9): Paddies, Prestige,
SunPo CI, hoá đơn Brugnoli đọc được; Vanessa không (con dấu đè lên bảng). Excel 97-2003 `.xls`
(Vinity) từ v11.2 được tự chuyển sang `.xlsx` trong trình duyệt (G10). **Không đọc được**: Freetex PKL
(.doc), Showwin (thư mục rỗng).

## G1. Nhận diện vai trò file

| File | Cách nhận |
|---|---|
| Inbound SAP | như cũ (`purchasing document`…) — luôn thử trước |
| Mẫu riêng | như cũ (`fabProfile`, PKL Excel theo PO, hoá đơn ITL/Inkava) |
| **Excel chung** | sheet có dòng tiêu đề bảng (≥ 2 cột nhận ra, trong đó có cột số lượng). Sheet có cột **giá/thành tiền** → *hoá đơn*; không có giá (hoặc tên sheet PACK/PKL/PL/装箱) → *packing list* |
| **PDF chung** | trang có tiêu đề `INVOICE`/`FATTURA`/`COMMERCIAL INVOICE` → hoá đơn; `PACKING LIST`/`DELIVERY NOTE`/`PHIẾU ĐÓNG HÀNG` → packing list; `PROFORMA INVOICE` → **bỏ qua**; `HÓA ĐƠN GIÁ TRỊ GIA TĂNG` → hoá đơn GTGT (xem G4) |
| PDF dạng ảnh | tổng số ký tự < 40 → báo *"PDF dạng ảnh (scan)"*, không xử lý |
| `.xls`, `.doc` | báo *"Lưu dưới dạng .xlsx"* / *"file Word — chưa hỗ trợ"* |

Hoá đơn GTGT dạng PDF vẫn được giữ làm **hoá đơn kèm** cho PKL Yubo (PHẦN III); chỉ khi không ghép
được với PKL vải nào và không theo mẫu ITL/Inkava thì nó tự trở thành chứng từ của bộ đọc chung.

## G2. Đọc bảng trong Excel

1. **Dòng tiêu đề** = dòng có điểm cao nhất trong 60 dòng đầu, điểm = số cột nhận ra theo từ khoá:
   `PO/P.O./Cust.PO#/Order No`, `ITEM/ARTICLE/ART.-NO/MODEL/STYLE/DESCRIPTION/MATERIAL/货号/品名`,
   `COLOUR/COLOR/CLR/SHADE/颜色`, `COLOUR CODE/LOT/BATCH/DYE LOT/缸号`, `SIZE/DIMENSION/尺码`,
   `QTY/QUANTITY/Q'TY/METERS/YDS/LENGTH/TTL/M/数量`, `UNIT/UOM`, `PRICE/U/P`, `AMOUNT/VALUE/金额`,
   `SURCHARGE/MCQ`, `ROLL`, `CTN/CARTON/BOX/BALE`. Cùng điểm → lấy dòng **sau** (tiêu đề con sát dữ liệu).
   Ô tiêu đề gộp ngang chỉ tính ở ô đầu. Cột `Item`/`No.` chỉ đánh số thứ tự không coi là mã hàng.
2. Cột số lượng: loại các cột `CTN`, `Qty/CTN`, `N.W`, `G.W`, `CBM`, `PRICE`…; nếu cột này gần như
   trống mà cột kề bên có số ở hầu hết dòng (mỗi dòng = 1 thùng) → dùng cột kề. Hai cột cùng tên
   `PO No` → chọn cột có mã SAP. Đơn vị lấy từ tiêu đề (`QTY/M`, `Q'ty(Y)`, `LENGTH(YARDS)`,
   `Quantity(PRS)`) hoặc ô ngay sau số lượng (`4379 | pc`).
3. **Dòng hàng** = dòng có số lượng > 0. Ô PO/mã/màu/size/lô trống **kéo xuống** từ dòng trên.
   Ô chữ ở cột không có tiêu đề vẫn được ghép vào mô tả (Cheung Hing ghi style `Panache-PAN3ST` ở cột trống).
   Bỏ qua: dòng `TOTAL`/`合计`, dòng surcharge/bank charge/freight, dòng chỉ có số lẻ loi (tổng phụ),
   dòng có `GROSS WEIGHT/ROLLS/CBM/BALES` mà không có PO/màu, dòng `Note:`; dừng ở `BANK/BENEFICIARY/SWIFT/AMOUNT IN WORDS`.
   Dòng cuộn chỉ gồm số (STT cuộn + mét + cân) được giữ nếu bảng có cột ROLL/LOT.
4. Số: `3.386,30` → 3386.3 · `1,673.0` → 1673 · `632,50` → 632.5 · `7.606` (file Việt) → 7606.
   Ô công thức cùng dòng (`=E54*F54`, `=F54`) được **tính lại** từ ô gốc. Giá/thành tiền ghi `USD` ở
   ô tiêu đề, số ở ô kế → lấy ô kế.
5. Số hoá đơn: nhãn `INVOICE NO / INV NO / CI NO / No. & date of invoice`, sau đó `SO No / Packing No / No.`;
   giá trị ở cùng ô (`NO:20260917`), ô bên phải, ô dưới, hoặc ô bên phải của ô dưới (`No. & date` → số + ngày).
   Loại giá trị giống nhãn (`INVOICE DATE`), giống mã PO SAP, có dấu phẩy, bắt đầu bằng `TEL/FAX`.
   Ngày: `7th,Agu,2026`, `SEP.01,2026`, `09/17/2026` (Mỹ → 17.09), `2026-09-22`, serial Excel.
6. Nhiều sheet: sheet hoá đơn có giá được ưu tiên (Fujian Honggang: Page 1/2 là packing, Page 3 là hoá đơn).

## G3. Đọc bảng trong PDF

1. Chữ được ghép thành dòng theo toạ độ (như PHẦN I); trang **xoay ngang** (Pioneer) được xoay lại trước.
2. **Tiêu đề bảng** tìm như Excel (ô là mẩu chữ); từ đó biết vị trí x của cột số lượng, size, màu, PO, giá.
   Không có tiêu đề đủ điểm → chấp nhận dòng ngắn chỉ có `METERS/QUANTITY/QTY`.
3. **Số lượng của một dòng** = số đứng ngay trước/sau một từ đơn vị (`309 YDS.`, `MT 632,50`,
   `1,673.0 meters`, `33 YDS`); ưu tiên số nằm trong cột số lượng; không có đơn vị thì lấy số dưới cột
   số lượng. Số có `0` đứng đầu (`03261` = mã màu) hay dài ≥ 8 chữ số (số cuộn) không phải số lượng.
   Đơn vị `KG` bị bỏ khi chứng từ tính bằng M/YD/PCS. Giá/thành tiền = hai số cuối sau số lượng nếu `giá × SL ≈ thành tiền`.
4. **Khối**: dòng chứa mã PO SAP mở khối mới (`PO.CAP0022700 …`, `775-775 1 PO: LHT0012200`,
   `DER0025700 D-PN28035S 419.70`). Chữ của khối (dòng trước, dòng sau khi khối chỉ có 1 dòng hàng,
   chữ phía trên bảng như `Your item: MFKSJQSL0127001`) là **ngữ cảnh** để dò mã hàng/màu/Material.
   Dòng hàng đứng ngay trên dòng PO mà chưa có PO (ô PO canh giữa dọc) được kéo sang khối mới.
   Dòng chỉ có số (size + số lượng + giá) thừa hưởng màu/mô tả của dòng hàng kề trong khối.
   Ô dưới cột PO ghi mã riêng (`PT.WINNER1-0667`) → **không** lấy PO của khối (tra ScaX hoặc báo LỖI).
   Khối + ngữ cảnh đi tiếp qua dòng `Total batch` và sang trang sau (bảng cuộn dài nhiều trang).
5. Bỏ qua dòng: `TOTAL…`, `N.W/G.W/CBM/WEIGHT`, `TARIFF/ORIGIN`, `DIFETTO/FAULT`, `DAYS/PAYMENT/REMARK/SHIPPING MARK`, bank.
6. Số hoá đơn: `INVOICE NO … 28493`, `Invoice: LHT-S-INV-2609/00159`, `Invoice Number 00268944`,
   `NR.DOCUMENTO` (giá trị ở dòng dưới, cùng cột x), `INVOICE NO.  DATE` / `CT-26-314T  AUGUST 19, 2026`
   (giá trị dưới nhãn). Không có nhãn → tên file (`MF2615049.PDF`, phải xuất hiện trong chữ file) hoặc
   tên thư mục (`invoice 100972 - …`), ghi rõ *"số HĐ lấy từ tên file/thư mục"*. Ngày: `24-Sep-26`,
   `22/07/26`, `AUGUST 19, 2026`, `dated Septenber 04th, 2026` (lỗi chính tả tháng vẫn hiểu).
7. Chỉ có packing list (hoá đơn là bản scan — Brugnoli): các dòng cuộn cùng PO + Material + màu + size
   được **gộp** thành một dòng; số hoá đơn lấy từ tên thư mục.

## G4. Hoá đơn GTGT Việt Nam không theo mẫu ITL (Hoa Nghiêm, Freetex…)

Dòng hàng = dòng bắt đầu bằng STT, có đơn vị (`M`, `CÁI`, `PCS`, `ĐÔI`, `KG`…) và ≥ 2 số sau đó:
số lượng, đơn giá, thành tiền (`1  M  7.606  1.632  12.412.992`). Mô tả nằm ở dòng kề trên/dưới
(tối đa 3 dòng, nhưng dòng cách 2–3 dòng mà sát dòng STT khác thì thuộc mặt hàng kia). Ký hiệu + số +
ngày + tổng tiền đọc như ITL. Mã Material SAP ghi trong mô tả (`AELATMFC0008019 - PO số FGV0000600`)
được dùng thẳng (G5.1). Mã hàng bị xuống dòng (`YN6044` / `8/3`) vẫn nhận ra vì so sau khi bỏ khoảng trắng.

## G5. Tìm đúng dòng inbound — thứ tự ưu tiên

1. **Mã Material SAP** có trên dòng chứng từ (Stretchline, Pioneer, Freetex, Brugnoli) → lấy thẳng
   dòng inbound có mã đó, kể cả khi PO ghi khác. Mã chỉ xuất hiện ở dòng mô tả chung (Chain Guan ghi 3
   mã trên một dòng) → không dùng.
2. **PO**: mã SAP trên chứng từ (`[A-Z]{3}\d{7}` hoặc `[A-Z]{2}\d{8}` như `CH10034200`; chấp nhận
   `PO.CAP0022700`, `PO: LHT0012200`, `P/O No: WIN0002500`). Chứng từ **không ghi PO** (Carvico, Hing Yip)
   hoặc ghi mã riêng (`FUJIANBAI1-2610`, `202608/A059/1`) mà inbound chỉ có **một** PO → dùng PO đó (ghi chú).
   Không ghi PO mà inbound nhiều PO → dò trên toàn bộ inbound. Dòng không ghi PO trong khi các dòng khác
   có → không tự gán.
3. **Mã hàng**: như vải (article ⊂ Material/Description/Spec/Supplier Ref) **hoặc ngược lại**: Supplier Ref /
   đầu Material Description của inbound (bỏ đuôi `Solid/Fab/Ela/Wire/Tape…`, ≥ 4 ký tự có số) nằm trong
   chữ chứng từ (`N.295/207/73`, `SW1235`, `RLI17042A01S3`, `232216`).
4. **Màu**: thang điểm của vải (PHẦN II) × 3, cộng +2 khi mã trong *Specification* (`PAN3ST`, `7283`)
   hoặc *Lapdip* (`LF610143B`, ≥ 5 ký tự) có trên chứng từ, +1 khi đơn giá bằng, +3 khi khớp Material.
5. **Size**: cột size, hoặc size trong mô tả (`36.9CM` ↔ `0369` ↔ `0309-30.9CM`, `75B`, `12DD`, `Bra cup M`).
   Có size mà PO không có size đó → **THIẾU DÒNG** (ghi các size PO có). Nhiều dòng cùng điểm → tách
   bằng số lượng như cũ.
6. Nhiều dòng cùng điểm, **khác Material** (các size của một mã): tổng còn nhận của chúng đúng bằng
   số chứng từ → điền từng dòng, trạng thái **KHỚP (chia theo size)**; không bằng → **CẦN KIỂM TAY** và
   **không điền gì** (v11 bỏ việc tự chia FIFO trong trường hợp này — FIFO chỉ còn cho cùng Material/nhiều PO của Yubo).
7. Chứng từ không ghi đơn vị → coi như đơn vị inbound (ghi chú). Chứng từ không có cột giá/thành tiền
   (packing list, Pioneer) → không kiểm giá trị, không báo lệch tổng tiền.

## G6. Packing list rời và hoá đơn ghi gộp

- Packing list rời (Excel/PDF, một hay nhiều file) gắn vào chứng từ chung theo điểm: cùng thư mục +6,
  tên file/số có số hoá đơn +3, mỗi PO trùng +1, trùng mã hàng/màu +2, cùng chủ hàng +1, chỉ có một
  cặp +1. Không có hoá đơn (scan) → packing list tự làm chứng từ.
- Ghép nhóm packing list với dòng chứng từ **thận trọng**: nhóm phải cộng đủ ≥ 3 điểm từ màu (+3/+2),
  mã hàng (+2, sai mã −3), size (+2, sai size −6), lô (bằng đúng +2, khác −4), PO (khác → loại).
  Nhiều dòng hoá đơn cùng trỏ về một nhóm (DJIC 6 + 106 ↔ 112) → so **tổng** của các dòng đó.
- Packing list đọc từ **PDF** chỉ để tham khảo: lệch thì ghi chú "chỉ để tham khảo", **không** hạ trạng thái
  (vì khó tách cuộn/lô/tổng phụ chính xác trong PDF).
- Hoá đơn ghi gộp theo mã hàng (SunPo: `IF-8307S 3105 pairs`) còn packing list ghi từng PO + size + màu:
  nếu tổng packing list của mã bằng tổng hoá đơn của mã đó → **tách** dòng hoá đơn theo packing list
  (giữ đơn giá nếu các dòng cùng giá), ghi chú *"tách theo packing list"*. Dòng hoá đơn đã có size/PO thì không tách.

## G7. Chọn file inbound cho chứng từ chung

Điểm như vải (PO trùng ×6, cùng thư mục +2, tên file +3, ngày gần +2.5/+1) cộng thêm +2 khi
*Partner Name* trong inbound bắt đầu bằng tên chủ hàng đọc từ chứng từ (Carvico không ghi PO), +0.5 khi
chỉ có một inbound và một chứng từ.

## G8. Những gì bộ đọc chung **không** làm

- OCR chỉ là phương án dự phòng (G9): số liệu từ bản scan luôn gắn nhãn OCR và phải đối chiếu bản gốc.
- Không đọc file Word `.doc` (Freetex PKL).
- Không đoán nhóm khi ô PO/mã canh giữa theo chiều dọc mà nhóm có > 3 dòng (Dongguan Uwork: dòng
  `12 40` nằm giữa hai mã) — dòng sai sẽ lộ ra dưới dạng VƯỢT DUNG SAI/THIẾU DÒNG để kiểm tay.
- Không tự chia size khi tổng không khớp (Cheung Hing nhãn `Evangeline-11434 58 DZ` cho 7 size).
- Packing list có thùng trộn size (Yibei thùng 24: 9 M + 170 L) → báo LỆCH PKL để kiểm tay.

## G9. PDF dạng ảnh (scan) — OCR (từ v11.1, 07.10.2026)

PDF không có lớp chữ (dưới 40 ký tự trích được) được coi là bản scan. Nếu ô **OCR** (cạnh ô kéo–thả)
đang bật, công cụ xử lý như sau; nếu tắt, file được báo "PDF dạng ảnh — đã tắt OCR" và bỏ qua.

1. **Tải bộ OCR khi cần.** Lần đầu gặp bản scan, trình duyệt tải tesseract.js (bộ nhận dạng ~4 MB + dữ liệu
   ngôn ngữ Anh + Việt ~4,5 MB) từ CDN jsDelivr; cần mạng ở lần này, các lần sau trình duyệt dùng
   bản đã lưu. Không tải được → file báo "OCR lỗi … cần mạng để tải bộ OCR". Bộ OCR **không** nhúng
   vào `index.html` để giữ file ~1 MB nén.
2. **Vẽ từng trang** bằng pdf.js ở độ phóng 3× (≈216 dpi) rồi nhận dạng; mỗi từ nhận được kèm toạ
   độ, được đổi về cùng hệ toạ độ như PDF có chữ, nên các luật G3 (tiêu đề cột, số lượng cạnh đơn vị,
   khối theo PO…) dùng lại nguyên vẹn.
3. **Sửa lỗi OCR hay gặp** trước khi dò: `O`→`0` và `I`/`l`→`1` trong các ô toàn số hoặc mã PO
   (`PDD00O34OO` → `PDD0003400`), bỏ ký tự `|` của đường kẻ bảng, tách từ dính như `141PAIRS` →
   `141` + `PAIRS`. Không sửa gì bên trong chữ (tên màu, mã hàng).
4. **Luôn đi qua bộ đọc chung** (PHẦN IV), không bao giờ đi vào mẫu packing list phụ liệu (mẫu đó cần
   chữ chính xác từng ký tự). Hoá đơn GTGT scan chỉ đọc được phần đầu (số, ngày) được ghép với
   chứng từ cùng thư mục theo G7/Y rồi lấy **ngày** từ đó (Paddies: ngày HĐ GTGT 11.09 thay cho ngày
   commercial invoice 10.09); số HĐ giữ theo ô "VAT Invoice#" trên chứng từ nếu có, vì OCR dễ đọc
   thiếu ký hiệu.
5. **Bản scan trùng với file Excel/PDF gốc** (SunPo gửi cả CI.xlsx lẫn CI scan): cùng thư mục và trùng
   số hoá đơn hoặc tổng số lượng lệch < 0,5 % → bản scan bị bỏ, chỉ giữ bản gốc. Nếu bản gốc chỉ là
   packing list (Brugnoli: PL có chữ + hoá đơn scan) thì giữ dòng hàng theo packing list, lấy số và
   ngày hoá đơn từ bản scan (`FATTURA N. 100972`), ghi chú "số hoá đơn lấy từ bản scan (OCR)".
6. **Không tự điền khi không chắc.** OCR không nhận ra bảng hàng (Vanessa: con dấu đè lên cột size/số
   lượng) → file báo "OCR đọc được chữ nhưng không nhận ra bảng hàng" kèm số HĐ nếu đọc được, không đưa
   vào đối chiếu. Dòng có số lượng nhưng không khớp đơn vị/tổng vẫn ra CẦN KIỂM TAY như thường.
7. **Gắn nhãn.** Hoá đơn có bất kỳ số liệu nào từ OCR được đánh dấu **OCR** trong báo cáo, kèm cảnh báo
   đỏ "Chứng từ là bản scan, đọc bằng OCR — kiểm tra kỹ số liệu với bản gốc trước khi import".

Kết quả kiểm thử: Paddies 5/5 dòng đúng số người điền (14 s), Prestige 1/1, Brugnoli 1/1 (số HĐ
100972, ngày 25.09), SunPo giống hệt khi chỉ có file Excel; Vanessa báo nhập tay.

## G10. Excel 97-2003 (`.xls`) — tự chuyển sang `.xlsx` (từ v11.2, 07.10.2026)

ExcelJS (thư viện đọc Excel của công cụ) chỉ đọc định dạng `.xlsx`. File `.xls` (định dạng nhị phân
BIFF của Excel 97-2003, Vinity/Vista Tex gửi `INVOICE.xls`) được xử lý như sau:

1. Thư viện SheetJS (bản core, nhúng sẵn trong `index.html`, chỉ nạp khi gặp file `.xls`) đọc file và
   ghi lại thành `.xlsx` **trong bộ nhớ trình duyệt** — không tạo file, không cần Excel, chạy offline.
   Giữ nguyên giá trị ô, ô gộp, ngày tháng; mất định dạng màu/khung (không cần cho việc đọc).
2. Bản `.xlsx` này đi tiếp đúng luồng của mọi file Excel khác (mẫu riêng → packing list theo PO →
   inbound → bộ đọc chung), nên hoá đơn, packing list hay cả file inbound dạng `.xls` đều dùng được.
3. Dòng thông báo dưới ô kéo–thả ghi `<tên file>: Excel 97-2003 — đã tự chuyển sang .xlsx để đọc`.
   File `.xls` hỏng hoặc không phải Excel thật (một số phần mềm xuất HTML rồi đặt đuôi `.xls` — SheetJS
   vẫn đọc được loại này) → báo "không chuyển được" kèm lý do, cần mở bằng Excel rồi *Lưu dưới dạng .xlsx*.

Cùng v11.2, bộ đọc chung Excel (G2) nhận thêm kiểu **dòng phụ ở DƯỚI**: bảng không có cột PO, màu và
PO ghi ở dòng ngay dưới dòng hàng (Vinity: `PADDING :FNJ08WHN1C236 /WIDHT 36"  YDS 236` rồi dòng
`COL.BEIGE /PO NO.VTC0000100`). Nhận ra khi dòng đầu tiên có mã PO SAP nằm **sau** dòng hàng đầu tiên;
khi đó dòng phụ được gán ngược cho dòng hàng phía trên (PO, màu — bỏ tiền tố `COL.` và đoạn `/PO NO.…`,
mã Material nếu có) và **không** kéo xuống dòng hàng kế tiếp. Các bảng kiểu cũ (PO/màu ghi ở dòng
trên hoặc có cột riêng) không đổi. Kết quả Vinity: 2/2 dòng khớp PO VTC0000100 BEIGE 236 / NAVY 127 YDS,
tổng 2.178 USD đúng hoá đơn.

## Thay đổi của v11.3 (07.10.2026) — PKL Yubo mẫu mới không có ô "HD:"

- PKL `PKL_SCAVI_6.10.xlsx` (10.2026) không còn ô `HD:` / `Ngày xuất HD:`; ngày ghi ở dòng
  `Ngày 6 tháng 10 Năm 2026` phía trên bảng. `parseDateCell` hiểu dạng "Ngày d tháng m năm yyyy"; reader
  Yubo lấy ngày ở các dòng trên tiêu đề bảng khi không có ô `HD:` → ghép được với hoá đơn GTGT PDF cùng
  ngày (luật Y, +3 cùng ngày, +1 mỗi PO trùng).
- Dự phòng ghép: khi chỉ còn đúng **một** hoá đơn GTGT PDF chưa ghép và đúng **một** chứng từ chưa có số
  HĐ, lại trùng PO trong nội dung PDF → ghép dù không trùng số/ngày. Trước đó bộ này tách thành hai hoá đơn
  (PKL "chưa có số HĐ" + PDF đọc bằng bộ đọc chung) — chính là lỗi "không nhận dạng được hoá đơn".
- Trạng thái mới cho PO đã giao một phần: đợt này vượt **phần còn lại** của PO nhưng cộng dồn vẫn trong
  dung sai → `KHỚP (trong dung sai)` với ghi chú "PO … đã giao … chỉ còn … — vượt … nhưng cộng dồn còn
  trong dung sai"; trước đây báo `KHỚP (giao thiếu)` với "còn lại" âm (khó hiểu). `VƯỢT DUNG SAI` không đổi.
- Kết quả bộ 1808: 4/4 dòng khớp PO J&H0009800, tổng 630.727.350 đ đúng hoá đơn, INB_1C26TYY-00001808.

## Thay đổi khác của v11

- Bỏ tự điền FIFO khi nhiều dòng inbound cùng điểm nhưng khác Material (trước đây ghi theo dung sai từng dòng — sai).
- File inbound có công thức dùng chung (Hing Yip cột Y): ô công thức được ghi thành giá trị trước khi xoá
  dòng, tránh lỗi "Shared Formula master must exist" khi tạo file INB.
- Tiêu đề inbound chấp nhận `Supplier Mat. No.` / `Supplier Material Number` (thay cho `Supplier Ref`)
  và `Input Quantity` (thay cho `Invoice Quantity`).
- Đơn vị: thêm `Y`, `PRS/PAIR/PAA`, `DZ/DOZ`, `SET`, `PIECE`.
- Mã PO SAP dạng 2 chữ + 8 số (`CH10034200`) được nhận.
- Ngày kiểu Mỹ `09/17/2026` → 17.09.2026; tên tháng viết sai (`Agu`, `Septenber`) vẫn hiểu.
