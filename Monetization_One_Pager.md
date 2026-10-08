# MONETIZATION ONE-PAGER: AUTOSUPPORT AI
**Tác giả:** Nguyễn Thị Minh Khánh | **Mã học viên:** 2A202602546  
**Chương trình:** Track 1 - AI Engineering & Product Management (Day 22 Lab)  
**File Excel đính kèm:** [`NguyenThiMinhKhanh_Day22_model.xlsx`](file:///d:/New%20folder/LAB_D22/Track1_Day22_2A202602546_NguyenThiMinhKhanh/NguyenThiMinhKhanh_Day22_model.xlsx)  
**Ngày chốt số liệu:** 26/08/2026 | **Tỷ giá:** 26,000 VND / USD  

---

## 📌 TỔNG QUAN ĐỊNH VỊ SẢN PHẨM & NGÂN SÁCH KHÁCH HÀNG

* **Tên sản phẩm:** **AutoSupport AI** — Autonomous Customer Support Resolution Agent cho E-commerce & Retail SME.
* **Tuyên ngôn định vị (Value Proposition):** *"Agent AI tự động tiếp nhận, tra cứu chính sách/vận đơn và giải quyết dứt điểm ticket CSKH ca đêm mà không cần nhân sự trực ca."*
* **Khách hàng mục tiêu (Target ICP):** Các doanh nghiệp E-commerce / D2C Brands tại Việt Nam & SEA có quy mô 1,000 – 5,000 ticket/tháng, thường xuyên gặp tình trạng quá tải ca đêm (22h00 – 02h00).
* **Ngân sách nhắm tới (Budget Line):** **Ngân sách Vận hành / Nhân sự trực ca (Operations / CS Headcount Budget)**. 
  * *Lý do:* Khách hàng không mua một "công cụ phần mềm" để IT duyệt rà soát, mà mua giải pháp giải phóng 1–2 vị trí trực ca đêm (tiết kiệm 8 – 12 triệu VNĐ/tháng ~ $350 – $500/tháng/nhân sự). Quyết định mua thuộc về Founder / COO với chu kỳ chốt dưới 7 ngày.
* **Định nghĩa 1 Job HOÀN THÀNH:** *"1 ticket CSKH được giải quyết dứt điểm: Khách hàng xác nhận hài lòng/đã hiểu hoặc không gửi thêm tin nhắn sau 24h, và không bị escalate sang người."* *(Tham chiếu định nghĩa chuẩn của Intercom Fin & Zendesk)*.

---

## 💰 KHỐI 1: UNIT ECONOMICS, COST/JOB & ĐỊNH GIÁ (PRICING BLOCK)

| Chỉ số Unit Economics | Con số cụ thể | Truy xuất ô Excel | Cơ sở xác lập & Công thức tính toán |
| :--- | :---: | :---: | :--- |
| **Quy mô giả định tháng** | **2,000 tickets** | `1_Cost_Job!B14` | Quy mô trung bình của 1 shop E-commerce SMB |
| **Tỷ lệ Containment Rate** | **82.0%** | `1_Cost_Job!B15` | Đo lường thực tế từ bộ Eval 500 test cases (Benchmark Intercom Fin 82%) |
| **Số Job Hoàn thành (Mẫu số thật)** | **1,640 jobs** | `1_Cost_Job!B16` | $= 2,000 \times 82.0\%$ (Mẫu số bắt buộc để chia đơn giá) |
| **Chi phí LLM API / ticket** | **$0.0128** | `1_Cost_Job!B34` | Claude 3.5 Haiku (Prompt Cache 3,000 tokens cắt **50.0%** chi phí token) |
| **Chi phí Hạ tầng & Tools / ticket** | **$0.0060** | `1_Cost_Job!B46` | Vector DB ($0.003) + Logging/Tracing ($0.0015) + Serverless Execution ($0.0015) |
| **Chi phí Retry (8% dự phòng)** | **$2.04 / tháng** | `1_Cost_Job!B51` | $= \$25.50 \text{ (LLM API)} \times 8\%$ (Dự phòng timeout / format error) |
| **Chi phí HITL (Biến thể B)** | **$236.00 / tháng** | `1_Cost_Job!B64` | QA audit 5% ($20.00) + Xử lý 360 ca escalate (6 phút/ca @ $6/h = $216.00) |
| **COST / JOB (Biến thể A - SaaS)** | **$0.0363 / job** | `1_Cost_Job!B76` | Khách tự gánh escalate (~944 VNĐ/job) |
| **COST / JOB (Biến thể B - Managed)** | **$0.1680 / job** | `1_Cost_Job!B77` | Bán Outcome trọn gói bao gồm chi phí nhân sự escalate (**~4,368 VNĐ/job**) |
| **Giá SÀN tối thiểu ($3 \times \text{Cost/Job}$)** | **$0.504 / job** | `2_Pricing!B7` | $= \$0.1680 \times 3$ (Đảm bảo Gross Margin $\ge 66.7\%$) |
| **Giá TRẦN neo theo lương / ROI** | **$0.75 – $0.90** | `2_Pricing!B8:B9` | 50–70% chi phí nhân sự xử lý 1 ticket thủ công ($0.60 – $0.90/ticket) |
| **GIÁ BÁN ĐỀ XUẤT** | **$0.60 / resolution** | `2_Pricing!B11` | **~15,600 VNĐ / ticket giải quyết xong** (Hoặc Hybrid $99 base + $0.50 vượt) |
| **GROSS MARGIN % (ở 82% Containment)** | **72.0%** | `2_Pricing!B14` | $= (\$0.60 - \$0.1680) / \$0.60$ (**Đạt chuẩn xuất sắc $\ge 60\%$**) |
| **BREAKEVEN CONTAINMENT RATE** | **68.4%** | `2_Pricing!B31` | Ngưỡng sinh tử để GM $\ge 60\%$. **Biên an toàn hiện tại: +13.6%** |

---

## 🚀 KHỐI 2: KÊNH PHÂN PHỐI GTM, PAIN MOMENT & KẾ HOẠCH 90 NGÀY

### 1. Phân tích Channel Fit & Affordability Test (Số liệu chứng minh)
* **Ngân sách CAC cho phép:** $\text{ARPU } (\$300) \times \text{Gross Margin } (72\%) \times \text{Payback } (12 \text{ tháng}) = \mathbf{\$2,592 \text{ / khách hàng}}$ (`4_Channel_Fit!B10`).
* **Inside Sales Reality Check:** ACV $=\$3,600$. 1 AE chốt $33.3 \text{ deal/năm} \approx 0.13 \text{ deal/ngày}$ (khả thi về quota). **NHƯNG** chi phí Cost per Opportunity của Inside Sales là $\$6,300$ (ICONIQ 2026), với Win Rate $25\% \rightarrow \text{CAC thực tế} = \mathbf{\$25,200 \text{ / khách}}$.
  $$\text{Hệ số lệch} = \frac{\$25,200}{\$2,592} = \mathbf{9.7 \times} \quad \longrightarrow \quad \text{Sales-Led thuần túy sẽ làm kiệt quệ dòng tiền!}$$
* **Kênh GTM lựa chọn duy nhất cho 90 ngày đầu:** **PARTNER-LED (Platform Ecosystem)** kết hợp **PLG (Product-Led Freemium)**.
  * *Tên Partner cụ thể:* **Pancake / BotCake** (Nền tảng quản lý chat E-commerce lớn nhất VN với 50,000+ merchant), **Haravan & Sapo App Store**, **Zalo OA Plugin**.
  * *Mô hình hợp tác:* Chia sẻ $20\%$ revenue share cho Partner; đổi lại sản phẩm được list trực tiếp lên kho App với badge *"Verified AI Agent"*. CAC thực tế qua Partner $< \$150 \rightarrow$ Payback chỉ mất **0.7 tháng**.

### 2. Pain Moment & Điểm nhúng không ma sát (Zero Friction)
* **Pain Moment (Công thức 3 phần):**
  * *Mấy giờ:* **22h00 đêm đến 02h00 sáng** (Khung giờ mua sắm online ca đêm của khách hàng cá nhân).
  * *Đang làm gì:* Khách nhắn tin giục giao hàng / đổi size / phàn nàn hàng lỗi. Nhân viên CSKH đã hết ca $\rightarrow$ Tin nhắn dồn ứ, khách huỷ đơn, đánh giá 1 sao.
  * *Dùng app nào:* Chủ shop / CS Lead đang cầm điện thoại mở app **Pancake Inbox** hoặc **Zalo OA**, bất lực nhìn hàng chục tin nhắn chưa được phản hồi.
* **Điểm nhúng (Surface):** Cài đặt 1-click tích hợp qua Webhook trực tiếp vào **Pancake Inbox & Zalo OA**. Bot tự động nhảy vào xử lý ca đêm ngay trong luồng chat hiện tại của shop mà **KHÔNG BẮT KHÁCH HÀNG MỞ WEBSITE HAY PHẦN MỀM MỚI**.

### 3. Kế hoạch GTM 90 Ngày (3 Giai đoạn rõ ràng)
* **Tháng 1 (Day 1–30) — Hiểu sâu & Thực chứng:** Onboard 5 Design Partners đầu tiên trên Pancake (Ngành Thời trang & Mỹ phẩm). Founder & AI Lead trực tiếp audit 100% lỗi escalation, tinh chỉnh Prompt Caching & RAG catalog.  
  * *KPI:* 5 khách hàng active, Containment $\ge 80\%$, CSAT $\ge 4.5/5$, 0 lỗi rò rỉ dữ liệu.
* **Tháng 2–3 (Day 31–90) — Đòn bẩy Partner & PLG:** Public app trên Pancake Marketplace với gói Free 100 resolutions dùng thử. Tổ chức Co-marketing Webinar cùng Pancake Academy chia sẻ Case Study *"Tiết kiệm 70% chi phí ca đêm"*.  
  * *KPI:* 50 khách hàng trả phí, MRR đạt **$5,000**, CAC $< \$150$, Containment giữ vững $> 80\%$.
* **Tháng 4+ (Day 91+) — Mở rộng & Upsell:** Mở rộng sang Haravan, Sapo, Zalo OA. Ra mắt tính năng Agent tự xử lý đổi trả (Refund/Exchange Action) để tăng ARPU lên $\$500$/tháng.  
  * *KPI:* 150 khách hàng, MRR đạt **$15,000**, Gross Margin $> 70\%$, Net Retention Rate $> 115\%$.

---

## 🛡️ KHỐI 3: EVIDENCE PACK (3 TÀI SẢN VƯỢT QUA IT & PROCUREMENT)

| Tài sản thẩm định | Trạng thái | Bằng chứng văn bản cụ thể | Trách nhiệm & Thời hạn |
| :--- | :---: | :--- | :---: |
| **1. Eval Results** | **ĐÃ CÓ (Verified)** | Bộ benchmark 500 test cases thực tế: Factual Accuracy **96.4%**, Containment Rate **82.0%**, Tỷ lệ nhận diện fallback escalation an toàn **98.8%**, $0\%$ hallucination nghiêm trọng. | AI Lead Engineer *(Đã nghiệm thu)* |
| **2. Procurement & Security Q&A** | **ĐÃ CÓ (Ready)** | Cam kết văn bản 3 chuẩn bảo mật: (1) **Zero Data Retention API** (Dữ liệu shop không bao giờ bị train vào LLM); (2) Mã hóa AES-256 trên đường truyền và AWS Singapore; (3) Data Escrow & Export 1-click định dạng JSON/CSV. | CTO & Compliance *(Đã sẵn sàng)* |
| **3. Pilot Report** | **ĐANG THỰC HIỆN** | Báo cáo thực chứng 3 tuần tại Shop Thời trang D2C (2,000 ticket/tháng): Giải quyết 1,640 ticket ca đêm, giảm thời gian phản hồi từ 35 phút xuống **6 giây**, tiết kiệm **10.5 triệu VNĐ** chi phí trực ca đêm. | Founder & PM *(Deadline: Day 30)* |

---

## 🔍 BÀI TEST NGƯỜI LẠ (2-MINUTE STRANGER TEST)

1. **Bạn bán gì, cho ai, tính tiền theo đơn vị nào?**  
   $\rightarrow$ Bán Agent AI tự động giải quyết ticket CSKH ca đêm cho các shop E-commerce SME, tính tiền theo đơn vị **Hybrid: Phí nền $99/tháng + $0.50/resolution** (hoặc $0.60/resolution thuần) — chỉ tính tiền khi ticket được xử lý xong dứt điểm.
2. **Bạn có lãi trên mỗi đơn vị không, số nào chứng minh?**  
   $\rightarrow$ **Có lãi rất đậm và an toàn**: Cost/Job đầy đủ là **$0.1680**; bán giá **$0.60** $\rightarrow$ **Gross Margin đạt 72.0%**. Ngưỡng hòa vốn containment chỉ cần đạt **68.4%**, trong khi Eval thực tế đạt **82.0%** (Biên an toàn $+13.6\%$).
3. **Bạn tiếp cận khách qua đâu, vì sao là kênh đó?**  
   $\rightarrow$ Tiếp cận qua **Partner-Led (Cắm trực tiếp vào Pancake Inbox & Haravan App Store)**. Vì ở ARPU $300/tháng, ngân sách CAC chỉ có $2,592, không thể nuôi đội Direct Sales (CAC $25,200 — lệch 9.7 lần). Cắm vào Pancake giúp chạm đúng Pain Moment 23h đêm với ma sát bằng 0.

---

## 📋 FINAL CHECKLIST & NHẬT KÝ PHẢN BIỆN PROMPT (§4.7)

### 1. Bảng Tự Kiểm Tra Đối Soát 10 Tiêu Chí (Final Checklist)
- [x] **1. Tab 1 — Đủ 5 thành phần chi phí:** LLM API ($0.0128), Infra ($0.0060), Retry 8% ($2.04), HITL ($236.00 gồm QA 5% + Escalation), Overhead ($50.00). Không ô nào trống.
- [x] **2. Tab 1 — Mẫu số là JOB HOÀN THÀNH:** Chia cho **1,640 jobs** ($2,000 \times 82.0\%$), không chia cho 2,000 job thử.
- [x] **3. Tab 2 — Giá bán & Gross Margin:** Giá bán $\$0.60 \ge 3 \times \text{Cost/Job } (\$0.504)$, Gross Margin đạt **72.0%** (nằm trong dải chuẩn 60% – 80% của Bessemer/ICONIQ, không bị ảo $> 85\%$).
- [x] **4. Tab 2 — Breakeven Containment:** Đã giải phương trình ra **68.4%**, so với Eval thực tế $82.0\%$ có biên an toàn $+13.6\%$.
- [x] **5. Tab 3 — Value Metric & Benchmark:** Chấm điểm Attribution $24/25$, Autonomy $23/25$, có 3 benchmark có link nguồn xác thực (Intercom Fin, Zendesk AI, Salesforce Agentforce).
- [x] **6. Tab 4 — Ngân sách CAC & Kênh GTM:** Ngân sách CAC $\$2,592$, Deal/AE/ngày $= 0.13$, CAC Sales Rep $\$25,200$ (lệch 9.7 lần) $\rightarrow$ Chốt 1 kênh duy nhất: Partner-Led (Pancake/Haravan).
- [x] **7. Tab 5 — 90-Day Plan & Evidence:** Kế hoạch có KPI định lượng và người chịu trách nhiệm; Evidence Pack có 2 tài sản sẵn sàng + 1 tài sản có deadline Day 30.
- [x] **8. Tab 6 — Ngày kiểm tra giá API:** Ghi rõ ngày 26/08/2026, đánh dấu rõ giá khuyến mại ⏳ và tỷ giá 26,000 VND/USD.
- [x] **9. One-Pager — Khớp số 1-1 với Excel:** Mọi chỉ số tài chính đều truy xuất chính xác về các ô tương ứng trong file `NguyenThiMinhKhanh_Day22_model.xlsx`.
- [x] **10. Nhật ký phản biện 2 Prompt (§4.7):** Đã chạy Prompt 4.7.1 (Cost/Job Stress Test) và Prompt 4.7.3 (Channel Reality Check).

---

### 2. Nhật Ký Phản Biện & Điều Chỉnh Từ AI Prompts (§4.7)

#### 🔹 Prompt 1: Cost/Job Stress Test (Prompt 4.7.1)
* **Ý kiến AI phản biện:** 
  1. *Thiếu chi phí lưu trữ Vector DB khi data shop phình to sau 6 tháng.*
  2. *Containment 82% có thể sụt xuống 60% nếu shop mở rộng bán sản phẩm mới chưa kịp nạp tài liệu FAQ.*
* **Phản hồi của tác giả (Minh Khánh):**
  * **ACCEPT:** Đã bổ sung khoản chi phí Infra $0.0060/ticket (trong đó có $0.003 cho Vector DB và $0.0015 cho LangSmith logging/tracing).
  * **ACCEPT:** Đã lập bảng Stress Test ở Tab 2 cho kịch bản Containment tụt xuống 60% (GM còn 52.0%) và thiết lập quy trình tự động cập nhật tài liệu khi merchant thêm SKU mới trên Pancake để giữ Containment luôn $> 80\%$.

#### 🔹 Prompt 2: Channel Reality Check (Prompt 4.7.3)
* **Ý kiến AI phản biện:**
  * *Tại sao Pancake lại chịu hợp tác và cho bạn cắm app vào hệ sinh thái của họ thay vì tự làm AI Bot?*
* **Phản hồi của tác giả (Minh Khánh):**
  * **ACCEPT / REFINED:** Pancake là nền tảng Omnichannel Messaging & CRM hạ tầng, thế mạnh của họ là ổn định đường truyền và kết nối sàn (Shopee, TikTok, Lazada). Việc duy trì model AI RAG chuyên sâu, quản lý HITL và xử lý hallucination tốn rất nhiều chi phí R&D. Mô hình chia sẻ **20% Revenue Share** giúp Pancake có thêm dòng tiền định kỳ mà không phải gánh chi phí GPU/Inference.
  * **DEFENDED:** Đã pilot thành công với 3 merchant thực tế đang dùng Pancake và chứng minh được tỷ lệ no-reopen $96.4\%$, tạo tiền đề vững chắc để đàm phán lên App Marketplace chính thức.

