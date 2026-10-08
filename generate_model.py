import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_monetization_workbook():
    wb = openpyxl.Workbook()
    default_sheet = wb.active
    
    # ------------------ STYLES & COLOR PALETTE ------------------
    font_title = Font(name="Calibri", size=13, bold=True, color="1F4E78")
    font_section = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_header = Font(name="Calibri", size=10, bold=True, color="000000")
    font_bold = Font(name="Calibri", size=10, bold=True, color="000000")
    font_normal = Font(name="Calibri", size=10, color="000000")
    font_italic = Font(name="Calibri", size=9, italic=True, color="595959")
    font_note = Font(name="Calibri", size=9, color="333333")
    
    # Fills
    fill_section_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid") # Dark Blue
    fill_subsection = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")     # Medium Blue
    fill_table_header = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")   # Light Steel Blue
    fill_yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")         # User Input (Yellow)
    fill_gray = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")           # Auto Formula (Gray)
    fill_green = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")         # Pass / Target (Light Green)
    fill_red = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")           # Alert / Warning (Light Peach/Red)
    fill_summary = PatternFill(start_color="D6DCE5", end_color="D6DCE5", fill_type="solid")       # Summary Row
    
    # Borders
    thin_border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    # Alignments
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_center = Alignment(horizontal="center", vertical="center")
    align_wrap_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # =========================================================================
    # TAB 0: 0_README
    # =========================================================================
    ws0 = wb.create_sheet(title="0_README")
    ws0.views.sheetView[0].showGridLines = True
    
    ws0["A1"] = "LAB THỰC CHIẾN DAY 22: AI MONETIZATION, UNIT ECONOMICS & GTM STRATEGY"
    ws0["A1"].font = font_title
    ws0["A2"] = "Học viên: Nguyễn Thị Minh Khánh | Mã học viên: 2A202602546 | Repository: Track1_Day22_2A202602546_NguyenThiMinhKhanh"
    ws0["A2"].font = font_italic
    ws0["A3"] = "Ngày chốt giá API & Benchmark: 26/08/2026 | Tỷ giá quy đổi: 26,000 VND / USD"
    ws0["A3"].font = font_bold
    
    # Legend Table
    ws0["A5"] = "QUY ƯỚC MÀU TRONG BẢNG TÍNH"
    ws0["A5"].font = font_section; ws0["A5"].fill = fill_section_header; ws0.merge_cells("A5:D5")
    
    legends = [
        ("Màu sắc", "Tên quy ước", "Ý nghĩa / Hành động của người làm bài", "Ví dụ áp dụng"),
        ("🟡 MÀU VÀNG", "User Input", "Ô người làm bài BẮT BUỘC phải điền số liệu / thông tin giả định", "Giá API, số turn, containment rate"),
        ("⬜ MÀU XÁM", "Formula / Calculated", "Công thức tự động tính toán - TUYỆT ĐỐI KHÔNG ĐƯỢC GHI ĐÈ", "Tổng LLM cost, Cost/Job, CAC Budget"),
        ("🟩 MÀU XANH LÁ", "Target Achieved / Healthy", "Chỉ số đạt ngưỡng chuẩn an toàn tài chính (Gross Margin >= 60%, LTV:CAC >= 3)", "Gross Margin 72%, Payback 4.5 tháng"),
        ("🟥 MÀU ĐỎ / CAM", "Alert / Unhealthy", "Chỉ số rơi vào vùng nguy hiểm tài chính - Cần điều chỉnh mô hình", "Gross Margin < 50%, Deal/AE > 1 deal/ngày")
    ]
    
    for r_idx, row in enumerate(legends, start=6):
        for c_idx, val in enumerate(row, start=1):
            cell = ws0.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_border
            if r_idx == 6:
                cell.font = font_header; cell.fill = fill_table_header; cell.alignment = align_center
            else:
                cell.font = font_normal
                if c_idx == 1:
                    cell.alignment = align_center
                    if "VÀNG" in val: cell.fill = fill_yellow; cell.font = font_bold
                    elif "XÁM" in val: cell.fill = fill_gray; cell.font = font_bold
                    elif "XANH" in val: cell.fill = fill_green; cell.font = font_bold
                    elif "ĐỎ" in val: cell.fill = fill_red; cell.font = font_bold
                else:
                    cell.alignment = align_left

    # Tabs Map
    ws0["A13"] = "BẢN ĐỒ CÁC TAB LÀM VIỆC VÀ NỘI DUNG CHÍNH"
    ws0["A13"].font = font_section; ws0["A13"].fill = fill_section_header; ws0.merge_cells("A13:D13")
    
    tabs_data = [
        ("Tab", "Tên Tab", "Vai trò & Nội dung chi tiết", "Đầu ra cốt lõi (Key Deliverables)"),
        ("Tab 0", "0_README", "Hướng dẫn quy ước, ngày chốt giá, danh mục các tab", "Bản đồ cấu trúc làm bài"),
        ("Tab 1", "1_Cost_Job", "Tab nặng nhất: Tính Cost/Job đầy đủ 5 thành phần (LLM + Infra + HITL + Retry + Overhead)", "Cost/Job hoàn thành (Biến thể A & B)"),
        ("Tab 2", "2_Pricing", "Tính giá sàn (3x Cost/Job), neo giá trần, xác định Gross Margin và Stress Test Breakeven", "Giá đề xuất ($0.60), Gross Margin (72%), Breakeven (68.4%)"),
        ("Tab 3", "3_Value_Metric", "Đánh giá ma trận Attribution x Autonomy, benchmark thị trường, chốt đơn vị tính tiền", "Value Metric bảo vệ được (Hybrid/Outcome)"),
        ("Tab 4", "4_Channel_Fit", "Tính Ngân sách CAC, Deal/AE capacity, Channel Affordability Test, chọn 1 kênh GTM", "Kênh Partner-Led (Pancake/Haravan) + CAC Budget ($2,592)"),
        ("Tab 5", "5_90Day_Plan", "Xác định Pain Moment 3 phần, Điểm nhúng zero-friction, Kế hoạch 90 ngày & Evidence Pack", "Pain Moment, 90-Day Plan, 3 Tài sản Evidence Pack"),
        ("Tab 6", "6_Benchmarks", "Bảng giá tham chiếu API chính thức và Benchmark tài chính B2B SaaS/AI từ Bessemer, ICONIQ", "Nguồn dữ liệu đối soát chuẩn mực")
    ]
    
    for r_idx, row in enumerate(tabs_data, start=14):
        for c_idx, val in enumerate(row, start=1):
            cell = ws0.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_border
            if r_idx == 14:
                cell.font = font_header; cell.fill = fill_table_header; cell.alignment = align_center
            else:
                cell.font = font_normal
                if c_idx == 1: cell.alignment = align_center; cell.font = font_bold
                elif c_idx == 2: cell.font = font_bold
                else: cell.alignment = align_left

    ws0.column_dimensions["A"].width = 16
    ws0.column_dimensions["B"].width = 24
    ws0.column_dimensions["C"].width = 65
    ws0.column_dimensions["D"].width = 45

    # =========================================================================
    # TAB 1: 1_Cost_Job
    # =========================================================================
    ws1 = wb.create_sheet(title="1_Cost_Job")
    ws1.views.sheetView[0].showGridLines = True
    
    ws1["A1"] = "TAB 1: TÍNH TOÁN COST/JOB ĐẦY ĐỦ 5 THÀNH PHẦN (UNIT ECONOMICS MODEL)"
    ws1["A1"].font = font_title
    ws1["A2"] = "Sản phẩm: AutoSupport AI | Job: Tự động giải quyết dứt điểm ticket CSKH đa kênh (E-commerce / D2C SME)"
    ws1["A2"].font = font_italic
    
    # Section 1: Job Definition
    ws1["A4"] = "SECTION 1: ĐỊNH NGHĨA JOB & ĐÓNG GÓI SẢN PHẨM"
    ws1["A4"].font = font_section; ws1["A4"].fill = fill_section_header; ws1.merge_cells("A4:E4")
    
    s1_data = [
        ("Tên sản phẩm AI", "AutoSupport AI - Autonomous Customer Support Resolution Agent"),
        ("Câu mô tả sản phẩm (Bản A - Tool)", "Công cụ AI Copilot hỗ trợ nhân viên CSKH tra cứu và soạn câu trả lời nhanh."),
        ("Câu mô tả sản phẩm (Bản B - Job)", "Agent AI tự động tiếp nhận, tra cứu chính sách/vận đơn và giải quyết dứt điểm ticket CSKH ca đêm mà không cần nhân sự trực ca."),
        ("Phiên bản lựa chọn & Ngân sách", "Chọn Phiên bản B: Đánh trực tiếp vào Ngân sách Vận hành / Nhân sự trực ca (Ops Headcount Budget) - Ngân sách lớn và quyết nhanh."),
        ("Định nghĩa 1 Job HOÀN THÀNH", "1 ticket CSKH được giải quyết dứt điểm: Khách hàng xác nhận vấn đề đã xong hoặc không phản hồi lại sau 24h, và không bị escalate sang người."),
        ("Phân loại trách nhiệm HITL", "Biến thể B (Managed Outcome - AI gánh vác xử lý cam kết trọn gói, có tính chi phí nhân sự escalation vào COGS)")
    ]
    for idx, (label, val) in enumerate(s1_data, start=5):
        ws1.cell(row=idx, column=1, value=label).font = font_bold
        ws1.cell(row=idx, column=1).border = thin_border
        cell_v = ws1.cell(row=idx, column=2, value=val)
        cell_v.font = font_normal; cell_v.fill = fill_yellow; cell_v.border = thin_border
        ws1.merge_cells(start_row=idx, start_column=2, end_row=idx, end_column=5)

    # Section 2: Volume & Containment
    ws1["A12"] = "SECTION 2: KHỐI LƯỢNG VẬN HÀNH & CONTAINMENT RATE (THÁNG)"
    ws1["A12"].font = font_section; ws1["A12"].fill = fill_section_header; ws1.merge_cells("A12:E12")
    
    headers_5col = ["Khoản mục", "Giá trị", "Đơn vị", "Công thức / Nguồn cơ sở", "Ghi chú"]
    for c_idx, h in enumerate(headers_5col, start=1):
        c = ws1.cell(row=13, column=c_idx, value=h); c.font = font_header; c.fill = fill_table_header; c.border = thin_border; c.alignment = align_center
        
    s2_rows = [
        ("Tổng số ticket tiếp nhận (Jobs Attempted)", 2000, "tickets/tháng", "Giả định quy mô 1 shop E-commerce SMB", "Input biến số quy mô", fill_yellow, font_bold, "#,##0"),
        ("Tỷ lệ AI tự xử lý thành công (Containment Rate)", 0.82, "%", "Benchmark Intercom Fin 82% / Kết quả Eval", "Biến sinh tử của bài toán", fill_yellow, font_bold, "0.0%"),
        ("Số Job AI HOÀN THÀNH (Jobs Completed - Mẫu số thật)", "=B14*B15", "jobs/tháng", "= B14 * B15", "MẪU SỐ BẮT BUỘC ĐỂ CHIA COST/JOB", fill_gray, font_bold, "#,##0"),
        ("Số ticket chuyển sang người (Escalations)", "=B14*(1-B15)", "tickets/tháng", "= B14 * (1 - B15)", "Cần người giải quyết", fill_gray, font_normal, "#,##0")
    ]
    for idx, (label, val, unit, form, note, cfill, cfont, nfmt) in enumerate(s2_rows, start=14):
        ws1.cell(row=idx, column=1, value=label).font = font_bold if "HOÀN THÀNH" in label else font_normal
        ws1.cell(row=idx, column=1).border = thin_border
        c = ws1.cell(row=idx, column=2, value=val); c.font = cfont; c.fill = cfill; c.border = thin_border; c.number_format = nfmt; c.alignment = align_right
        ws1.cell(row=idx, column=3, value=unit).font = font_normal; ws1.cell(row=idx, column=3).border = thin_border
        ws1.cell(row=idx, column=4, value=form).font = font_italic; ws1.cell(row=idx, column=4).border = thin_border
        ws1.cell(row=idx, column=5, value=note).font = font_note; ws1.cell(row=idx, column=5).border = thin_border

    # Section 3: LLM API Cost
    ws1["A19"] = "SECTION 3: CHI PHÍ LLM API (CÓ PROMPT CACHING & FRESH TOKENS)"
    ws1["A19"].font = font_section; ws1["A19"].fill = fill_section_header; ws1.merge_cells("A19:E19")
    
    for c_idx, h in enumerate(headers_5col, start=1):
        c = ws1.cell(row=20, column=c_idx, value=h); c.font = font_header; c.fill = fill_table_header; c.border = thin_border; c.alignment = align_center
        
    s3_rows = [
        ("Mô hình LLM chính sử dụng", "Anthropic Claude 3.5 Haiku (hoặc Gemini 3.7 Flash)", "text", "Anthropic Pricing 26/08/2026", "Model tối ưu speed/cost", fill_yellow, font_normal, "@"),
        ("Giá Fresh Input per 1M tokens", 1.00, "$/1M", "Anthropic List Pricing", "Giá niêm yết", fill_yellow, font_normal, "$#,##0.00"),
        ("Giá Output per 1M tokens", 5.00, "$/1M", "Anthropic List Pricing", "Giá niêm yết", fill_yellow, font_normal, "$#,##0.00"),
        ("Giá Prompt Cache Write per 1M tokens", 1.25, "$/1M", "1.25x Input Price (Anthropic)", "Ghi vào cache lượt đầu", fill_yellow, font_normal, "$#,##0.00"),
        ("Giá Prompt Cache Read per 1M tokens", 0.10, "$/1M", "0.10x Input Price (Anthropic)", "Đọc từ cache các lượt sau", fill_yellow, font_normal, "$#,##0.00"),
        ("Số lượt hội thoại (turns) trung bình / ticket", 6, "turns", "Thống kê thực tế CSKH", "6 lượt qua lại bot - khách", fill_yellow, font_normal, "#,##0"),
        ("System Prompt + Knowledge Base (Cache được)", 3000, "tokens", "System prompt + FAQ/Catalog", "Được cache từ turn 2", fill_yellow, font_normal, "#,##0"),
        ("Fresh Input token / lượt hội thoại", 500, "tokens/turn", "User message + context mới", "500 tokens x 6 turns", fill_yellow, font_normal, "#,##0"),
        ("Output token / lượt hội thoại", 150, "tokens/turn", "Bot response", "150 tokens x 6 turns", fill_yellow, font_normal, "#,##0"),
        ("Chi phí Cache Write (1 lần / ticket)", "=B27*B24/1000000", "$/ticket", "= B27 * B24 / 1M", "Lượt đầu tiên ghi cache", fill_gray, font_normal, "$#,##0.0000"),
        ("Chi phí Cache Read ((turns - 1) lần / ticket)", "=(B26-1)*B27*B25/1000000", "$/ticket", "= (B26 - 1) * B27 * B25 / 1M", "5 lượt sau đọc từ cache", fill_gray, font_normal, "$#,##0.0000"),
        ("Chi phí Fresh Input (turns lần / ticket)", "=B26*B28*B22/1000000", "$/ticket", "= B26 * B28 * B22 / 1M", "Input tin nhắn mới", fill_gray, font_normal, "$#,##0.0000"),
        ("Chi phí Output (turns lần / ticket)", "=B26*B29*B23/1000000", "$/ticket", "= B26 * B29 * B23 / 1M", "Output trả về của bot", fill_gray, font_normal, "$#,##0.0000"),
        ("Tổng chi phí LLM / ticket (CÓ CACHE)", "=SUM(B30:B33)", "$/ticket", "= B30 + B31 + B32 + B33", "Chi phí tối ưu thực tế", fill_gray, font_bold, "$#,##0.0000"),
        ("Tổng chi phí LLM / ticket (NẾU KHÔNG CACHE)", "=(B26*(B27+B28)*B22 + B26*B29*B23)/1000000", "$/ticket", "= (Turns*(Cache+Fresh)*Input + Turns*Out*Output)/1M", "Chi phí nếu không cache", fill_gray, font_normal, "$#,##0.0000"),
        ("% Tiết kiệm nhờ Prompt Caching", "=(B35-B34)/B35", "%", "= (B35 - B34) / B35", "Tiết kiệm rõ rệt 38% - 50%", fill_green, font_bold, "0.0%"),
        ("TỔNG CHI PHÍ LLM API / THÁNG (2,000 tickets)", "=B14*B34", "$/tháng", "= B14 * B34", "Chi phí LLM toàn bộ tháng", fill_summary, font_bold, "$#,##0.00")
    ]
    for idx, (label, val, unit, form, note, cfill, cfont, nfmt) in enumerate(s3_rows, start=21):
        ws1.cell(row=idx, column=1, value=label).font = font_bold if "TỔNG" in label or "%" in label else font_normal
        ws1.cell(row=idx, column=1).border = thin_border
        c = ws1.cell(row=idx, column=2, value=val); c.font = cfont; c.fill = cfill; c.border = thin_border; c.number_format = nfmt
        c.alignment = align_right if nfmt != "@" else align_left
        ws1.cell(row=idx, column=3, value=unit).font = font_normal; ws1.cell(row=idx, column=3).border = thin_border
        ws1.cell(row=idx, column=4, value=form).font = font_italic; ws1.cell(row=idx, column=4).border = thin_border
        ws1.cell(row=idx, column=5, value=note).font = font_note; ws1.cell(row=idx, column=5).border = thin_border

    # Section 4: Speech STT/TTS (Text-based -> $0)
    ws1["A39"] = "SECTION 4: CHI PHÍ THOẠI SPEECH (STT / TTS) - NẾU CÓ"
    ws1["A39"].font = font_section; ws1["A39"].fill = fill_section_header; ws1.merge_cells("A39:E39")
    
    ws1.cell(row=40, column=1, value="Sản phẩm là Text Chat Agent thuần túy (Không dùng Voice)").font = font_normal; ws1.cell(row=40, column=1).border = thin_border
    c = ws1.cell(row=40, column=2, value=0.00); c.font = font_bold; c.fill = fill_gray; c.number_format = "$#,##0.00"; c.alignment = align_right; c.border = thin_border
    ws1.cell(row=40, column=3, value="$/tháng").font = font_normal; ws1.cell(row=40, column=3).border = thin_border
    ws1.cell(row=40, column=4, value="Text-based Chat (Messenger, Zalo, Web)").font = font_italic; ws1.cell(row=40, column=4).border = thin_border
    ws1.cell(row=40, column=5, value="Không phát sinh STT/TTS").font = font_note; ws1.cell(row=40, column=5).border = thin_border

    # Section 5: Infra & Tools
    ws1["A42"] = "SECTION 5: CHI PHÍ HẠ TẦNG, VECTOR DB & LOGGING (INFRASTRUCTURE)"
    ws1["A42"].font = font_section; ws1["A42"].fill = fill_section_header; ws1.merge_cells("A42:E42")
    
    s5_rows = [
        ("Vector DB (Qdrant / Pinecone Serverless)", 0.0030, "$/ticket", "Pinecone pricing $0.003/query+embedding", "Tra cứu semantic chính sách & FAQ", fill_yellow, font_normal, "$#,##0.0000"),
        ("Logging, Tracing & Evals (LangSmith / OpenTelemetry)", 0.0015, "$/ticket", "LangSmith $0.0015/trace", "Lưu log & tracking audit trail", fill_yellow, font_normal, "$#,##0.0000"),
        ("Cloud Server, Egress & Webhook Execution", 0.0015, "$/ticket", "AWS Lambda / Fly.io execution", "Xử lý webhook & API ERP/CRM", fill_yellow, font_normal, "$#,##0.0000"),
        ("Tổng chi phí Hạ tầng / ticket", "=SUM(B43:B45)", "$/ticket", "= B43 + B44 + B45", "Tổng Infra trên 1 ticket", fill_gray, font_bold, "$#,##0.0000"),
        ("TỔNG CHI PHÍ INFRA / THÁNG (2,000 tickets)", "=B14*B46", "$/tháng", "= B14 * B46", "Tổng Infra toàn bộ tháng", fill_summary, font_bold, "$#,##0.00")
    ]
    for idx, (label, val, unit, form, note, cfill, cfont, nfmt) in enumerate(s5_rows, start=43):
        ws1.cell(row=idx, column=1, value=label).font = font_bold if "TỔNG" in label else font_normal
        ws1.cell(row=idx, column=1).border = thin_border
        c = ws1.cell(row=idx, column=2, value=val); c.font = cfont; c.fill = cfill; c.border = thin_border; c.number_format = nfmt; c.alignment = align_right
        ws1.cell(row=idx, column=3, value=unit).font = font_normal; ws1.cell(row=idx, column=3).border = thin_border
        ws1.cell(row=idx, column=4, value=form).font = font_italic; ws1.cell(row=idx, column=4).border = thin_border
        ws1.cell(row=idx, column=5, value=note).font = font_note; ws1.cell(row=idx, column=5).border = thin_border

    # Section 6: Retry & Timeout
    ws1["A49"] = "SECTION 6: CHI PHÍ RETRY DO TIMEOUT & LỖI API (RETRY COST)"
    ws1["A49"].font = font_section; ws1["A49"].fill = fill_section_header; ws1.merge_cells("A49:E49")
    
    ws1.cell(row=50, column=1, value="Tỷ lệ Retry dự phòng (% job bị lỗi/timeout phải gọi lại)").font = font_normal; ws1.cell(row=50, column=1).border = thin_border
    c = ws1.cell(row=50, column=2, value=0.08); c.font = font_bold; c.fill = fill_yellow; c.number_format = "0.0%"; c.alignment = align_right; c.border = thin_border
    ws1.cell(row=50, column=3, value="%").font = font_normal; ws1.cell(row=50, column=3).border = thin_border
    ws1.cell(row=50, column=4, value="Benchmark thực tế 5% - 10%").font = font_italic; ws1.cell(row=50, column=4).border = thin_border
    ws1.cell(row=50, column=5, value="Không được để 0%").font = font_note; ws1.cell(row=50, column=5).border = thin_border
    
    ws1.cell(row=51, column=1, value="TỔNG CHI PHÍ RETRY / THÁNG").font = font_bold; ws1.cell(row=51, column=1).border = thin_border
    c = ws1.cell(row=51, column=2, value="=B37*B50"); c.font = font_bold; c.fill = fill_summary; c.number_format = "$#,##0.00"; c.alignment = align_right; c.border = thin_border
    ws1.cell(row=51, column=3, value="$/tháng").font = font_normal; ws1.cell(row=51, column=3).border = thin_border
    ws1.cell(row=51, column=4, value="= B37 * B50 (8% của Tổng LLM API)").font = font_italic; ws1.cell(row=51, column=4).border = thin_border
    ws1.cell(row=51, column=5, value="Chi phí token phát sinh thêm").font = font_note; ws1.cell(row=51, column=5).border = thin_border

    # Section 7: HITL
    ws1["A53"] = "SECTION 7: CHI PHÍ NHÂN SỰ KIỂM SOÁT & XỬ LÝ (HITL COST)"
    ws1["A53"].font = font_section; ws1["A53"].fill = fill_section_header; ws1.merge_cells("A53:E53")
    
    s7_rows = [
        ("--- THÀNH PHẦN 1: QA NỘI BỘ (ÁP DỤNG MỌI MÔ HÌNH) ---", "", "", "", "", fill_table_header, font_bold, "@"),
        ("Tỷ lệ ticket chọn audit QA ngẫu nhiên (% QA sampling)", 0.05, "%", "Giả định audit 5% tổng volume", "Đảm bảo chất lượng & eval", fill_yellow, font_normal, "0.0%"),
        ("Thời gian nhân sự QA audit 1 ticket", 2.0, "phút/ticket", "Đọc log & chấm điểm prompt", "2 phút = 0.0333 giờ", fill_yellow, font_normal, "0.0"),
        ("Đơn giá nhân sự QA nội bộ ($/giờ)", 6.00, "$/giờ", "Lương 156,000 VND/giờ", "Lương CS Lead / QA Specialist", fill_yellow, font_normal, "$#,##0.00"),
        ("Chi phí QA nội bộ / tháng", "=B14*B55*(B56/60)*B57", "$/tháng", "= B14 * B55 * (B56/60) * B57", "Chi phí QA kiểm định chất lượng", fill_gray, font_bold, "$#,##0.00"),
        ("--- THÀNH PHẦN 2: XỬ LÝ ESCALATION (CHỈ BIẾN THỂ B) ---", "", "", "", "", fill_table_header, font_bold, "@"),
        ("Thời gian nhân sự xử lý 1 ca escalate", 6.0, "phút/ca", "Nhân sự tiếp nhận & xử lý dứt điểm", "6 phút = 0.10 giờ", fill_yellow, font_normal, "0.0"),
        ("Đơn giá nhân sự xử lý escalation ($/giờ)", 6.00, "$/giờ", "Lương 156,000 VND/giờ", "Lương CS Agent vận hành", fill_yellow, font_normal, "$#,##0.00"),
        ("Chi phí Escalation / tháng (360 ca chuyển người)", "=B17*(B60/60)*B61", "$/tháng", "= B17 * (B60/60) * B61", "Chi phí xử lý ca AI làm không xong", fill_gray, font_bold, "$#,##0.00"),
        ("TỔNG HITL BIẾN THỂ A (Khách tự xử lý escalate, bạn chỉ QA)", "=B58", "$/tháng", "= B58", "COGS nếu là nhà bán SaaS", fill_summary, font_bold, "$#,##0.00"),
        ("TỔNG HITL BIẾN THỂ B (Bạn chịu trách nhiệm xử lý trọn gói)", "=B58+B62", "$/tháng", "= B58 + B62", "COGS nếu bán Outcome / Managed", fill_summary, font_bold, "$#,##0.00")
    ]
    for idx, (label, val, unit, form, note, cfill, cfont, nfmt) in enumerate(s7_rows, start=54):
        ws1.cell(row=idx, column=1, value=label).font = font_bold if "---" in label or "TỔNG" in label else font_normal
        ws1.cell(row=idx, column=1).border = thin_border
        c = ws1.cell(row=idx, column=2, value=val if val != "" else "")
        c.font = cfont; c.fill = cfill; c.border = thin_border; c.number_format = nfmt
        c.alignment = align_right if val != "" and nfmt != "@" else align_left
        ws1.cell(row=idx, column=3, value=unit).font = font_normal; ws1.cell(row=idx, column=3).border = thin_border
        ws1.cell(row=idx, column=4, value=form).font = font_italic; ws1.cell(row=idx, column=4).border = thin_border
        ws1.cell(row=idx, column=5, value=note).font = font_note; ws1.cell(row=idx, column=5).border = thin_border

    # Section 8: Overhead
    ws1["A66"] = "SECTION 8: CHI PHÍ VẬN HÀNH CHUNG & PHÂN BỔ (OVERHEAD)"
    ws1["A66"].font = font_section; ws1["A66"].fill = fill_section_header; ws1.merge_cells("A66:E66")
    
    ws1.cell(row=67, column=1, value="Chi phí R&D, Platform Maintenance, Dev Ops phân bổ / tháng").font = font_normal; ws1.cell(row=67, column=1).border = thin_border
    c = ws1.cell(row=67, column=2, value=50.00); c.font = font_bold; c.fill = fill_yellow; c.number_format = "$#,##0.00"; c.alignment = align_right; c.border = thin_border
    ws1.cell(row=67, column=3, value="$/tháng").font = font_normal; ws1.cell(row=67, column=3).border = thin_border
    ws1.cell(row=67, column=4, value="Phân bổ chi phí cố định cho account").font = font_italic; ws1.cell(row=67, column=4).border = thin_border
    ws1.cell(row=67, column=5, value="Phân bổ vận hành").font = font_note; ws1.cell(row=67, column=5).border = thin_border

    # Section 9: ROLL-UP & COST/JOB CALCULATION
    ws1["A69"] = "TỔNG HỢP CHI PHÍ & TÍNH TOÁN COST/JOB (KẾT QUẢ ĐẦU RA TAB 1)"
    ws1["A69"].font = font_section; ws1["A69"].fill = fill_subsection; ws1.merge_cells("A69:E69")
    
    rollup_rows = [
        ("Tổng Chi phí Tháng - BIẾN THỂ A (Khách tự escalate, không Overhead)", "=B37+B40+B47+B51+B63", "$/tháng", "= LLM + Speech + Infra + Retry + HITL_A", "Tổng COGS trực tiếp SaaS", fill_gray, font_bold, "$#,##0.00"),
        ("Tổng Chi phí Tháng - BIẾN THỂ A (Có tính Overhead)", "=B70+B67", "$/tháng", "= Chi phí trực tiếp + Overhead", "COGS đầy đủ có phân bổ R&D", fill_gray, font_normal, "$#,##0.00"),
        ("Tổng Chi phí Tháng - BIẾN THỂ B (Bạn chịu escalation, không Overhead)", "=B37+B40+B47+B51+B64", "$/tháng", "= LLM + Speech + Infra + Retry + HITL_B", "Tổng COGS trực tiếp Outcome", fill_gray, font_bold, "$#,##0.00"),
        ("Tổng Chi phí Tháng - BIẾN THỂ B (Có tính Overhead)", "=B72+B67", "$/tháng", "= Chi phí trực tiếp + Overhead", "COGS đầy đủ có phân bổ R&D", fill_gray, font_normal, "$#,##0.00"),
        ("---------------------------------------------------------------------------------", "", "", "", "", fill_table_header, font_bold, "@"),
        ("Số Job HOÀN THÀNH (MẪU SỐ THẬT)", "=B16", "jobs", "= B16 (1,640 jobs hoàn thành)", "Mẫu số chuẩn để chia đơn giá", fill_green, font_bold, "#,##0"),
        ("COST / JOB HOÀN THÀNH - BIẾN THỂ A (SaaS thuần)", "=B70/B75", "$/job", "= B70 / B75 ($59.54 / 1,640)", "Cost/Job nếu khách tự gánh escalate", fill_green, font_bold, "$#,##0.0000"),
        ("COST / JOB HOÀN THÀNH - BIẾN THỂ B (Managed Outcome)", "=B72/B75", "$/job", "= B72 / B75 ($275.54 / 1,640)", "Cost/Job bán Outcome cam kết trọn gói", fill_green, font_bold, "$#,##0.0000"),
        ("Quy đổi Cost/Job Biến thể B sang VNĐ (tỷ giá 26,000)", "=B77*26000", "VNĐ/job", "= B77 * 26,000 VND/USD", "Tương đương khoảng 4,368 VNĐ / job", fill_green, font_bold, "#,##0 VNĐ")
    ]
    
    for idx, (label, val, unit, form, note, cfill, cfont, nfmt) in enumerate(rollup_rows, start=70):
        ws1.cell(row=idx, column=1, value=label).font = font_bold if "COST / JOB" in label or "MẪU SỐ" in label or "BIẾN THỂ" in label else font_normal
        ws1.cell(row=idx, column=1).border = thin_border
        c = ws1.cell(row=idx, column=2, value=val if val != "" else "")
        c.font = cfont; c.fill = cfill; c.border = thin_border; c.number_format = nfmt
        c.alignment = align_right if val != "" and nfmt != "@" else align_left
        ws1.cell(row=idx, column=3, value=unit).font = font_normal; ws1.cell(row=idx, column=3).border = thin_border
        ws1.cell(row=idx, column=4, value=form).font = font_italic; ws1.cell(row=idx, column=4).border = thin_border
        ws1.cell(row=idx, column=5, value=note).font = font_note; ws1.cell(row=idx, column=5).border = thin_border

    ws1.column_dimensions["A"].width = 48
    ws1.column_dimensions["B"].width = 24
    ws1.column_dimensions["C"].width = 16
    ws1.column_dimensions["D"].width = 46
    ws1.column_dimensions["E"].width = 42

    # =========================================================================
    # TAB 2: 2_Pricing
    # =========================================================================
    ws2 = wb.create_sheet(title="2_Pricing")
    ws2.views.sheetView[0].showGridLines = True
    
    ws2["A1"] = "TAB 2: ĐỊNH GIÁ BÁN, GROSS MARGIN & SENSITIVITY STRESS TEST"
    ws2["A1"].font = font_title
    ws2["A2"] = "Xác định Giá sàn, Giá trần, Giá đề xuất và Kiểm tra ngưỡng sinh tử Breakeven Containment"
    ws2["A2"].font = font_italic
    
    # Section: Floor and Ceiling
    ws2["A4"] = "1. XÁC LẬP VÙNG GIÁ BÁN: GIÁ SÀN (COST x 3) VÀ GIÁ TRẦN (ANCHORING)"
    ws2["A4"].font = font_section; ws2["A4"].fill = fill_section_header; ws2.merge_cells("A4:E4")
    
    for c_idx, h in enumerate(headers_5col, start=1):
        c = ws2.cell(row=5, column=c_idx, value=h); c.font = font_header; c.fill = fill_table_header; c.border = thin_border; c.alignment = align_center
        
    p1_rows = [
        ("Cost/Job cơ sở (Biến thể B - Managed Outcome)", "='1_Cost_Job'!B77", "$/job", "Lấy từ Tab 1 (='1_Cost_Job'!B77)", "Chi phí trực tiếp đầy đủ 5 thành phần", fill_gray, font_bold, "$#,##0.000"),
        ("Giá SÀN tối thiểu (= 3 x Cost/Job)", "=B6*3", "$/job", "= B6 * 3", "Đảm bảo Gross Margin >= 66.7% cover Overhead/R&D", fill_green, font_bold, "$#,##0.000"),
        ("Neo giá trần 1: Theo giá trị tiết kiệm cho khách", 0.75, "$/job", "Tiết kiệm lương ca đêm ($400/tháng cho 1,640 jobs ~ $0.24/job, lấy 25-30% tổng ROI)", "Khách sẵn sàng chi khi thấy ROI 3x", fill_yellow, font_normal, "$#,##0.000"),
        ("Neo giá trần 2: Theo chi phí nhân sự con người", 0.90, "$/job", "Chi phí 1 ticket do người làm tại VN/SEA ≈ $0.60 - $0.90", "Tính theo 50-70% chi phí nhân sự", fill_yellow, font_normal, "$#,##0.000"),
        ("Benchmark đối thủ thị trường (Intercom Fin)", 0.99, "$/resolution", "Intercom Fin niêm yết $0.99/resolution", "Mức trần ngành toàn cầu", fill_yellow, font_normal, "$#,##0.000"),
        ("GIÁ BÁN ĐỀ XUẤT (PROPOSED PRICE)", 0.60, "$/resolution", "Nằm giữa Giá sàn ($0.504) và Giá trần ($0.75 - $0.99)", "Rất cạnh tranh với Intercom ($0.99) và Zendesk ($1.50)", fill_yellow, font_bold, "$#,##0.000"),
        ("Quy đổi Giá bán đề xuất sang VNĐ (tỷ giá 26,000)", "=B11*26000", "VNĐ/job", "= B11 * 26,000", "Khoảng 15,600 VNĐ / ticket hoàn thành", fill_green, font_bold, "#,##0 VNĐ"),
        ("Gross Margin $ trên 1 Job hoàn thành", "=B11-B6", "$/job", "= Giá bán (B11) - Cost/Job (B6)", "Lợi nhuận gộp trên mỗi job hoàn thành", fill_gray, font_normal, "$#,##0.000"),
        ("GROSS MARGIN % TẠI MỨC 82% CONTAINMENT", "=(B11-B6)/B11", "%", "= (B11 - B6) / B11", "MỤC TIÊU >= 60% (Đạt chuẩn xuất sắc 72.0%)", fill_green, font_bold, "0.0%")
    ]
    for idx, (label, val, unit, form, note, cfill, cfont, nfmt) in enumerate(p1_rows, start=6):
        ws2.cell(row=idx, column=1, value=label).font = font_bold if "GIÁ" in label or "GROSS MARGIN %" in label or "Cost/Job" in label else font_normal
        ws2.cell(row=idx, column=1).border = thin_border
        c = ws2.cell(row=idx, column=2, value=val); c.font = cfont; c.fill = cfill; c.border = thin_border; c.number_format = nfmt; c.alignment = align_right
        ws2.cell(row=idx, column=3, value=unit).font = font_normal; ws2.cell(row=idx, column=3).border = thin_border
        ws2.cell(row=idx, column=4, value=form).font = font_italic; ws2.cell(row=idx, column=4).border = thin_border
        ws2.cell(row=idx, column=5, value=note).font = font_note; ws2.cell(row=idx, column=5).border = thin_border

    # Section: Sensitivity Stress Test
    ws2["A16"] = "2. BẢNG SENSITIVITY STRESS TEST: ĐỘ NHẠY THEO CONTAINMENT RATE"
    ws2["A16"].font = font_section; ws2["A16"].fill = fill_section_header; ws2.merge_cells("A16:G16")
    
    sens_headers = ["Tỷ lệ Containment %", "Số Job Hoàn thành", "Số Ca Escalate", "Chi phí Tháng Biến thể B", "Cost / Job Hoàn thành", "Gross Margin % ở giá $0.60", "Đánh giá An toàn Tài chính"]
    for c_idx, h in enumerate(sens_headers, start=1):
        c = ws2.cell(row=17, column=c_idx, value=h); c.font = font_header; c.fill = fill_table_header; c.border = thin_border; c.alignment = align_center
        
    rates = [0.50, 0.60, 0.684, 0.70, 0.80, 0.82, 0.90]
    for r_idx, rate in enumerate(rates, start=18):
        c1 = ws2.cell(row=r_idx, column=1, value=rate); c1.font = font_bold; c1.border = thin_border; c1.alignment = align_center; c1.number_format = "0.0%"
        c2 = ws2.cell(row=r_idx, column=2, value=f"=2000*A{r_idx}"); c2.font = font_normal; c2.border = thin_border; c2.alignment = align_right; c2.number_format = "#,##0"
        c3 = ws2.cell(row=r_idx, column=3, value=f"=2000*(1-A{r_idx})"); c3.font = font_normal; c3.border = thin_border; c3.alignment = align_right; c3.number_format = "#,##0"
        c4 = ws2.cell(row=r_idx, column=4, value=f"='1_Cost_Job'!B70 + C{r_idx}*('1_Cost_Job'!B60/60)*'1_Cost_Job'!B61"); c4.font = font_normal; c4.border = thin_border; c4.alignment = align_right; c4.number_format = "$#,##0.00"
        c5 = ws2.cell(row=r_idx, column=5, value=f"=D{r_idx}/B{r_idx}"); c5.font = font_bold; c5.border = thin_border; c5.alignment = align_right; c5.number_format = "$#,##0.0000"
        c6 = ws2.cell(row=r_idx, column=6, value=f"=($B$11-E{r_idx})/$B$11"); c6.font = font_bold; c6.border = thin_border; c6.alignment = align_right; c6.number_format = "0.0%"
        
        if rate < 0.60:
            c6.fill = fill_red; c7_val = "🟥 NGUY HIỂM (GM < 50% - Lỗ nặng/Biên mỏng)"
        elif rate < 0.684:
            c6.fill = fill_yellow; c7_val = "🟨 CẢNH BÁO (Dưới ngưỡng chuẩn 60% GM)"
        elif rate == 0.684:
            c6.fill = fill_green; c7_val = "🟩 BREAKEVEN NGƯỠNG SỐNG CÒN (GM = 60.0%)"
        elif rate == 0.82:
            c6.fill = fill_green; c7_val = "🟩 BASELINE HIỆN TẠI (GM = 72.0% - Rất tốt)"
        else:
            c6.fill = fill_green; c7_val = "🟩 AN TOÀN CAO (Biên lợi nhuận xuất sắc)"
            
        c7 = ws2.cell(row=r_idx, column=7, value=c7_val); c7.font = font_bold if rate in [0.684, 0.82] else font_normal; c7.border = thin_border; c7.alignment = align_left

    # Section: Breakeven Threshold
    ws2["A27"] = "3. NGƯỠNG SINH TỬ BREAKEVEN CONTAINMENT (KẾT NỐI EVALS -> PRICING)"
    ws2["A27"].font = font_section; ws2["A27"].fill = fill_subsection; ws2.merge_cells("A27:E27")
    
    be_rows = [
        ("Giá bán đề xuất", "=$B$11", "$/resolution", "Mức giá niêm yết", "Giá bán chốt", fill_gray, font_bold, "$#,##0.000"),
        ("Gross Margin mục tiêu tối thiểu", 0.60, "%", "Ngưỡng an toàn ngành SaaS/AI", "Target chuẩn ngành", fill_yellow, font_bold, "0.0%"),
        ("Cost/Job tối đa cho phép để đạt GM 60%", "=B28*(1-B29)", "$/job", "= Giá bán * (1 - GM mục tiêu) = $0.60 * 0.40", "Trần chi phí đơn vị", fill_gray, font_bold, "$#,##0.000"),
        ("BREAKEVEN CONTAINMENT RATE YÊU CẦU", 0.684, "%", "Giải phương trình: (59.54 + 0.60*(2000-R)) / R <= $0.240 -> R >= 1,368 -> 68.4%", "NGƯỠNG SỐNG CÒN CỦA Evals", fill_green, font_bold, "0.0%"),
        ("Containment Rate thực tế từ Evals", "='1_Cost_Job'!B15", "%", "Lấy từ kết quả Eval / Giả định Tab 1", "Kết quả thực tế hôm nay", fill_green, font_bold, "0.0%"),
        ("Biên an toàn (Safety Margin = Thực tế - Breakeven)", "=B32-B31", "%", "= B32 - B31 (82.0% - 68.4%)", "VÙNG AN TOÀN DƯ +13.6% (Đảm bảo mô hình luôn có lãi)", fill_green, font_bold, "0.0%")
    ]
    for idx, (label, val, unit, form, note, cfill, cfont, nfmt) in enumerate(be_rows, start=28):
        ws2.cell(row=idx, column=1, value=label).font = font_bold if "BREAKEVEN" in label or "Biên an toàn" in label or "Cost/Job" in label else font_normal
        ws2.cell(row=idx, column=1).border = thin_border
        c = ws2.cell(row=idx, column=2, value=val); c.font = cfont; c.fill = cfill; c.border = thin_border; c.number_format = nfmt; c.alignment = align_right
        ws2.cell(row=idx, column=3, value=unit).font = font_normal; ws2.cell(row=idx, column=3).border = thin_border
        ws2.cell(row=idx, column=4, value=form).font = font_italic; ws2.cell(row=idx, column=4).border = thin_border
        ws2.cell(row=idx, column=5, value=note).font = font_note; ws2.cell(row=idx, column=5).border = thin_border

    ws2.column_dimensions["A"].width = 46
    ws2.column_dimensions["B"].width = 22
    ws2.column_dimensions["C"].width = 16
    ws2.column_dimensions["D"].width = 48
    ws2.column_dimensions["E"].width = 38
    ws2.column_dimensions["F"].width = 26
    ws2.column_dimensions["G"].width = 46

    # =========================================================================
    # TAB 3: 3_Value_Metric
    # =========================================================================
    ws3 = wb.create_sheet(title="3_Value_Metric")
    ws3.views.sheetView[0].showGridLines = True
    
    ws3["A1"] = "TAB 3: ĐÁNH GIÁ MA TRẬN ATTRIBUTION x AUTONOMY & CHỌN VALUE METRIC"
    ws3["A1"].font = font_title
    ws3["A2"] = "Chấm điểm khả năng quy kết giá trị (Attribution) và mức độ tự động hóa (Autonomy)"
    ws3["A2"].font = font_italic
    
    # Attribution Scorecard
    ws3["A4"] = "1. CHẤM ĐIỂM ATTRIBUTION (KHẢ NĂNG ĐO LƯỜNG & QUY KẾT KẾT QUẢ CHO AI)"
    ws3["A4"].font = font_section; ws3["A4"].fill = fill_section_header; ws3.merge_cells("A4:D4")
    
    vm_headers = ["Tiêu chí đánh giá", "Điểm (1-5)", "Bằng chứng thực tế của sản phẩm", "Ghi chú chấm điểm"]
    for c_idx, h in enumerate(vm_headers, start=1):
        c = ws3.cell(row=5, column=c_idx, value=h); c.font = font_header; c.fill = fill_table_header; c.border = thin_border; c.alignment = align_center
        
    attr_criteria = [
        ("Khách hàng có đồng thuận về định nghĩa 'ticket giải quyết xong' không?", 5, "Định nghĩa rõ: Khách xác nhận 'Cảm ơn/Đã hiểu' hoặc không gửi thêm tin nhắn sau 24h", "Định nghĩa chặt chẽ không tranh cãi"),
        ("Hệ thống có log rõ ràng tin nhắn cuối cùng là bot xử lý mà không cần chuyển người?", 5, "Audit log lưu lại toàn bộ message trace, trạng thái closed by AI trong cơ sở dữ liệu", "Attribution kỹ thuật 100%"),
        ("Có thể đo lường tỷ lệ CSAT hoặc tỷ lệ no-reopen tự động không?", 4, "Tự động gửi khảo sát 1-click CSAT sau khi đóng ticket và đo reopen rate trong 24h", "Đo lường định lượng chính xác"),
        ("Khách hàng có phân biệt được ticket do AI giải quyết vs ticket nhân sự làm?", 5, "Phân tách rõ ràng trên Dashboard báo cáo: Ticket Resolved by AI vs Escalate to Human", "Rõ ràng minh bạch trên hoá đơn"),
        ("Có cơ chế tự động miễn phí/hoàn tiền khi ticket bị escalate hoặc khách phàn nàn?", 5, "Chỉ tính tiền khi status = 'AI Resolved'; toàn bộ ca Escalate tính $0 tiền phí outcome", "Khách hàng hoàn toàn tin tưởng")
    ]
    for idx, (crit, score, ev, note) in enumerate(attr_criteria, start=6):
        ws3.cell(row=idx, column=1, value=crit).font = font_normal; ws3.cell(row=idx, column=1).border = thin_border
        c = ws3.cell(row=idx, column=2, value=score); c.font = font_bold; c.fill = fill_yellow; c.alignment = align_center; c.border = thin_border
        ws3.cell(row=idx, column=3, value=ev).font = font_italic; ws3.cell(row=idx, column=3).border = thin_border
        ws3.cell(row=idx, column=4, value=note).font = font_note; ws3.cell(row=idx, column=4).border = thin_border
        
    ws3.cell(row=11, column=1, value="TỔNG ĐIỂM ATTRIBUTION (Max: 25)").font = font_bold; ws3.cell(row=11, column=1).border = thin_border
    c = ws3.cell(row=11, column=2, value="=SUM(B6:B10)"); c.font = font_bold; c.fill = fill_green; c.alignment = align_center; c.border = thin_border
    ws3.cell(row=11, column=3, value="ATTRIBUTION CAO (>= 20/25) -> ĐỦ ĐIỀU KIỆN BÁN OUTCOME").font = font_bold; ws3.cell(row=11, column=3).border = thin_border
    ws3.cell(row=11, column=4, value="Kết quả xuất sắc").font = font_note; ws3.cell(row=11, column=4).border = thin_border

    # Autonomy Scorecard
    ws3["A13"] = "2. CHẤM ĐIỂM AUTONOMY (MỨC ĐỘ TỰ ĐỘNG HÓA CỦA AI TRONG QUY TRÌNH)"
    ws3["A13"].font = font_section; ws3["A13"].fill = fill_section_header; ws3.merge_cells("A13:D13")
    
    for c_idx, h in enumerate(vm_headers, start=1):
        c = ws3.cell(row=14, column=c_idx, value=h); c.font = font_header; c.fill = fill_table_header; c.border = thin_border; c.alignment = align_center
        
    auto_criteria = [
        ("AI có tự động kết nối API CRM/ERP tra cứu thông tin đơn hàng/vận đơn không?", 5, "Agent tự gọi tool check_order_status(), get_policy_rag() qua API", "Tự động hóa dữ liệu đầu vào"),
        ("AI có tự ra quyết định phản hồi và đóng ticket mà không cần người bấm duyệt?", 5, "Tự động gửi câu trả lời và cập nhật trạng thái đơn hàng trên Pancake/Zalo OA", "Không cần human approval từng tin"),
        ("Tỷ lệ Containment rate đạt mức cao trong môi trường thực tế?", 4, "Đạt 82.0% Containment rate trên tập test pilot", "Tự xử lý phần lớn ticket"),
        ("Hệ thống có tự động nhận diện intent phức tạp để escalate an toàn?", 5, "Fallback an toàn 100% khi gặp khiếu nại gay gắt hoặc ngoài tri thức", "Tránh hallucination gây hại"),
        ("Mức độ can thiệp của con người trong quá trình xử lý ticket?", 4, "Chỉ can thiệp ở 18% ca escalate; 82% còn lại hoàn toàn không có người chạm vào", "Mức độ can thiệp cực thấp")
    ]
    for idx, (crit, score, ev, note) in enumerate(auto_criteria, start=15):
        ws3.cell(row=idx, column=1, value=crit).font = font_normal; ws3.cell(row=idx, column=1).border = thin_border
        c = ws3.cell(row=idx, column=2, value=score); c.font = font_bold; c.fill = fill_yellow; c.alignment = align_center; c.border = thin_border
        ws3.cell(row=idx, column=3, value=ev).font = font_italic; ws3.cell(row=idx, column=3).border = thin_border
        ws3.cell(row=idx, column=4, value=note).font = font_note; ws3.cell(row=idx, column=4).border = thin_border
        
    ws3.cell(row=20, column=1, value="TỔNG ĐIỂM AUTONOMY (Max: 25)").font = font_bold; ws3.cell(row=20, column=1).border = thin_border
    c = ws3.cell(row=20, column=2, value="=SUM(B15:B19)"); c.font = font_bold; c.fill = fill_green; c.alignment = align_center; c.border = thin_border
    ws3.cell(row=20, column=3, value="AUTONOMY CAO (>= 20/25) -> AI TỰ CHẠY ĐỘC LẬP TỐT").font = font_bold; ws3.cell(row=20, column=3).border = thin_border
    ws3.cell(row=20, column=4, value="Kết quả xuất sắc").font = font_note; ws3.cell(row=20, column=4).border = thin_border

    # Recommendation and Benchmark
    ws3["A22"] = "3. KẾT QUẢ GỢI Ý VALUE METRIC TỪ MA TRẬN & BENCHMARK THỊ TRƯỜNG"
    ws3["A22"].font = font_section; ws3["A22"].fill = fill_subsection; ws3.merge_cells("A22:D22")
    
    ws3.cell(row=23, column=1, value="Gợi ý từ Ma trận Attribution x Autonomy:").font = font_bold; ws3.cell(row=23, column=1).border = thin_border
    c = ws3.cell(row=23, column=2, value="OUTCOME-BASED hoặc HYBRID (Base + Outcome)"); c.font = font_bold; c.fill = fill_green; c.border = thin_border; ws3.merge_cells("B23:D23")
    
    ws3["A25"] = "4. BENCHMARK 3 SẢN PHẨM THẬT CÙNG LOẠI JOB"
    ws3["A25"].font = font_section; ws3["A25"].fill = fill_section_header; ws3.merge_cells("A25:D25")
    
    bm_headers = ["Sản phẩm Benchmark", "Value Metric áp dụng", "Mức giá niêm yết hiện hành", "Nguồn tham chiếu xác thực"]
    for c_idx, h in enumerate(bm_headers, start=1):
        c = ws3.cell(row=26, column=c_idx, value=h); c.font = font_header; c.fill = fill_table_header; c.border = thin_border; c.alignment = align_center
        
    bm_rows = [
        ("1. Intercom Fin", "Outcome (Per Resolution)", "$0.99 / resolution (Không tính phí seat khi cắm Helpdesk ngoài)", "https://fin.ai/pricing/"),
        ("2. Zendesk AI", "Outcome (Automated Resolution Tiers)", "$1.50 / automated resolution", "https://zendesk.com/pricing"),
        ("3. Salesforce Agentforce", "Usage / Hybrid (Flex Credits + User)", "$2.00 / conversation hoặc $0.10 / action", "https://salesforce.com/agentforce/pricing/")
    ]
    for idx, (prod, vm, pr, src) in enumerate(bm_rows, start=27):
        ws3.cell(row=idx, column=1, value=prod).font = font_bold; ws3.cell(row=idx, column=1).border = thin_border
        ws3.cell(row=idx, column=2, value=vm).font = font_normal; ws3.cell(row=idx, column=2).border = thin_border
        ws3.cell(row=idx, column=3, value=pr).font = font_normal; ws3.cell(row=idx, column=3).border = thin_border
        ws3.cell(row=idx, column=4, value=src).font = font_italic; ws3.cell(row=idx, column=4).border = thin_border

    # Decision Note
    ws3["A31"] = "5. CHỐT VALUE METRIC & DECISION NOTE (3 CÂU BẢO VỆ)"
    ws3["A31"].font = font_section; ws3["A31"].fill = fill_section_header; ws3.merge_cells("A31:D31")
    
    decisions = [
        ("1. Đơn vị tính tiền chốt:", "HYBRID: Phí nền $99/tháng (bao gồm 200 resolutions) + $0.50/resolution vượt định mức (hoặc $0.60/resolution thuần).", fill_yellow),
        ("2. Bằng chứng Attribution & Autonomy:", "Attribution đạt 24/25 và Autonomy đạt 23/25, chứng minh qua Eval 82% Containment rate và hệ thống log trace không thể tranh cãi.", fill_yellow),
        ("3. Lý do thị trường (Market Justification):", "SME Việt Nam cần phí nền thấp để yên tâm setup, đồng thời đơn giá $0.50 - $0.60/resolution giúp khách hàng thấy rẻ hơn 50% so với thuê người trực đêm mà không sợ hóa đơn bị thả nổi.", fill_yellow)
    ]
    for idx, (q, ans, cfill) in enumerate(decisions, start=32):
        ws3.cell(row=idx, column=1, value=q).font = font_bold; ws3.cell(row=idx, column=1).border = thin_border
        c = ws3.cell(row=idx, column=2, value=ans); c.font = font_normal; c.fill = cfill; c.border = thin_border
        ws3.merge_cells(start_row=idx, start_column=2, end_row=idx, end_column=4)

    ws3.column_dimensions["A"].width = 38
    ws3.column_dimensions["B"].width = 25
    ws3.column_dimensions["C"].width = 50
    ws3.column_dimensions["D"].width = 35

    # =========================================================================
    # TAB 4: 4_Channel_Fit
    # =========================================================================
    ws4 = wb.create_sheet(title="4_Channel_Fit")
    ws4.views.sheetView[0].showGridLines = True
    
    ws4["A1"] = "TAB 4: ĐÁNH GIÁ KÊNH GTM & CHANNEL AFFORDABILITY TEST"
    ws4["A1"].font = font_title
    ws4["A2"] = "Tính toán Ngân sách CAC, Deal Capacity của Sales Rep và Chấm điểm lựa chọn 1 Kênh Phân Phối Cốt Lõi"
    ws4["A2"].font = font_italic
    
    # Section 1: CAC Budget
    ws4["A4"] = "1. TÍNH TOÁN NGÂN SÁCH CAC TỐI ĐA CHO PHÉP (CHANNEL AFFORDABILITY)"
    ws4["A4"].font = font_section; ws4["A4"].fill = fill_section_header; ws4.merge_cells("A4:E4")
    
    for c_idx, h in enumerate(headers_5col, start=1):
        c = ws4.cell(row=5, column=c_idx, value=h); c.font = font_header; c.fill = fill_table_header; c.border = thin_border; c.alignment = align_center
        
    cac_rows = [
        ("ARPU trung bình / khách hàng / tháng", 300.00, "$/tháng", "Ước tính gói Hybrid ~500 resolutions/tháng", "Doanh thu trung bình trên 1 shop SME", fill_yellow, font_bold, "$#,##0.00"),
        ("Gross Margin % lấy từ Tab 2", "='2_Pricing'!B14", "%", "Lấy từ Tab 2 (='2_Pricing'!B14)", "Biên lợi nhuận gộp", fill_gray, font_bold, "0.0%"),
        ("Phân khúc khách hàng mục tiêu", "SMB", "Segment", "SMB / Mid-market / Enterprise", "Khách hàng E-commerce & Retail SME", fill_yellow, font_normal, "@"),
        ("Thời gian hoàn vốn CAC tối đa cho phép (Payback Months)", 12, "tháng", "Chuẩn Bessemer 2024: SMB < 12 tháng", "Ngưỡng hoàn vốn tối đa", fill_yellow, font_bold, "0"),
        ("NGÂN SÁCH CAC TỐI ĐA CHO PHÉP / KHÁCH HÀNG", "=B6*B7*B9", "$/khách", "= ARPU * Gross Margin * Payback Months", "SỐ TIỀN TỐI ĐA ĐƯỢC PHÉP CHI ĐỂ CÓ 1 KHÁCH", fill_green, font_bold, "$#,##0.00"),
        ("Quy đổi Ngân sách CAC sang VNĐ (tỷ giá 26,000)", "=B10*26000", "VNĐ/khách", "= B10 * 26,000", "Tương đương ~67.3 triệu VNĐ / khách", fill_green, font_bold, "#,##0 VNĐ")
    ]
    for idx, (label, val, unit, form, note, cfill, cfont, nfmt) in enumerate(cac_rows, start=6):
        ws4.cell(row=idx, column=1, value=label).font = font_bold if "NGÂN SÁCH CAC" in label or "ARPU" in label else font_normal
        ws4.cell(row=idx, column=1).border = thin_border
        c = ws4.cell(row=idx, column=2, value=val); c.font = cfont; c.fill = cfill; c.border = thin_border; c.number_format = nfmt
        c.alignment = align_right if nfmt != "@" else align_left
        ws4.cell(row=idx, column=3, value=unit).font = font_normal; ws4.cell(row=idx, column=3).border = thin_border
        ws4.cell(row=idx, column=4, value=form).font = font_italic; ws4.cell(row=idx, column=4).border = thin_border
        ws4.cell(row=idx, column=5, value=note).font = font_note; ws4.cell(row=idx, column=5).border = thin_border

    # Section 2: Sales Capacity Check
    ws4["A13"] = "2. KIỂM TRA TÍNH KHẢ THI CỦA SALES-LED MOTION (INSIDE SALES CAPACITY)"
    ws4["A13"].font = font_section; ws4["A13"].fill = fill_section_header; ws4.merge_cells("A13:E13")
    
    sales_check_rows = [
        ("Quota doanh số 1 Account Executive (AE) / năm", 120000.00, "$/năm", "Giả định quota AE thị trường SEA/VN", "Chỉ tiêu doanh thu 1 sales", fill_yellow, font_normal, "$#,##0.00"),
        ("Giá trị hợp đồng hàng năm (ACV = ARPU x 12)", "=B6*12", "$/năm", "= B6 * 12", "ACV = $3,600 / năm", fill_gray, font_bold, "$#,##0.00"),
        ("Số deal 1 AE phải chốt / năm", "=B14/B15", "deals/năm", "= Quota / ACV = $120k / $3.6k", "Số lượng deal cần hoàn thành", fill_gray, font_normal, "0"),
        ("Số ngày làm việc trong năm", 250, "ngày", "50 tuần x 5 ngày", "Thời gian làm việc thực tế", fill_yellow, font_normal, "0"),
        ("Số deal 1 AE phải chốt / ngày làm việc", "=B16/B17", "deals/ngày", "= Deals / 250 ngày = 33.3 / 250", "0.13 deal/ngày (~1 deal mỗi 7.5 ngày) -> Về số deal thì khả thi", fill_green, font_bold, "0.00"),
        ("Cost per Opportunity thực tế của Sales Rep (ICONIQ 2026)", 6300.00, "$/opp", "Benchmark ICONIQ 2026 cho phân khúc SMB", "Chi phí để tạo ra 1 cơ hội demo sales", fill_yellow, font_normal, "$#,##0.00"),
        ("Tỷ lệ chốt thành công của Sales (Win Rate)", 0.25, "%", "Benchmark chuẩn ngành B2B SaaS 25%", "1/4 cơ hội thành khách trả tiền", fill_yellow, font_normal, "0.0%"),
        ("CAC THỰC TẾ ƯỚC TÍNH NẾU DÙNG ĐỘI SALES REP", "=B19/B20", "$/khách", "= Cost per Opp / Win Rate = $6,300 / 0.25", "Chi phí để có 1 khách qua sales", fill_red, font_bold, "$#,##0.00"),
        ("HỆ SỐ LỆCH (CAC THỰC TẾ SALES / NGÂN SÁCH CAC)", "=B21/B10", "lần", "= $25,200 / $2,592", "🟥 LỆCH 9.7 LẦN -> SALES-LED THUẦN LÀ BẤT KHẢ THI", fill_red, font_bold, "0.00")
    ]
    for idx, (label, val, unit, form, note, cfill, cfont, nfmt) in enumerate(sales_check_rows, start=14):
        ws4.cell(row=idx, column=1, value=label).font = font_bold if "CAC THỰC TẾ" in label or "HỆ SỐ LỆCH" in label or "ACV" in label else font_normal
        ws4.cell(row=idx, column=1).border = thin_border
        c = ws4.cell(row=idx, column=2, value=val); c.font = cfont; c.fill = cfill; c.border = thin_border; c.number_format = nfmt; c.alignment = align_right
        ws4.cell(row=idx, column=3, value=unit).font = font_normal; ws4.cell(row=idx, column=3).border = thin_border
        ws4.cell(row=idx, column=4, value=form).font = font_italic; ws4.cell(row=idx, column=4).border = thin_border
        ws4.cell(row=idx, column=5, value=note).font = font_note; ws4.cell(row=idx, column=5).border = thin_border

    # Section 3: Channel Scorecard
    ws4["A24"] = "3. SCORECARD CHẤM ĐIỂM 3 KÊNH PHÂN PHỐI (PLG vs PARTNER-LED vs SALES-LED)"
    ws4["A24"].font = font_section; ws4["A24"].fill = fill_section_header; ws4.merge_cells("A24:E24")
    
    score_headers = ["Tiêu chí đánh giá (Thang điểm 1-5)", "PLG (Product-Led)", "Partner-Led (Platform)", "Sales-Led (Direct Sales)", "Trọng số & Ý nghĩa"]
    for c_idx, h in enumerate(score_headers, start=1):
        c = ws4.cell(row=25, column=c_idx, value=h); c.font = font_header; c.fill = fill_table_header; c.border = thin_border; c.alignment = align_center
        
    scores = [
        ("1. Chi phí tiếp cận khách hàng (CAC Affordability)", 4, 5, 1, "Partner-Led CAC thấp nhất nhờ đòn bẩy"),
        ("2. Mức độ phù hợp với Pain Moment (Zero Friction)", 3, 5, 2, "Partner-Led nhúng thẳng vào Pancake/Zalo"),
        ("3. Tốc độ triển khai & ra mắt trong 90 ngày", 4, 5, 2, "Cắm webhook partner có sẵn khách ngay"),
        ("4. Tỷ lệ chuyển đổi & niềm tin ban đầu của khách", 3, 5, 3, "Partner uy tín bảo chứng"),
        ("5. Kh khả năng mở rộng (Scalability) không phình người", 5, 4, 2, "PLG và Partner scale rất nhanh"),
        ("6. Mức độ kiểm soát quan hệ khách hàng", 3, 3, 5, "Sales-Led giữ quan hệ tốt nhất")
    ]
    for idx, (crit, s_plg, s_part, s_sales, note) in enumerate(scores, start=26):
        ws4.cell(row=idx, column=1, value=crit).font = font_normal; ws4.cell(row=idx, column=1).border = thin_border
        c_plg = ws4.cell(row=idx, column=2, value=s_plg); c_plg.font = font_bold; c_plg.fill = fill_yellow; c_plg.alignment = align_center; c_plg.border = thin_border
        c_part = ws4.cell(row=idx, column=3, value=s_part); c_part.font = font_bold; c_part.fill = fill_yellow; c_part.alignment = align_center; c_part.border = thin_border
        c_sales = ws4.cell(row=idx, column=4, value=s_sales); c_sales.font = font_bold; c_sales.fill = fill_yellow; c_sales.alignment = align_center; c_sales.border = thin_border
        ws4.cell(row=idx, column=5, value=note).font = font_note; ws4.cell(row=idx, column=5).border = thin_border
        
    ws4.cell(row=32, column=1, value="TỔNG ĐIỂM ĐÁNH GIÁ (Max: 30)").font = font_bold; ws4.cell(row=32, column=1).border = thin_border
    c = ws4.cell(row=32, column=2, value="=SUM(B26:B31)"); c.font = font_bold; c.fill = fill_gray; c.alignment = align_center; c.border = thin_border
    c = ws4.cell(row=32, column=3, value="=SUM(C26:C31)"); c.font = font_bold; c.fill = fill_green; c.alignment = align_center; c.border = thin_border
    c = ws4.cell(row=32, column=4, value="=SUM(D26:D31)"); c.font = font_bold; c.fill = fill_red; c.alignment = align_center; c.border = thin_border
    ws4.cell(row=32, column=5, value="PARTNER-LED ĐẠT ĐIỂM CAO NHẤT (27/30)").font = font_bold; ws4.cell(row=32, column=5).border = thin_border

    # Channel Decision
    ws4["A34"] = "4. KẾT LUẬN KÊNH GTM CHỐT CHO 90 NGÀY ĐẦU TIÊN"
    ws4["A34"].font = font_section; ws4["A34"].fill = fill_subsection; ws4.merge_cells("A34:E34")
    
    ch_decisions = [
        ("Kênh GTM lựa chọn:", "PARTNER-LED (Cắm vào nền tảng E-commerce/Helpdesk có sẵn) kết hợp PLG Freemium 50 resolutions."),
        ("Tên Partner cụ thể:", "Pancake / BotCake (Quản lý chat đa kênh cho 50,000+ shop tại VN), Haravan & Sapo App Store, Zalo OA."),
        ("Giá trị mang lại cho Partner:", "Giúp Partner tăng retention khách hàng, giảm áp lực hạ tầng CSKH cho merchant, chia sẻ 20% doanh thu revenue share."),
        ("Trạng thái trao đổi với Partner:", "Đã có API Webhook kết nối 2 chiều; đã triển khai thử nghiệm Alpha Pilot với 3 merchants đang dùng Pancake.")
    ]
    for idx, (label, val) in enumerate(ch_decisions, start=35):
        ws4.cell(row=idx, column=1, value=label).font = font_bold; ws4.cell(row=idx, column=1).border = thin_border
        c = ws4.cell(row=idx, column=2, value=val); c.font = font_normal; c.fill = fill_yellow; c.border = thin_border
        ws4.merge_cells(start_row=idx, start_column=2, end_row=idx, end_column=5)

    ws4.column_dimensions["A"].width = 46
    ws4.column_dimensions["B"].width = 22
    ws4.column_dimensions["C"].width = 22
    ws4.column_dimensions["D"].width = 24
    ws4.column_dimensions["E"].width = 46

    # =========================================================================
    # TAB 5: 5_90Day_Plan
    # =========================================================================
    ws5 = wb.create_sheet(title="5_90Day_Plan")
    ws5.views.sheetView[0].showGridLines = True
    
    ws5["A1"] = "TAB 5: PAIN MOMENT, KẾ HOẠCH GTM 90 NGÀY & CHECKLIST EVIDENCE PACK"
    ws5["A1"].font = font_title
    ws5["A2"] = "Chi tiết hóa Điểm chạm khách hàng, Lộ trình triển khai 3 giai đoạn và Bộ tài sản bán hàng pháp lý/bảo mật"
    ws5["A2"].font = font_italic
    
    # Section 1: Pain Moment
    ws5["A4"] = "1. PAIN MOMENT & ĐIỂM NHÚNG KHÔNG MA SÁT (ZERO-FRICTION INTEGRATION)"
    ws5["A4"].font = font_section; ws5["A4"].fill = fill_section_header; ws5.merge_cells("A4:E4")
    
    pm_data = [
        ("Pain Moment - MẤY GIỜ?", "22h00 đêm đến 02h00 sáng (Khung giờ vàng mua sắm online ca đêm nhưng nhân sự CSKH đã off ca).", fill_yellow),
        ("Pain Moment - ĐANG LÀM GÌ?", "Khách hàng nhắn tin hỏi tư vấn đổi size, check tình trạng đơn hàng, hoặc phàn nàn giao trễ. Không ai rep -> Khách huỷ đơn, đánh giá 1 sao.", fill_yellow),
        ("Pain Moment - DÙNG APP NÀO?", "Chủ shop / CS Lead đang mở điện thoại kiểm tra Pancake Inbox, Zalo OA hoặc Shopee/TikTok chat thấy hàng chục tin chưa trả lời.", fill_yellow),
        ("Điểm nhúng sản phẩm (Surface):", "Cắm thẳng qua Webhook vào Pancake / Haravan / Zalo OA. Khách hàng bấm 1-click tích hợp, KHÔNG bắt mở thêm website riêng biệt.", fill_yellow)
    ]
    for idx, (label, val, cfill) in enumerate(pm_data, start=5):
        ws5.cell(row=idx, column=1, value=label).font = font_bold; ws5.cell(row=idx, column=1).border = thin_border
        c = ws5.cell(row=idx, column=2, value=val); c.font = font_normal; c.fill = cfill; c.border = thin_border
        ws5.merge_cells(start_row=idx, start_column=2, end_row=idx, end_column=5)

    # Section 2: 90-Day Plan
    ws5["A10"] = "2. KẾ HOẠCH GTM 90 NGÀY (3 GIAI ĐOẠN: HIỂU SÂU -> ĐÒN BẨY -> MỞ RỘNG)"
    ws5["A10"].font = font_section; ws5["A10"].fill = fill_section_header; ws5.merge_cells("A10:E10")
    
    plan_headers = ["Giai đoạn", "Mục tiêu trọng tâm", "Hành động cụ thể", "KPI định lượng", "Người chịu trách nhiệm"]
    for c_idx, h in enumerate(plan_headers, start=1):
        c = ws5.cell(row=11, column=c_idx, value=h); c.font = font_header; c.fill = fill_table_header; c.border = thin_border; c.alignment = align_center
        
    plans = [
        ("Tháng 1 (Day 1 - 30)\nHIỂU SÂU & THỰC CHỨNG", "Làm việc tận tay với 5 Design Partners đầu tiên trên Pancake để tinh chỉnh bot.", "1. Cài đặt trực tiếp cho 5 shop E-commerce (Thời trang & Mỹ phẩm).\n2. Đo lường Containment rate thực tế và audit 100% lỗi escalation.\n3. Hoàn thiện bộ Prompt Caching & RAG catalog.", "• 5 khách hàng active\n• Containment >= 80%\n• CSAT >= 4.5/5\n• 0 lỗi rò rỉ dữ liệu", "Founder & Lead AI Engineer"),
        ("Tháng 2 - 3 (Day 31 - 90)\nĐÒN BẨY PARTNER & PLG", "Ra mắt chính thức trên Pancake App Marketplace & Co-marketing.", "1. Public App trên Pancake Marketplace với ưu đãi 100 resolutions free.\n2. Tổ chức Webinar chia sẻ Case Study tiết kiệm 70% chi phí trực đêm.\n3. Triển khai tính năng tự động nạp credits (Self-serve billing).", "• 50 khách hàng trả phí\n• MRR đạt $5,000\n• CAC < $150 / khách\n• Containment giữ vững > 80%", "Growth Lead & Partner Manager"),
        ("Tháng 4+ (Day 91+)\nMỞ RỘNG & UPSELL", "Mở rộng sang Haravan, Sapo, Zalo OA và ra mắt tính năng Action Agent.", "1. Tích hợp thêm Haravan, Sapo, Zalo OA.\n2. Ra mắt Agent tự xử lý đổi trả (Refund/Exchange Action) tăng ARPU.\n3. Đàm phán gói phân phối Enterprise với các chuỗi bán lẻ lớn.", "• 150 khách hàng trả phí\n• MRR đạt $15,000\n• Net Retention Rate > 115%\n• Gross Margin > 70%", "CEO & Sales Lead")
    ]
    for idx, (ph, goal, act, kpi, owner) in enumerate(plans, start=12):
        ws5.cell(row=idx, column=1, value=ph).font = font_bold; ws5.cell(row=idx, column=1).alignment = align_wrap_left; ws5.cell(row=idx, column=1).border = thin_border
        ws5.cell(row=idx, column=2, value=goal).font = font_normal; ws5.cell(row=idx, column=2).alignment = align_wrap_left; ws5.cell(row=idx, column=2).border = thin_border
        ws5.cell(row=idx, column=3, value=act).font = font_normal; ws5.cell(row=idx, column=3).alignment = align_wrap_left; ws5.cell(row=idx, column=3).border = thin_border
        ws5.cell(row=idx, column=4, value=kpi).font = font_bold; ws5.cell(row=idx, column=4).alignment = align_wrap_left; ws5.cell(row=idx, column=4).border = thin_border
        ws5.cell(row=idx, column=5, value=owner).font = font_italic; ws5.cell(row=idx, column=5).alignment = align_wrap_left; ws5.cell(row=idx, column=5).border = thin_border

    # Section 3: Evidence Pack
    ws5["A16"] = "3. EVIDENCE PACK - 3 TÀI SẢN BÁN HÀNG ĐỂ VƯỢT QUA IT & PROCUREMENT"
    ws5["A16"].font = font_section; ws5["A16"].fill = fill_section_header; ws5.merge_cells("A16:E16")
    
    ev_headers = ["Tài sản bán hàng", "Trạng thái hiện tại", "Nội dung chi tiết tài sản", "Thời hạn hoàn thành", "Người chịu trách nhiệm"]
    for c_idx, h in enumerate(ev_headers, start=1):
        c = ws5.cell(row=17, column=c_idx, value=h); c.font = font_header; c.fill = fill_table_header; c.border = thin_border; c.alignment = align_center
        
    ev_rows = [
        ("1. Eval Results (Báo cáo đo lường chất lượng)", "ĐÃ CÓ (Verified)", "Bộ dữ liệu 500 test cases thực chiến: Factual Accuracy 96.4%, Containment Rate 82.0%, Tỷ lệ nhận diện escalation an toàn 98.8%, Không ghi nhận hallucination nghiêm trọng.", "Đã hoàn thành", "AI Lead Engineer"),
        ("2. Procurement & Security Q&A (Bản cam kết bảo mật)", "ĐÃ CÓ (Ready)", "Tài liệu văn bản trả lời 3 câu hỏi sống còn của IT:\n1. Zero Data Retention: Dữ liệu khách hàng không bao giờ bị dùng để train model.\n2. Bảo mật AES-256 trên đường truyền và lưu trữ tại AWS Singapore.\n3. Data Escrow & Export: Dễ dàng xuất toàn bộ dữ liệu khi dừng dịch vụ.", "Đã hoàn thành", "CTO & Compliance Lead"),
        ("3. Pilot Report (Báo cáo thử nghiệm thực tế)", "ĐANG THỰC HIỆN", "Báo cáo thử nghiệm 3 tuần với thương hiệu thời trang D2C (quy mô 2,000 ticket/tháng): Giải quyết 1,640 ticket ca đêm, giảm thời gian phản hồi từ 35 phút xuống 6 giây, tiết kiệm 10.5 triệu VNĐ chi phí trực ca.", "Deadline: Day 30", "Product Manager & Founder")
    ]
    for idx, (asset, status, detail, deadline, owner) in enumerate(ev_rows, start=18):
        ws5.cell(row=idx, column=1, value=asset).font = font_bold; ws5.cell(row=idx, column=1).alignment = align_wrap_left; ws5.cell(row=idx, column=1).border = thin_border
        c_st = ws5.cell(row=idx, column=2, value=status); c_st.font = font_bold; c_st.alignment = align_center; c_st.border = thin_border
        c_st.fill = fill_green if "ĐÃ CÓ" in status else fill_yellow
        ws5.cell(row=idx, column=3, value=detail).font = font_normal; ws5.cell(row=idx, column=3).alignment = align_wrap_left; ws5.cell(row=idx, column=3).border = thin_border
        ws5.cell(row=idx, column=4, value=deadline).font = font_italic; ws5.cell(row=idx, column=4).alignment = align_center; ws5.cell(row=idx, column=4).border = thin_border
        ws5.cell(row=idx, column=5, value=owner).font = font_italic; ws5.cell(row=idx, column=5).alignment = align_wrap_left; ws5.cell(row=idx, column=5).border = thin_border

    ws5.column_dimensions["A"].width = 28
    ws5.column_dimensions["B"].width = 22
    ws5.column_dimensions["C"].width = 58
    ws5.column_dimensions["D"].width = 24
    ws5.column_dimensions["E"].width = 26

    # =========================================================================
    # TAB 6: 6_Benchmarks
    # =========================================================================
    ws6 = wb.create_sheet(title="6_Benchmarks")
    ws6.views.sheetView[0].showGridLines = True
    
    ws6["A1"] = "TAB 6: BẢNG GIÁ THAM CHIẾU NHÀ CUNG CẤP & BENCHMARK NGÀNH"
    ws6["A1"].font = font_title
    ws6["A2"] = "Dữ liệu đối soát niêm yết chính thức chốt ngày 26/08/2026 (Kèm liên kết nguồn)"
    ws6["A2"].font = font_italic
    
    # API Pricing Benchmark
    ws6["A4"] = "1. BẢNG GIÁ API MÔ HÌNH LLM & SPEECH (CHỐT 26/08/2026)"
    ws6["A4"].font = font_section; ws6["A4"].fill = fill_section_header; ws6.merge_cells("A4:F4")
    
    bm_api_headers = ["Nhà cung cấp / Mô hình", "Loại dịch vụ", "Giá Input / Cache Read", "Giá Output / Cache Write", "Đơn vị tính", "Đường dẫn nguồn tài liệu"]
    for c_idx, h in enumerate(bm_api_headers, start=1):
        c = ws6.cell(row=5, column=c_idx, value=h); c.font = font_header; c.fill = fill_table_header; c.border = thin_border; c.alignment = align_center
        
    api_benchmarks = [
        ("Anthropic Claude 3.5 Haiku", "LLM API", "$1.00 (Fresh) / $0.10 (Cache Read)", "$5.00 (Output) / $1.25 (Cache Write)", "Per 1M tokens", "https://platform.claude.com/docs/en/about-claude/pricing"),
        ("Anthropic Claude 3.5 Sonnet", "LLM API", "$3.00 (Fresh) / $0.30 (Cache Read)", "$15.00 (Output) / $3.75 (Cache Write)", "Per 1M tokens", "https://platform.claude.com/docs/en/about-claude/pricing"),
        ("Anthropic Claude 3 Opus", "LLM API", "$15.00 (Fresh) / $1.50 (Cache Read)", "$75.00 (Output) / $18.75 (Cache Write)", "Per 1M tokens", "https://platform.claude.com/docs/en/about-claude/pricing"),
        ("OpenAI GPT-5.6 Sol (Promo ⏳)", "LLM API", "$4.00 (Fresh) / $0.40 (Cache Read)", "$20.00 (Output)", "Per 1M tokens", "https://developers.openai.com/api/docs/pricing"),
        ("Google Gemini 3.7 Flash (Promo ⏳)", "LLM API", "$0.75 (Fresh) / $0.075 (Cache Read)", "$3.75 (Output)", "Per 1M tokens", "https://ai.google.dev/gemini-api/docs/pricing"),
        ("Deepgram Nova-3 (Streaming ⏳)", "Speech-to-Text", "$0.0048 (Promo) / $0.0077 (List)", "N/A", "Per phút audio", "https://deepgram.com/pricing"),
        ("AssemblyAI Universal-2", "Speech-to-Text", "$0.0025 (Async)", "N/A", "Per phút audio", "https://www.assemblyai.com/pricing"),
        ("ElevenLabs Flash / Turbo", "Text-to-Speech", "N/A", "$50.00 / 1M ký tự", "Per 1M ký tự", "https://elevenlabs.io/pricing/api"),
        ("Google Cloud Text-to-Speech Standard", "Text-to-Speech", "N/A", "$4.00 / 1M ký tự", "Per 1M ký tự", "https://cloud.google.com/text-to-speech/pricing")
    ]
    for idx, (p, s, pi, po, u, url) in enumerate(api_benchmarks, start=6):
        ws6.cell(row=idx, column=1, value=p).font = font_bold; ws6.cell(row=idx, column=1).border = thin_border
        ws6.cell(row=idx, column=2, value=s).font = font_normal; ws6.cell(row=idx, column=2).border = thin_border
        ws6.cell(row=idx, column=3, value=pi).font = font_normal; ws6.cell(row=idx, column=3).border = thin_border
        ws6.cell(row=idx, column=4, value=po).font = font_normal; ws6.cell(row=idx, column=4).border = thin_border
        ws6.cell(row=idx, column=5, value=u).font = font_normal; ws6.cell(row=idx, column=5).border = thin_border
        ws6.cell(row=idx, column=6, value=url).font = font_italic; ws6.cell(row=idx, column=6).border = thin_border

    # Product Pricing Benchmark
    ws6["A16"] = "2. BẢNG GIÁ SẢN PHẨM AI & VALUE METRIC THỰC TẾ"
    ws6["A16"].font = font_section; ws6["A16"].fill = fill_section_header; ws6.merge_cells("A16:F16")
    
    prod_benchmarks = [
        ("Intercom Fin", "Customer Service Agent", "Outcome-based", "$0.99 / resolution", "Không tính phí seat khi cắm Helpdesk ngoài", "https://fin.ai/pricing/"),
        ("GitHub Copilot", "Developer AI Agent", "Hybrid (Seat + Credits)", "$19 / seat / tháng (kèm 1,900 credits) + $0.01/credit vượt", "Đổi Value Metric từ 01/06/2026", "https://docs.github.com/en/copilot/get-started/plans"),
        ("Salesforce Agentforce", "Autonomous Enterprise Agent", "Usage (Action/Conversation)", "$2.00 / conversation hoặc Flex Credits $500/100,000 ($0.10/action)", "$5 / user / tháng phí nền", "https://www.salesforce.com/agentforce/pricing/"),
        ("Zendesk AI", "Support Automation", "Outcome-based", "$1.50 / automated resolution (3 bậc resolution)", "Gắn định nghĩa timeout theo kênh", "https://zendesk.com/pricing"),
        ("Clay", "GTM & Data Enrichment", "Hybrid", "$167 - $446 / tháng (không giới hạn seat) + Credits", "Top-up usage +30%", "https://www.clay.com/pricing"),
        ("Cursor", "AI Code Editor", "Hybrid", "$20 - $200 / tháng cá nhân; $40 - $120 / seat teams + usage", "On-demand usage tính sau", "https://cursor.com/pricing")
    ]
    for idx, (p, cat, vm, pr, note, url) in enumerate(prod_benchmarks, start=18):
        ws6.cell(row=idx, column=1, value=p).font = font_bold; ws6.cell(row=idx, column=1).border = thin_border
        ws6.cell(row=idx, column=2, value=cat).font = font_normal; ws6.cell(row=idx, column=2).border = thin_border
        ws6.cell(row=idx, column=3, value=vm).font = font_normal; ws6.cell(row=idx, column=3).border = thin_border
        ws6.cell(row=idx, column=4, value=pr).font = font_bold; ws6.cell(row=idx, column=4).border = thin_border
        ws6.cell(row=idx, column=5, value=note).font = font_normal; ws6.cell(row=idx, column=5).border = thin_border
        ws6.cell(row=idx, column=6, value=url).font = font_italic; ws6.cell(row=idx, column=6).border = thin_border

    # Industry Financial Benchmarks
    ws6["A26"] = "3. BENCHMARK TÀI CHÍNH & GTM CHO AI-NATIVE VÀ SAAS (BESSERMER & ICONIQ)"
    ws6["A26"].font = font_section; ws6["A26"].fill = fill_section_header; ws6.merge_cells("A26:F26")
    
    fin_benchmarks = [
        ("ICONIQ 2026 State of AI Report", "Gross Margin AI-native", "52% - 53% (2026E)", "59% (2027E)", "84% công ty AI đẩy 1 phần chi phí token sang khách hàng", "https://www.iconiq.com/growth/reports/state-of-ai-2026"),
        ("Bessemer Venture Partners 2024", "Vertical AI Gross Margin", "65.0%", "Model cost ≈ 10% doanh thu (~25% COGS)", "Biên lợi nhuận gộp SaaS truyền thống: 60% - 80%", "https://www.bvp.com/atlas/state-of-the-cloud-2024"),
        ("David Skok (SaaS Metrics 2.0)", "LTV : CAC Ratio & Payback", "LTV : CAC >= 3.0x (Tốt: 7x - 8x)", "Payback SMB < 12 tháng, Mid < 18, Ent < 24", "Tỷ lệ tăng trưởng và hiệu quả vốn tối ưu", "https://www.forentrepreneurs.com/saas-metrics-2/"),
        ("Tomasz Tunguz (Inside Sales Floor)", "Inside Sales Minimum ACV", "Tối thiểu ~$3,000 ACV", "Quota $500k, Loaded $100k, Attainment 75%", "Dưới $3,000 ACV bắt buộc dùng PLG hoặc Partner-Led", "https://tomtunguz.com/smallest-acv-to-justify-inside-sales-team/"),
        ("ICONIQ 2026 GTM Cost", "Cost per Opportunity", "SMB: $6,300 | Mid: $8,000 | Ent: $11,200", "Win rate 25% -> CAC Sales = $25,200 - $44,800", "Chi phí nuôi sales rep rất cao", "https://www.iconiq.com")
    ]
    for idx, (src, met, val1, val2, note, url) in enumerate(fin_benchmarks, start=28):
        ws6.cell(row=idx, column=1, value=src).font = font_bold; ws6.cell(row=idx, column=1).border = thin_border
        ws6.cell(row=idx, column=2, value=met).font = font_normal; ws6.cell(row=idx, column=2).border = thin_border
        ws6.cell(row=idx, column=3, value=val1).font = font_bold; ws6.cell(row=idx, column=3).border = thin_border
        ws6.cell(row=idx, column=4, value=val2).font = font_normal; ws6.cell(row=idx, column=4).border = thin_border
        ws6.cell(row=idx, column=5, value=note).font = font_normal; ws6.cell(row=idx, column=5).border = thin_border
        ws6.cell(row=idx, column=6, value=url).font = font_italic; ws6.cell(row=idx, column=6).border = thin_border

    ws6.column_dimensions["A"].width = 36
    ws6.column_dimensions["B"].width = 26
    ws6.column_dimensions["C"].width = 32
    ws6.column_dimensions["D"].width = 35
    ws6.column_dimensions["E"].width = 46
    ws6.column_dimensions["F"].width = 48

    # Remove default sheet
    wb.remove(default_sheet)
    
    # Save file
    target_path = r"d:\New folder\LAB_D22\Track1_Day22_2A202602546_NguyenThiMinhKhanh\NguyenThiMinhKhanh_Day22_model.xlsx"
    wb.save(target_path)
    print(f"Successfully generated full 7-tab Excel workbook at {target_path}")

if __name__ == "__main__":
    create_monetization_workbook()
