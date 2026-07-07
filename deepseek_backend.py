import os
import re
import sys
import time
from datetime import date
import requests
import urllib3
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from google import genai

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
time.sleep(2)


def build_prompt(property_description: str) -> str:
    return (
        f"你是一位高成交導向的房仲銷售顧問。請根據以下物件資訊，嚴格依照此六段式骨架撰寫文案，禁止廢話：\n\n"
        "一句抓眼球的標題：要包含地標或最大優勢，如「[行政區]稀缺低總價，首購首選」。\n\n"
        "一句最強差異：直接點出這間房跟同區其他物件不一樣的特點。\n\n"
        "三個買方在意的理由：針對物件優勢條列，如「1. 步行5分鐘進捷運；2. 永久景觀棟距；3. 屋主誠售可議」。\n\n"
        "一段生活場景：用描寫式文字寫出居住感，例如「下班回家下樓就有公園，回家不用找車位」。\n\n"
        "一段完整規格：必須清楚列出「坪數、格局、樓層、管理費、車位類型」，若資訊不足請標示「未提供」。\n\n"
        "一個明確邀請：結尾強制寫「本週僅釋出 3 組帶看名額，歡迎私訊看照片並預約賞屋時段」。\n\n"
        "【執行規則】\n\n"
        "1. 文案要口語化、精簡，絕對不要出現「格局方正、採光佳」這種無效廢話。\n\n"
        "2. 【硬性禁止】絕對禁止輸出任何內部筆記、狀態提示、檢查備註（例如：數據需人工確認、等待核對、讀取中、文案生成中）。\n\n"
        "3. 【誠實原則】若資料缺失，請直接標示「未提供」即可，禁止輸出任何關於系統或數據狀態的描述。\n\n"
        "4. 【直接產出】你的輸出必須直接作為最終交付給客戶的內容，不要包含對系統的對話或備註。\n\n"
        f"待處理資訊：{property_description}"
    )


def resolve_input_text(user_input: str) -> tuple[str, str]:
    if "591.com.tw" not in user_input:
        return user_input, ""
    print("正在解析網址...")
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0 Safari/537.36",
    }
    response = requests.get(user_input, headers=headers, timeout=15, verify=False)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")
    title = soup.title.get_text(strip=True) if soup.title else ""
    text = soup.get_text(separator=" ", strip=True)
    return text[:2000], title


def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r"[\\/:*?\"<>|]", "_", value)
    cleaned = re.sub(r"\s+", "_", cleaned).strip("_")
    return cleaned or "物件"


def split_inputs(raw_input: str) -> list[str]:
    parts = re.split(r"[,\n]+", raw_input)
    return [item.strip() for item in parts if item.strip()]


def extract_sections(content: str) -> dict[str, str]:
    sections = {
        "【591 專業版】": "",
        "【FB 社團吸粉版】": "",
        "【LINE/限動秒殺版】": "",
    }
    current_key = None
    for line in content.splitlines():
        line_strip = line.strip()
        if line_strip in sections:
            current_key = line_strip
            continue
        if current_key:
            sections[current_key] += line + "\n"
    return sections


def filter_generic_landmarks(lines: list[str]) -> list[str]:
    blocked = ["便利商店", "加油站"]
    filtered: list[str] = []
    for line in lines:
        if any(keyword in line for keyword in blocked):
            continue
        filtered.append(line)
    return filtered


def limit_lines(lines: list[str], max_lines: int) -> list[str]:
    return lines[:max_lines]


def extract_intel_section(content: str) -> str:
    # adapt to new client-facing section title
    start_marker = "物件規格解析"
    start_index = content.find(start_marker)
    if start_index == -1:
        return content
    # return from that marker to next major section if exists
    return content[start_index:]


def clean_markdown(text: str) -> str:
    # remove repeated markdown markers and collapse blank lines
    text = re.sub(r"\*+", "", text)
    text = re.sub(r"\n\s*\n+", "\n\n", text)
    return text


def remove_internal_prompts(text: str) -> str:
    """
    Remove all internal verification prompts and check phrases using strict regex patterns.
    These are not meant for client delivery.
    
    Uses fuzzy regex matching to catch variations of status messages.
    """
    # 定義更嚴格的關鍵字組合與正則表達式
    banned_patterns = [
        r"\(.*?數據.*?確認.*?\)",      # 匹配 (數據需人工確認) 及變體
        r"\[.*?數據.*?確認.*?\]",      # 匹配 [數據需人工確認] 及變體
        r"【.*?數據.*?確認.*?】",      # 匹配 【數據需人工確認】 及變體
        r"數據.*?確認",                # 匹配 數據需人工確認、數據待確認等
        r"文案生成中",                 # 匹配 文案生成中
        r"真實物理數據讀取中",         # 匹配 真實物理數據讀取中
        r"等待.*?核對",                # 匹配 等待人工核對、等待核對等
        r"待\s*AI\s*解析",             # 匹配 待 AI 解析
        r"⚠️\s*數據.*?確認",           # 匹配 ⚠️ 數據需人工確認
        r"需人工確認",                 # 匹配 需人工確認
        r"人工確認",                   # 匹配 人工確認
        r"第\s*\d+\s*頁.*",            # 匹配 第 1 頁... 及其他頁碼
        r"數據.*?讀取中",              # 匹配 數據讀取中等
        r"確認中",                     # 匹配 確認中
        r"生成中",                     # 匹配 生成中
    ]
    
    for pattern in banned_patterns:
        text = re.sub(pattern, "", text)
    
    # 最後刪除空行，確保排版乾淨
    text = re.sub(r"\n\s*\n+", "\n\n", text)
    return text.strip()


def add_field(paragraph, instruction: str) -> None:
    # intentionally retained but NOT used: do not insert page fields into output
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), instruction)
    paragraph._p.append(field)


def add_paragraph_with_highlight(document: Document, text: str) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.line_spacing = 1.5
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pattern = re.compile(r"(總價[^，。\n]*|單價[^，。\n]*|完工日期[^，。\n]*)")
    last_index = 0
    for match in pattern.finditer(text):
        start, end = match.span()
        if start > last_index:
            paragraph.add_run(text[last_index:start])
        highlighted = paragraph.add_run(text[start:end])
        highlighted.font.highlight_color = WD_COLOR_INDEX.YELLOW
        last_index = end
    if last_index < len(text):
        paragraph.add_run(text[last_index:])


def apply_styles(document: Document) -> None:
    normal = document.styles["Normal"]
    normal.font.name = "Microsoft JhengHei"
    normal.font.size = Pt(11)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft JhengHei")
    normal._element.rPr.rFonts.set(qn("w:ascii"), "Arial")
    heading1 = document.styles["Heading 1"]
    heading1.font.name = "Microsoft JhengHei"
    heading1.font.size = Pt(18)
    heading1.font.bold = True
    heading1.font.color.rgb = RGBColor(0, 51, 102)
    heading1._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft JhengHei")
    heading1._element.rPr.rFonts.set(qn("w:ascii"), "Arial")
    heading2 = document.styles["Heading 2"]
    heading2.font.name = "Microsoft JhengHei"
    heading2.font.size = Pt(14)
    heading2.font.bold = True
    heading2.font.color.rgb = RGBColor(0, 0, 0)
    heading2._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft JhengHei")
    heading2._element.rPr.rFonts.set(qn("w:ascii"), "Arial")


def get_contact_info() -> tuple[str, str]:
    """
    Get contact information from environment variables.
    If not available, return default fallback contact info.
    """
    contact_name = os.getenv("CONTACT_NAME", "服務專員").strip()
    contact_phone = os.getenv("CONTACT_PHONE", "").strip()
    
    # If phone is not set, use fallback message
    if not contact_phone:
        contact_phone = "歡迎私訊洽詢"
    
    return contact_name, contact_phone


def set_header_footer(document: Document, property_name: str) -> None:
    # Minimal header (property name and date). No page numbers, no warnings.
    section = document.sections[0]
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    header = section.header
    header_table = header.add_table(rows=1, cols=2, width=Inches(6.5))
    header_table.autofit = True
    left_cell = header_table.cell(0, 0)
    right_cell = header_table.cell(0, 1)
    left_cell.text = f"{property_name}"
    right_cell.text = date.today().strftime("%Y/%m/%d")
    right_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Footer: include professional contact info for client delivery
    footer = section.footer
    footer_paragraph = footer.paragraphs[0]
    footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    contact_name, contact_phone = get_contact_info()
    footer_paragraph.text = f"業務聯繫人：{contact_name} ｜ 電話：{contact_phone}"


def extract_region_and_name(title: str, fallback: str) -> tuple[str, str]:
    region_candidates = [
        "台北市", "新北市", "桃園市", "台中市", "台南市", "高雄市", "基隆市", "新竹市", "嘉義市",
        "台北", "新北", "桃園", "台中", "台南", "高雄", "基隆", "新竹", "嘉義",
        "宜蘭", "花蓮", "台東", "雲林", "彰化", "南投", "屏東", "苗栗",
    ]
    region = "全台"
    for candidate in region_candidates:
        if candidate in title:
            region = candidate
            break
    name_source = title or fallback
    name_source = re.sub(r"591|租|售|買|房屋|物件|出售", "", name_source)
    name_source = re.sub(r"[\[\]\(\)【】]", "", name_source)
    name = sanitize_filename(name_source)[:12] or "物件"
    return region, name


def generate_listing(client: genai.Client, property_description: str) -> str:
    prompt = build_prompt(property_description)
    model_sequence = [
        "models/gemini-2.5-flash",
        "models/gemini-2.0-flash",
        "models/gemini-1.5-flash",
        "models/gemma-3-12b",
    ]
    last_error: Exception | None = None
    for model_name in model_sequence:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config={"temperature": 0.7},
            )
            content = getattr(response, "text", None)
            if content:
                return content.strip()
            last_error = ValueError("Gemini response content is empty")
        except Exception as error:
            last_error = error
            continue
    if last_error:
        raise last_error
    raise ValueError("Gemini response content is empty")


def save_docx(title: str, content: str) -> str:
    # Remove all internal verification prompts first
    content = remove_internal_prompts(content)
    
    # Collapse multiple blank lines
    content = re.sub(r"\n\s*\n+", "\n\n", content)

    output_dir = os.path.join(os.getcwd(), "Outputs")
    os.makedirs(output_dir, exist_ok=True)
    
    cleaned_content = clean_markdown(content)
    # Ensure no internal prompts in cleaned content
    cleaned_content = remove_internal_prompts(cleaned_content)
    
    # Prefer client-facing section title
    intel_section = clean_markdown(extract_intel_section(cleaned_content))
    intel_section = remove_internal_prompts(intel_section)
    
    sections = extract_sections(cleaned_content)
    
    document = Document()
    apply_styles(document)
    region, property_name = extract_region_and_name(title, intel_section[:12])
    set_header_footer(document, property_name)

    # Ensure no internal prompts exist in header/footer
    for section in document.sections:
        for p in section.header.paragraphs:
            for run in list(p.runs):
                run.text = remove_internal_prompts(run.text or "")
        for p in section.footer.paragraphs:
            for run in list(p.runs):
                run.text = remove_internal_prompts(run.text or "")

    # Final cleanup of content
    cleaned_content = remove_internal_prompts(cleaned_content)
    cleaned_content = re.sub(r"\n\s*\n+", "\n\n", cleaned_content)
    intel_section = clean_markdown(extract_intel_section(cleaned_content))
    intel_section = remove_internal_prompts(intel_section)
    sections = extract_sections(cleaned_content)

    # Client-facing heading
    document.add_heading("物件規格解析", level=1)
    document.add_paragraph("#物件規格 #市場解析 #銷售要點")
    table = document.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "本案單價"
    table.cell(0, 1).text = "區域行情"
    table.cell(1, 0).text = "價差判讀"
    table.cell(1, 1).text = "詳見下方說明"
    intel_lines = limit_lines(filter_generic_landmarks(intel_section.splitlines()), 12)
    for line in intel_lines:
        line_strip = line.strip()
        if line_strip:
            add_paragraph_with_highlight(document, line_strip)
    document.add_page_break()

    # 591 專業版
    document.add_heading("591 專業版", level=1)
    document.add_paragraph("#成交戰術 #數據精準 #專業建議")
    professional_content = sections["【591 專業版】"].strip() or cleaned_content
    professional_content = remove_internal_prompts(professional_content)
    professional_lines = limit_lines(filter_generic_landmarks(professional_content.splitlines()), 16)
    for line in professional_lines:
        line_strip = line.strip()
        if line_strip:
            add_paragraph_with_highlight(document, line_strip)
    document.add_page_break()

    # FB 社團吸粉版
    document.add_heading("FB 社團吸粉版", level=1)
    document.add_paragraph("#在地社群 #吸粉曝光 #熱區生活")
    fb_content = sections["【FB 社團吸粉版】"].strip()
    fb_content = remove_internal_prompts(fb_content)
    fb_lines = limit_lines(filter_generic_landmarks(fb_content.splitlines()), 10)
    for line in fb_lines:
        line_strip = line.strip()
        if line_strip:
            add_paragraph_with_highlight(document, line_strip)
    document.add_page_break()

    # LINE / 限動版
    document.add_heading("LINE/限動秒殺版", level=1)
    document.add_paragraph("#VIP急售 #限量釋出 #稀缺搶手")
    line_content = sections["【LINE/限動秒殺版】"].strip()
    line_content = remove_internal_prompts(line_content)
    line_lines = limit_lines(filter_generic_landmarks(line_content.splitlines()), 6)
    for line in line_lines:
        line_strip = line.strip()
        if line_strip:
            add_paragraph_with_highlight(document, line_strip)

    # Contact block at end with automated fallback
    contact_para = document.add_paragraph()
    contact_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    contact_name, contact_phone = get_contact_info()
    contact_text = f"業務聯繫人：{contact_name}　|　電話：{contact_phone}"
    contact_run = contact_para.add_run(contact_text)
    contact_run.bold = True

    filename = f"【成品報告】{region}_{property_name}_{date.today().strftime('%m%d')}.docx"
    file_path = os.path.join(output_dir, filename)
    document.save(file_path)
    return file_path


def main() -> None:
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise SystemExit("請設定 GEMINI_API_KEY")
    client = genai.Client(api_key=api_key)
    user_input = input("請直接貼上 591 網址或房屋資訊：")
    inputs = split_inputs(user_input)
    if not inputs:
        raise SystemExit("請提供房地產物件描述文字")
    for item in inputs:
        description, title = resolve_input_text(item)
        if not description:
            continue
        output = generate_listing(client, description)
        base_title = title or (description[:12] if description else "物件")
        save_docx(base_title, output)
        print("✅ 成品報告已生成在 Outputs 資料夾（可直接傳給客戶）")


if __name__ == "__main__":
    main()
