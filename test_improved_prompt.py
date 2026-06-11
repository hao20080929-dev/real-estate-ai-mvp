#!/usr/bin/env python3
"""
測試改進的 Prompt 生成內容的格式驗證
"""
import os
os.environ['STREAMLIT_CLIENT_LOGGER_LEVEL'] = 'error'

from app import build_prompt, extract_sections, save_pdf

print("=" * 60)
print("測試改進的 Prompt")
print("=" * 60)

# 測試輸入
property_info = "台北市內湖區文湖街21巷118弄12號，7樓華廈，39坪2房2廳，開價2688萬，車位180萬"

print("\n第一步: 查看改進後的 Prompt")
print("-" * 60)

prompt = build_prompt(property_info)
print(prompt[:500])
print("\n... [Prompt 內容已截斷] ...\n")

# 驗證 Prompt 中是否包含重要的結構化指導
required_elements = [
    "【物件核心規格與黃金價值】",
    "【591 專業優化版】",
    "標題：",
    "內文：",
    "【FB 社團吸粉版】",
    "【LINE/限動秒殺版】",
    "不超過 100 字"
]

print("Prompt 檢查清單：")
for element in required_elements:
    if element in prompt:
        print(f"  ✓ 包含: {element}")
    else:
        print(f"  ✗ 缺少: {element}")

# 模擬 AI 會生成的內容（基於改進的 Prompt 指導）
print("\n" + "=" * 60)
print("第二步: 模擬改進後 Prompt 應生成的內容結構")
print("=" * 60)

simulated_ai_output = """【物件核心規格與黃金價值】
- 精準地段：台北市內湖區文湖街21巷118弄12號
- 物理規格：地上7層華廈，39坪
- 格局：2房2廳1衛
- 定價：開價2688萬
- 車位：180萬

【591 專業優化版】
標題：內湖文湖街39坪2房精品華廈｜開價2688萬｜隨時交屋
內文：
台北市內湖區文湖街精華地段，本案為 7 樓華廈。
目前推出 39 坪 2 房 2 廳 1 衛產品，格局方正實用。
開價 2688 萬，車位另計 180 萬。
近捷運、生活機能完善。

【FB 社團吸粉版】
內湖文湖街全新華廈！39坪2房只要2688萬！
別再租房了，買房比你想的更容易！
格局方正、採光充足、馬上可以入住
快來看實品屋！

【LINE/限動秒殺版】
內湖文湖街39坪2房華廈，開價2688萬。地點優越，格局方正。車位180萬。即刻交屋！"""

print("提取的內容：")
sections = extract_sections(simulated_ai_output)
for key in ["核心規格", "591版", "FB版", "LINE版"]:
    content = sections[key].strip()
    if content:
        print(f"\n{key}: ✓ 提取成功 ({len(content)} 字元)")
        lines = content.split('\n')
        for line in lines[:2]:
            if line.strip():
                print(f"  > {line.strip()[:60]}...")
    else:
        print(f"\n{key}: ✗ 未提取到內容")

print("\n" + "=" * 60)
print("第三步: 生成 PDF 驗證")
print("=" * 60)

try:
    filename, bio = save_pdf(
        "內湖文湖街華廈",
        simulated_ai_output,
        "業務員",
        "0912345678"
    )
    print(f"✓ PDF 生成成功")
    print(f"  檔案: {filename}")
    print(f"  大小: {len(bio.getvalue())} bytes")
except Exception as e:
    print(f"✗ PDF 生成失敗: {e}")

print("\n" + "=" * 60)
print("改進 Prompt 總結")
print("=" * 60)
print("""
✓ 改進內容：
1. 明確指定 4 個區塊的結構（用【】符號）
2. 591 版明確要求「標題：」和「內文：」兩部分
3. 強調不省略任何區塊
4. 重複提醒數據 100% 基於輸入資訊
5. 提供具體的內容格式指導

預期效果：
- AI 會更準確地生成符合預期的結構化內容
- extract_sections 會正確識別和提取各區塊
- PDF 會正確呈現完整的文案內容

下一步：
將改進的 Prompt 應用於實際應用程式
用戶上傳房屋信息後，AI 應生成符合預期格式的內容
""")

print("=" * 60)
