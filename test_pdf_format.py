#!/usr/bin/env python3
"""
測試 save_pdf 對符合範例格式的內容的處理
"""
import os
os.environ['STREAMLIT_CLIENT_LOGGER_LEVEL'] = 'error'

from app import save_pdf, extract_sections

print("=" * 60)
print("測試 save_pdf 內容呈現")
print("=" * 60)

# 模擬 AI 生成的內容（基於 PDF 範例格式）
test_content = """【物件核心規格與黃金價值】
精準地段：台北市內湖區文湖街21巷118弄12號，坐落於西湖生活圈。
物理規格：地上 7 層、地下 2 層之獨棟華廈。基地面積 132坪，建蔽率 45%，公設比 43%。
戶數規劃：全棟僅 1 棟，共 16戶住家，出入單純。

【591專業優化版】
標題：文湖寶翠西湖生活圈全新華廈｜倒數最後2戶 39坪2+1房2衛｜隨時交屋
內文：
台北市內湖區文湖街精華地段，博泓建設、豐田營造實力打造文湖寶翠。本案為地上 7 層、地下 2 層的獨棟華廈，全社區僅極致單純的 16戶住家。
目前全案銷售已進入最後階段，其餘產品皆已售罄，現正釋出倒數最後兩戶：2+1房 2衛 39坪產品。

【FB社團吸粉版】
內湖西湖生活圈全新完工，全棟16戶純住，最後2戶直接對決！
別再看那些看得到摸不到的預售屋了！
博泓建設文湖寶翠全新成屋，隨時可以交屋！

【LINE/限動秒殺版】
內湖文湖寶翠全新完工，倒數最後 2戶！文湖街 21 巷 118 弄 12 號，開價 68~78萬/坪。僅剩 39坪 2+1房 2衛，機械車位 180萬。"""

print("\n第一步: 測試 extract_sections")
print("-" * 60)

sections = extract_sections(test_content)
for key in ["核心規格", "591版", "FB版", "LINE版"]:
    content = sections[key].strip()
    print(f"\n{key}:")
    if content:
        print(f"  ✓ 成功提取 ({len(content)} 字元)")
        print(f"  預覽: {content[:80]}...")
    else:
        print(f"  ✗ 為空")

print("\n" + "=" * 60)
print("第二步: 測試 save_pdf 呈現")
print("=" * 60)

try:
    filename, bio = save_pdf(
        "文湖寶翠 - 台北市內湖區",
        test_content,
        "吳先生",
        "0928345567"
    )
    
    print(f"\n✓ PDF 生成成功")
    print(f"  檔案名: {filename}")
    print(f"  檔案大小: {len(bio.getvalue())} bytes")
    print(f"  存儲位置: Outputs/{filename}")
    
    # 驗證檔案是否存在
    file_path = f"Outputs/{filename}"
    if os.path.exists(file_path):
        print(f"  ✓ 檔案已存儲")
        file_size = os.path.getsize(file_path)
        print(f"  實際大小: {file_size} bytes")
    
except Exception as e:
    print(f"\n✗ PDF 生成失敗: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("測試完成")
print("=" * 60)
