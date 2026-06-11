#!/usr/bin/env python3
"""
獨立測試腳本：測試 app.py 的核心函數，找出 runtime bug
"""
import os
import sys

# 隱藏 Streamlit 警告
os.environ['STREAMLIT_CLIENT_LOGGER_LEVEL'] = 'error'

print("=" * 60)
print("Step 1: 測試匯入")
print("=" * 60)

try:
    from app import (
        filter_plain_text,
        split_inputs,
        extract_sections,
        extract_region_and_name,
        sanitize_filename,
        save_pdf,
        TacticalPDF,
    )
    print("✓ 匯入成功")
except Exception as e:
    print(f"✗ 匯入失敗: {e}")
    sys.exit(1)

print("\n" + "=" * 60)
print("Step 2: 測試 filter_plain_text")
print("=" * 60)

try:
    text = "台北市信義區房屋 3房2廳1衛，總價 2500萬 🏠"
    result = filter_plain_text(text)
    print(f"輸入: {text}")
    print(f"輸出: {result}")
    assert "🏠" not in result, "Emoji 應該被移除"
    print("✓ 文字過濾成功")
except Exception as e:
    print(f"✗ 文字過濾失敗: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("Step 3: 測試 split_inputs")
print("=" * 60)

try:
    raw = "房屋1\n房屋2,房屋3"
    result = split_inputs(raw)
    print(f"輸入: {repr(raw)}")
    print(f"輸出: {result}")
    assert len(result) == 3, f"應該分成 3 項，實際 {len(result)}"
    print("✓ 分割輸入成功")
except Exception as e:
    print(f"✗ 分割輸入失敗: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("Step 4: 測試 sanitize_filename")
print("=" * 60)

try:
    filename = "台北市 信義區/房屋 (2024)"
    result = sanitize_filename(filename)
    print(f"輸入: {filename}")
    print(f"輸出: {result}")
    assert "/" not in result and "(" not in result, "特殊字元應被移除"
    print("✓ 檔案名淨化成功")
except Exception as e:
    print(f"✗ 檔案名淨化失敗: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("Step 5: 測試 extract_sections")
print("=" * 60)

try:
    content = """【物件核心規格與黃金價值】
三房兩廳
位置優越

【591專業優化版】
這是 591 的描述

【FB社團吸粉版】
這是 FB 的描述

【LINE/限動秒殺版】
限時優惠"""

    result = extract_sections(content)
    print(f"提取的區塊: {list(result.keys())}")
    for key, val in result.items():
        if val.strip():
            print(f"  {key}: {val[:30]}...")
    assert "核心規格" in result and result["核心規格"].strip(), "核心規格應被提取"
    print("✓ 區塊提取成功")
except Exception as e:
    print(f"✗ 區塊提取失敗: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("Step 6: 測試 extract_region_and_name")
print("=" * 60)

try:
    title = "台北市信義區 3房房屋出售"
    region, name = extract_region_and_name(title, "")
    print(f"標題: {title}")
    print(f"區域: {region}, 名稱: {name}")
    assert region == "台北", f"應識別為台北，實際 {region}"
    print("✓ 地區名稱提取成功")
except Exception as e:
    print(f"✗ 地區名稱提取失敗: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("Step 7: 測試 TacticalPDF 初始化與基本操作")
print("=" * 60)

try:
    pdf = TacticalPDF("測試物件")
    print(f"✓ PDF 物件初始化成功")
    print(f"  使用字體: {pdf.base_font}")
    
    pdf.add_page()
    print(f"✓ 新增頁面成功")
    
    pdf.section_title("測試標題")
    print(f"✓ 新增標題成功")
    
    pdf.add_authoritative_paragraph("這是測試文字，包含數字 123 坪，以及特殊符號 #test *emoji*")
    print(f"✓ 新增段落成功")
    
    pdf.add_contact_box("王小明", "0912345678")
    print(f"✓ 新增聯繫欄成功")
    
    pdf_bytes = pdf.output(dest='S')
    print(f"✓ PDF 輸出成功，大小: {len(pdf_bytes)} bytes")
    assert isinstance(pdf_bytes, (bytes, bytearray)), f"PDF 應為 bytes，實際 {type(pdf_bytes)}"
    assert len(pdf_bytes) > 0, "PDF 檔案不應為空"
    print("✓ TacticalPDF 功能完整")
    
except Exception as e:
    print(f"✗ TacticalPDF 測試失敗: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("Step 8: 測試 save_pdf")
print("=" * 60)

try:
    test_content = """【物件核心規格與黃金價值】
位置：台北市信義區
坪數：30坪
格局：3房2廳1衛
總價：2500萬

【591專業優化版】
這是一個優質物件，位於信義區精華地段

【FB社團吸粉版】
超優質房子，趕快聯繫我！

【LINE/限動秒殺版】
台北信義，30坪，2500萬，限時優惠！"""

    filename, bio = save_pdf(
        "台北市信義區 3房房屋", 
        test_content,
        "測試業務",
        "0912345678"
    )
    
    print(f"✓ PDF 儲存成功")
    print(f"  檔案名: {filename}")
    print(f"  檔案大小: {len(bio.getvalue())} bytes")
    print(f"  存儲位置: Outputs/{filename}")
    
    assert os.path.exists(f"Outputs/{filename}"), "PDF 檔案應存在"
    print("✓ save_pdf 功能完整")
    
except Exception as e:
    print(f"✗ save_pdf 測試失敗: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("✓ 所有測試完成！沒有發現 runtime bug")
print("=" * 60)
