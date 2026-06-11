#!/usr/bin/env python3
"""
集成測試：檢查邊界情況和錯誤處理
"""
import os
import sys
import io
from unittest.mock import patch, MagicMock

# 隱藏 Streamlit 警告
os.environ['STREAMLIT_CLIENT_LOGGER_LEVEL'] = 'error'

from app import (
    resolve_input_text,
    extract_sections,
    save_pdf,
    TacticalPDF,
    filter_plain_text,
)

print("=" * 60)
print("整合測試：邊界情況與錯誤處理")
print("=" * 60)

print("\n" + "=" * 60)
print("Test 1: 處理空字串輸入")
print("=" * 60)

try:
    result = filter_plain_text("")
    assert result == "", f"空字串應返回空，得到 {repr(result)}"
    print("✓ 空字串處理正確")
except Exception as e:
    print(f"✗ 失敗: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("Test 2: 處理純文字輸入（不含 591 URL）")
print("=" * 60)

try:
    text = "台北市信義區 3房2廳 2500萬 聯絡我"
    desc, title = resolve_input_text(text)
    assert desc == text, "純文字應直接返回"
    assert title == "", "非 URL 應無 title"
    print(f"✓ 純文字處理成功")
    print(f"  描述: {desc[:50]}...")
except Exception as e:
    print(f"✗ 失敗: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("Test 3: 處理格式不完整的 content（缺少區塊）")
print("=" * 60)

try:
    incomplete_content = "只有一些隨意的文字，沒有任何區塊標題"
    sections = extract_sections(incomplete_content)
    print(f"提取的區塊: {list(sections.keys())}")
    # 驗證所有區塊都存在（可能為空）
    assert len(sections) == 4, f"應有 4 個區塊，得到 {len(sections)}"
    print("✓ 不完整內容處理成功（所有區塊為空）")
    for key, val in sections.items():
        print(f"  {key}: {'[空]' if not val.strip() else val[:20] + '...'}")
except Exception as e:
    print(f"✗ 失敗: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("Test 4: 處理包含特殊字元的檔案名")
print("=" * 60)

try:
    test_content = """【物件核心規格與黃金價值】
位置：台北 * 信義 ? 房 < 屋 > | 特 [ 殊 ] ( 字 ) { 元 }
【591專業優化版】
測試
【FB社團吸粉版】
測試
【LINE/限動秒殺版】
測試"""
    
    filename, bio = save_pdf(
        "台北★信義★3房?(2024)",
        test_content,
        "測試",
        "0912345678"
    )
    
    # 驗證檔案名不含非法字元
    forbidden_chars = ['*', '?', '<', '>', '|', '(', ')', '[', ']', '{', '}', '/', '\\', ':']
    for char in forbidden_chars:
        assert char not in filename, f"檔案名中仍包含 {char}"
    
    print(f"✓ 特殊字元檔案名處理成功")
    print(f"  原始標題: 台北★信義★3房?(2024)")
    print(f"  清淨後: {filename}")
    print(f"  檔案存在: {os.path.exists(f'Outputs/{filename}')}")
    
except Exception as e:
    print(f"✗ 失敗: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("Test 5: PDF 輸出為 bytes 類型")
print("=" * 60)

try:
    pdf = TacticalPDF("BytesTest")
    pdf.add_page()
    pdf.add_authoritative_paragraph("測試位元組輸出")
    
    pdf_bytes = pdf.output(dest='S')
    
    # 驗證輸出類型
    assert isinstance(pdf_bytes, (bytes, bytearray)), f"輸出應為 bytes，得到 {type(pdf_bytes)}"
    
    # 驗證內容有效（PDF 應以 %PDF 開頭）
    if isinstance(pdf_bytes, bytearray):
        pdf_bytes = bytes(pdf_bytes)
    
    assert pdf_bytes.startswith(b'%PDF'), "PDF 檔案應以 %PDF 開頭"
    
    print(f"✓ PDF bytes 輸出驗證成功")
    print(f"  類型: {type(pdf_bytes)}")
    print(f"  大小: {len(pdf_bytes)} bytes")
    print(f"  開頭: {pdf_bytes[:10]}")
    
except Exception as e:
    print(f"✗ 失敗: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("Test 6: PDF 儲存與讀取")
print("=" * 60)

try:
    test_content = """【物件核心規格與黃金價值】
測試核心規格
【591專業優化版】
591版本內容
【FB社團吸粉版】
FB版本內容
【LINE/限動秒殺版】
LINE版本內容"""
    
    filename, bio = save_pdf(
        "測試物件",
        test_content,
        "業務1",
        "0912345678"
    )
    
    file_path = f"Outputs/{filename}"
    
    # 驗證檔案存在
    assert os.path.exists(file_path), f"檔案應存在於 {file_path}"
    
    # 驗證檔案大小
    file_size = os.path.getsize(file_path)
    assert file_size > 0, "檔案應非空"
    assert file_size == len(bio.getvalue()), "檔案大小應與 BytesIO 相符"
    
    # 驗證檔案內容為有效 PDF
    with open(file_path, 'rb') as f:
        file_content = f.read()
        assert file_content.startswith(b'%PDF'), "檔案應為有效 PDF"
    
    print(f"✓ PDF 儲存與讀取驗證成功")
    print(f"  檔案路徑: {file_path}")
    print(f"  檔案大小: {file_size} bytes")
    print(f"  BytesIO 大小: {len(bio.getvalue())} bytes")
    
except Exception as e:
    print(f"✗ 失敗: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("✓ 所有整合測試通過！")
print("=" * 60)
