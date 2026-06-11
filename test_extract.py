#!/usr/bin/env python3
"""
驗證改進的 extract_sections 函數
"""
import os
os.environ['STREAMLIT_CLIENT_LOGGER_LEVEL'] = 'error'

from app import extract_sections

print("=" * 60)
print("測試改進的 extract_sections 函數")
print("=" * 60)

test_cases = [
    {
        "name": "標題和內容在同一行",
        "input": "【物件核心規格與黃金價值】3房2廳1衛，位置優越，格局方正",
        "expected_key": "核心規格",
    },
    {
        "name": "標題獨佔一行，內容在下一行",
        "input": """【591專業優化版】
這是一個優質物件，位於信義區精華地段
面積寬敞，採光充足""",
        "expected_key": "591版",
    },
    {
        "name": "多個區塊混合",
        "input": """【物件核心規格與黃金價值】3房2廳
位置優越
【591專業優化版】591 描述內容
【FB社團吸粉版】FB 吸粉文案
【LINE/限動秒殺版】限時優惠""",
        "expected_keys": ["核心規格", "591版", "FB版", "LINE版"],
    },
]

for test_case in test_cases:
    print("\n" + "=" * 60)
    print(f"Test: {test_case['name']}")
    print("=" * 60)
    
    try:
        result = extract_sections(test_case["input"])
        
        if "expected_key" in test_case:
            key = test_case["expected_key"]
            content = result[key].strip()
            if content:
                print(f"✓ {key} 提取成功")
                print(f"  內容: {content[:60]}...")
            else:
                print(f"✗ {key} 未能提取內容")
                print(f"  輸入: {test_case['input'][:60]}...")
        
        if "expected_keys" in test_case:
            all_success = True
            for key in test_case["expected_keys"]:
                content = result[key].strip()
                status = "✓" if content else "✗"
                print(f"  {status} {key}: {content[:40] if content else '[空]'}...")
                if not content:
                    all_success = False
            
            if all_success:
                print("✓ 所有區塊提取成功")
    
    except Exception as e:
        print(f"✗ 測試失敗: {e}")
        import traceback
        traceback.print_exc()

print("\n" + "=" * 60)
print("✓ extract_sections 驗證完成")
print("=" * 60)
