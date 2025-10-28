import os

# 使用現有的辦公專案路徑
current_dir = os.path.dirname(os.path.abspath(__file__))  # 獲取當前檔案所在目錄
photo_module_path = os.path.join(current_dir, "photo_to_word")

try:
    # 建立新的模組目錄
    os.makedirs(photo_module_path, exist_ok=True)
    print(f"建立資料夾: {photo_module_path}")

    # 建立 __init__.py
    init_path = os.path.join(photo_module_path, "__init__.py")
    with open(init_path, 'w', encoding='utf-8') as f:
        pass
    print(f"建立檔案: {init_path}")

    print("\n設定完成！")
    print(f"請將 main.py 複製到: {photo_module_path}")

except Exception as e:
    print(f"發生錯誤: {str(e)}")

input("\n按 Enter 鍵結束...")  # 等待使用者按 Enter
