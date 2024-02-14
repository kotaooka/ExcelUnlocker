import msoffcrypto
import pandas as pd
import io
import itertools
import tkinter as tk
from tkinter import filedialog
from concurrent.futures import ThreadPoolExecutor
import os

print("Excelファイルのパスワードを解析します")
print("このコードを使用する際は、必ず法律と倫理を遵守してください")

def open_password_protected_excel(file_path, password):
    # ファイルを開く
    file = msoffcrypto.OfficeFile(open(file_path, "rb"))

    # パスワードを設定
    file.load_key(password=password)

    # データをデコード
    decrypted = io.BytesIO()
    file.decrypt(decrypted)

    return decrypted.getvalue()  # バイト列を返す

def try_passwords(passwords, file_path):
    for password in passwords:
        print(f"現在試行しているパスワード: {password}")
        try:
            decrypted_data = open_password_protected_excel(file_path, password)
            print(f"解析したパスワード: {password}")
            
            # 解除されたExcelファイルを別名で保存
            file_name, file_extension = os.path.splitext(file_path)
            new_file_path = f"{file_name}_unlocked{file_extension}"
            with open(new_file_path, "wb") as f:
                f.write(decrypted_data)
            
            print(f"パスワードを解除して別名で保存しました: {new_file_path}")
            return password
        except Exception as e:
            continue
    return None

# 一般的なパスワードのリストをテキストファイルから読み込む
if os.path.exists('password.txt'):
    with open('password.txt', 'r', encoding='utf-8') as f:
        common_passwords = [line.strip() for line in f]
else:
    print("'password.txt'が存在しません。")
    common_passwords = []

# パスワードの組み合わせを生成
chars = '0123456789abcdefghijklmnopqrstuvwxyz'

# ダイアログボックスでファイルを選択
root = tk.Tk()
root.withdraw()  # ダイアログボックスのみ表示
file_path = filedialog.askopenfilename()

# パスワードがかかっていない場合は表示
try:
    pd.read_excel(file_path)
    print("このExcelファイルは暗号化されていません")
except Exception as e:
    with ThreadPoolExecutor(max_workers=10) as executor:
        # 一般的なパスワードを試す
        results = executor.map(lambda x: try_passwords(x, file_path), [common_passwords])
        for password in results:
            if password:
                break
        else:
            # 一般的なパスワードが見つからなかった場合、半角英数字1から8桁までのパスワードを試す
            for i in range(1, 9):
                passwords = [''.join(p) for p in itertools.product(chars, repeat=i)]
                results = executor.map(lambda x: try_passwords(x, file_path), [passwords])
                for password in results:
                    if password:
                        break