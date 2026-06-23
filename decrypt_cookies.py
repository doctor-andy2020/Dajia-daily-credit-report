"""从 CEF 浏览器缓存中提取 DM 网站 cookie"""
import sqlite3
import os
import json

try:
    import win32crypt
except ImportError:
    print("需要 pywin32，正在安装...")
    import subprocess
    subprocess.check_call(["pip", "install", "pywin32"])
    import win32crypt

db_path = os.path.expandvars(r'%APPDATA%\DM-Native\cache\Network\Cookies')

# 复制数据库（避免被锁定）
import shutil
tmp_db = os.path.expandvars(r'%TEMP%\dm_cookies.db')
shutil.copy2(db_path, tmp_db)

db = sqlite3.connect(tmp_db)
cur = db.cursor()

# 获取所有 innodealing 相关 cookie
cur.execute("""
    SELECT host_key, name, encrypted_value
    FROM cookies
    WHERE host_key LIKE '%innodealing%'
    ORDER BY host_key, name
""")

cookies = {}
for host, name, enc_val in cur.fetchall():
    try:
        # Windows DPAPI 解密
        decrypted = win32crypt.CryptUnprotectData(enc_val, None, None, None, 0)[1]
        value = decrypted.decode('utf-8', errors='replace')
        cookies[f"{host}|{name}"] = value
        print(f"{host} | {name} = {value}")
    except Exception as e:
        print(f"{host} | {name} = [解密失败: {e}]")

db.close()
os.unlink(tmp_db)
print("\n--- 所有 cookie 提取完毕 ---")
