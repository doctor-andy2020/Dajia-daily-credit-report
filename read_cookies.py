import sqlite3
import os

db_path = os.path.expandvars(r'%APPDATA%\DM-Native\cache\Network\Cookies')
db = sqlite3.connect(db_path)
cur = db.cursor()

# List all cookies
cur.execute("SELECT host_key, name, encrypted_value FROM cookies WHERE host_key LIKE '%innodealing%'")
rows = cur.fetchall()
for r in rows:
    print(f'host: {r[0]}, name: {r[1]}, enc_val_len: {len(r[2]) if r[2] else 0}')

print("\n--- All hosts ---")
cur.execute("SELECT DISTINCT host_key FROM cookies")
for r in cur.fetchall():
    print(r[0])

db.close()
