import sqlite3
import os

db_path = 'robo_boah.db'
if os.path.exists(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = cursor.fetchall()
    print("Tabelas encontradas:")
    for t in tables:
        print(f"\n--- {t[0]} ---")
        cursor.execute(f"PRAGMA table_info({t[0]});")
        columns = cursor.fetchall()
        for col in columns:
            print(f"  {col[1]} ({col[2]})")
    conn.close()
else:
    print("Banco de dados não encontrado.")
