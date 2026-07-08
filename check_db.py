import sqlite3

conn = sqlite3.connect('agno.db')
c = conn.cursor()

c.execute("SELECT name FROM sqlite_master WHERE type='table'")
tables = c.fetchall()
print("=== Tables ===")
for t in tables:
    print(t[0])

for table_name in [t[0] for t in tables]:
    c.execute(f"SELECT * FROM {table_name}")
    rows = c.fetchall()
    print(f"\n=== {table_name} ({len(rows)} rows) ===")
    for row in rows:
        print(row)

conn.close()
