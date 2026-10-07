import sqlite3
import pandas as pd

df = pd.read_csv("sales_project/data/sales_clean.csv", parse_dates=["Order Date"])
conn = sqlite3.connect("sales_project/sales.db")
df.to_sql("sales", conn, if_exists="replace", index=False)
conn.execute('CREATE INDEX IF NOT EXISTS idx_territory ON sales("Territory");')
conn.execute('CREATE INDEX IF NOT EXISTS idx_date ON sales("Order Date");')
conn.execute('CREATE INDEX IF NOT EXISTS idx_customer ON sales("Customer Name");')
conn.commit()
print(f"Loaded {conn.execute('SELECT COUNT(*) FROM sales').fetchone()[0]} rows")
conn.close()
