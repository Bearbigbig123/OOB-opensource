import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import sqlalchemy

# 連接到現有的 SQLite DB
DATABASE_URL = "sqlite:///kpi_tracking.db"
engine = sqlalchemy.create_engine(DATABASE_URL)


# 檢查並新增 ChartImagePath 欄位（如不存在）
with engine.connect() as connection:
    col_info = pd.read_sql("PRAGMA table_info(kpi_definitions)", connection)
    if 'ChartImagePath' not in col_info['name'].tolist():
        connection.execute(sqlalchemy.text("ALTER TABLE kpi_definitions ADD COLUMN ChartImagePath TEXT"))
        print("已自動新增 ChartImagePath 欄位到 kpi_definitions")
    df = pd.read_sql("SELECT KpiDefID, GroupName, ChartName, ChartType FROM kpi_definitions", connection)

# 建立 assets/spc_charts 資料夾（Dash 靜態目錄）
assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
output_dir = os.path.join(assets_dir, 'spc_charts')
os.makedirs(output_dir, exist_ok=True)

# 依每一筆產生一張常態分布 SPC chart
image_paths = []
for idx, row in df.iterrows():
    np.random.seed(row['KpiDefID'])
    x = np.arange(1, 21)
    mu = 10 + row['ChartType']
    sigma = 1 + (row['ChartType'] % 3)
    y = np.random.normal(mu, sigma, size=len(x))
    plt.figure(figsize=(8, 4))
    plt.plot(x, y, marker='o', label='SPC Data')
    plt.title(f"{row['GroupName']}_{row['ChartName']}_{row['ChartType']}")
    plt.xlabel('樣本編號')
    plt.ylabel('測量值')
    plt.tight_layout()
    img_filename = f"spc_{row['KpiDefID']}.png"
    img_path = os.path.join(output_dir, img_filename)
    plt.savefig(img_path)
    plt.close()
    db_img_path = f"spc_charts/{img_filename}"
    image_paths.append((row['KpiDefID'], db_img_path))
    print(f"產生: {img_path}")

# 將圖片路徑寫回 DB（需有 ChartImagePath 欄位）
with engine.connect() as connection:
    for kpi_id, db_img_path in image_paths:
        connection.execute(
            sqlalchemy.text("UPDATE kpi_definitions SET ChartImagePath = :img WHERE KpiDefID = :kid"),
            {"img": db_img_path, "kid": kpi_id}
        )
    connection.commit()

print("所有 SPC 圖片已產生並寫入資料庫 ChartImagePath 欄位！")

# 額外：啟動一個 Flask 伺服器來 serve 圖片
from flask import Flask, send_from_directory
app = Flask(__name__)

@app.route('/spc_charts/<path:filename>')
def serve_spc_chart(filename):
    return send_from_directory(os.path.join(os.path.dirname(__file__), 'assets', 'spc_charts'), filename)

if __name__ == '__main__':
    print("啟動 Flask 圖片伺服器，請用 http://127.0.0.1:5000/spc_charts/xxx.png 存取圖片")
    app.run(port=5000)
