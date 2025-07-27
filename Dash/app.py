import dash
from dash import dcc, html, Input, Output, State, dash_table, callback
from dash.exceptions import PreventUpdate
import pandas as pd
import sqlalchemy
import json
import datetime
import dash_bootstrap_components as dbc
from dash import MATCH, ALL, callback_context
from flask import Response

# --- 數據庫設定 ---
DATABASE_URL = "sqlite:///kpi_tracking.db"
engine = sqlalchemy.create_engine(DATABASE_URL)

# --- 初始化數據庫表格和數據 ---
def initialize_db():
    feedback_sql = '''
    CREATE TABLE IF NOT EXISTS kpi_feedback (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        KpiDefID INTEGER,
        KPICol TEXT,
        feedback TEXT,
        action TEXT,
        user TEXT,
        timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
    );
    '''
    kpi_definitions_sql = """
    CREATE TABLE IF NOT EXISTS kpi_definitions (
        KpiDefID INTEGER PRIMARY KEY AUTOINCREMENT,
        FAB TEXT NOT NULL,
        GroupName TEXT,
        ChartName TEXT NOT NULL,
        ChartType INTEGER,
        HL_CHI TEXT,
        HL_Bimode TEXT,
        HL_PartRisk TEXT,
        HL_Kshift TEXT,
        HL_Zombie TEXT
    );
    """
    kpi_status_sql = """
    CREATE TABLE IF NOT EXISTS kpi_status (
        StatusID INTEGER PRIMARY KEY AUTOINCREMENT,
        KpiDefID INTEGER NOT NULL,
        Status_CHI TEXT,
        Status_Bimode TEXT,
        Status_PartRisk TEXT,
        Remark TEXT,
        LastUpdatedDate DATETIME,
        UpdatedBy TEXT,
        FOREIGN KEY (KpiDefID) REFERENCES kpi_definitions(KpiDefID)
    );
    """

    with engine.connect() as connection:
        connection.execute(sqlalchemy.text(kpi_definitions_sql))
        connection.execute(sqlalchemy.text(kpi_status_sql))
        connection.execute(sqlalchemy.text(feedback_sql))
        connection.commit()

        if connection.execute(sqlalchemy.text("SELECT COUNT(*) FROM kpi_definitions")).scalar() == 0:
            # 修改為您的實際 CSV 文件路徑
            csv_file_path = r"C:\\Users\\hsa00\\Desktop\\data.csv"
            try:
                initial_kpis = pd.read_csv(csv_file_path)
                if 'FAC' in initial_kpis.columns:
                    initial_kpis = initial_kpis.rename(columns={'FAC': 'FAB'})
                # 欄位名稱同步更名
                initial_kpis = initial_kpis.rename(columns={
                    'HL_K_Defined': 'HL_CHI',
                    'HL_B_Defined': 'HL_Bimode',
                    'HL_P_Defined': 'HL_PartRisk',
                    'HL_S_Defined': 'HL_Kshift',
                    'HL_W_Defined': 'HL_Zombie'
                })
                # 取得目前 DB 欄位
                db_cols_info = pd.read_sql("PRAGMA table_info(kpi_definitions)", connection)
                db_col_names = db_cols_info['name'].tolist()
                # 自動補齊 CSV 有但 DB 沒有的欄位
                for col in initial_kpis.columns:
                    if col not in db_col_names:
                        alter_sql = f"ALTER TABLE kpi_definitions ADD COLUMN '{col}' TEXT"
                        connection.execute(sqlalchemy.text(alter_sql))
                        connection.commit()
                        print(f"已自動新增欄位: {col}")
                        db_col_names.append(col)
                # 只匯入 DB 已有的欄位
                initial_kpis = initial_kpis[[col for col in initial_kpis.columns if col in db_col_names]]
                # 確保 HL 欄位存在
                for col in ['HL_CHI', 'HL_Bimode', 'HL_PartRisk', 'HL_Kshift', 'HL_Zombie']:
                    if col not in initial_kpis.columns:
                        initial_kpis[col] = ''
                initial_kpis.to_sql('kpi_definitions', connection, if_exists='append', index=False)
                connection.commit()
                print(f"成功從 {csv_file_path} 載入初始 KPI 定義到資料庫。")
            except FileNotFoundError:
                raise RuntimeError(f"錯誤：找不到初始 KPI 定義檔案 '{csv_file_path}'。請檢查路徑或檔案名稱。")
            except Exception as e:
                raise RuntimeError(f"從 CSV 載入 KPI 定義時發生錯誤: {e}")

        defined_kpis_df = pd.read_sql("SELECT KpiDefID FROM kpi_definitions", connection)
        for kpi_id in defined_kpis_df['KpiDefID']:
            if connection.execute(sqlalchemy.text(f"SELECT COUNT(*) FROM kpi_status WHERE KpiDefID = {kpi_id}")).scalar() == 0:
                connection.execute(
                    sqlalchemy.text("INSERT INTO kpi_status (KpiDefID, Status_CHI, Status_Bimode, Status_PartRisk, Remark, LastUpdatedDate, UpdatedBy) VALUES (:kpi_id, :sk, :sb, :sp, :remark, :date, :by)"),
                    {"kpi_id": kpi_id, "sk": "N", "sb": "N", "sp": "N", "remark": "", "date": datetime.datetime.now(), "by": "System"}
                )
        connection.commit()

# --- Dash 應用程式設定 ---
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP, dbc.icons.FONT_AWESOME], suppress_callback_exceptions=True)

# --- 讀取數據的輔助函數 ---
def get_kpi_data(fab=None):
    with engine.connect() as connection:
        query = """
        SELECT
            kd.KpiDefID,
            kd.FAB,
            kd.GroupName,
            kd.ChartName,
            kd.ChartType,
            kd.Module,
            kd.HL_CHI,
            kd.HL_Bimode,
            kd.HL_PartRisk,
            kd.HL_Kshift,
            kd.HL_Zombie,
            ks.Remark,
            ks.LastUpdatedDate,
            ks.UpdatedBy
        FROM kpi_definitions kd
        LEFT JOIN kpi_status ks ON kd.KpiDefID = ks.KpiDefID
        """
        if fab:
            query += f" WHERE kd.FAB = '{fab}'"
        df = pd.read_sql(query, connection)

        df['LastUpdatedDate'] = pd.to_datetime(df['LastUpdatedDate'])
        for col in ['HL_CHI', 'HL_Bimode', 'HL_PartRisk', 'HL_Kshift', 'HL_Zombie']:
            df[col] = df[col].fillna("").replace({None: ""})
        df['Remark'] = df['Remark'].fillna('')
        def format_date_with_week(dt):
            if pd.isna(dt):
                return ''
            date_str = dt.strftime('%Y-%m-%d')
            iso = dt.isocalendar()
            # 週數前加上年份最後一碼
            week_str = f"W{str(iso.year)[-1]}{iso.week:02d}"
            return f"{date_str} ({week_str})"
        df['LastUpdatedDate'] = df['LastUpdatedDate'].apply(format_date_with_week)
    return df

# --- 帳號密碼設定 ---
USER_DB = {
    "admin": "8888",
    "user1": "1234",
    "user2": "5678"
}

# --- 登入頁面內容 ---
login_layout = dbc.Container([
    dbc.Row([
        dbc.Col(width=4),
        dbc.Col([
            dbc.Card([
                dbc.CardBody([
                    html.H2("KPI 追蹤系統", className="mb-4 text-center text-primary fw-bold"),
                    html.P("請輸入您的帳號密碼登入", className="text-muted text-center mb-3"),
                    dbc.Input(id='input-username', placeholder='帳號', type='text', size="lg", className="mb-3 border-primary"), # 更大、有邊框
                    dbc.Input(id='input-password', placeholder='密碼', type='password', size="lg", className="mb-3 border-primary"),
                    dbc.Button('登入', id='btn-login', color='primary', size="lg", className="w-100 mb-3 fw-bold"), # 更大、粗體
                    html.Div(id='login-message', className="text-danger text-center fw-bold small"),
                ])
            ], className="shadow-lg p-5 bg-white rounded mt-5") # 更大的陰影、更多填充
        ], width=4),
        dbc.Col(width=4)
    ], className="align-items-center min-vh-100") # 垂直置中
], fluid=True, className="bg-light login-container") # 整個登入背景也變淺色, 增加 class for CSS

# --- 主頁面 layout 內容 ---
def get_kpi_matrix_figure(fab=None, groupname=None, chartname=None, charttype=None):
    import plotly.graph_objs as go
    from plotly.subplots import make_subplots
    df = get_kpi_data(fab=fab)
    if groupname:
        df = df[df['GroupName'] == groupname]
    if chartname:
        df = df[df['ChartName'] == chartname]
    if charttype:
        df = df[df['ChartType'] == charttype]
    kpi_cols = ['HL_CHI', 'HL_Bimode', 'HL_PartRisk', 'HL_Kshift', 'HL_Zombie']
    # 取得所有 Module（僅用於欄位名稱）
    with engine.connect() as connection:
        all_modules = pd.read_sql("SELECT DISTINCT Module FROM kpi_definitions", connection)['Module'].dropna().astype(str).unique().tolist()
    modules = sorted([m for m in all_modules if str(m).strip() != ''])
    # 取得所有周次（x軸）
    # 先將 LastUpdatedDate 轉回 datetime
    df['LastUpdatedDate_dt'] = pd.to_datetime(df['LastUpdatedDate'].str[:10], errors='coerce')
    # 取所有週次（格式 W526）
    def get_week_str(dt):
        if pd.isna(dt):
            return ''
        iso = dt.isocalendar()
        return f"W{str(iso.year)[-1]}{iso.week:02d}"
    df['Week'] = df['LastUpdatedDate_dt'].apply(get_week_str)
    week_list = sorted([w for w in df['Week'].unique() if w], key=lambda x: (x[1:], x[0]))
    if not week_list:
        week_list = ['W00']
    n_weeks = len(week_list)
    # --- 統計每個 KPI x week x module 的數量 ---
    # 結構：row=KPI, col=Module，x軸=week
    # 先建立一個 dict: {(kpi, module): {week: count}}，kpi 用原始 kpi_cols
    stat_dict = {}
    for kpi in ['Total'] + kpi_cols:
        for module in modules:
            stat_dict[(kpi, module)] = {w: 0 for w in week_list}
    for module in modules:
        df_mod = df[df['Module'].astype(str) == module]
        for kpi in kpi_cols:
            for week in week_list:
                if not df_mod.empty:
                    count = df_mod[(df_mod['Week'] == week) & (~df_mod[kpi].isin(["", "Finish", "Waive"]))][kpi].count()
                else:
                    count = 0
                stat_dict[(kpi, module)][week] = count
        # Total 行：該 module 該週所有 KPI 欄位的有效數量總和
        for week in week_list:
            total_count = 0
            if not df_mod.empty:
                for kpi in kpi_cols:
                    total_count += df_mod[(df_mod['Week'] == week) & (~df_mod[kpi].isin(["", "Finish", "Waive"]))][kpi].count()
            stat_dict[('Total', module)][week] = total_count
    # y_labels 不顯示 HL_ 前綴，但統計時仍用原始 kpi_cols
    # 先計算每個 row/col 的總和
    row_sums = []
    col_sums = []
    # 計算 row sum（每個KPI across modules & weeks）
    for idx, kpi in enumerate(['Total'] + kpi_cols):
        total = 0
        for module in modules:
            for week in week_list:
                total += stat_dict[(kpi, module)][week]
        row_sums.append(total)
    # 計算 col sum（每個module across KPI & weeks）(不包含'Total'行)
    for module in modules:
        total = 0
        for kpi in kpi_cols:  # 不包含'Total'
            for week in week_list:
                total += stat_dict[(kpi, module)][week]
        col_sums.append(total)
    # y_labels 顯示 KPI 名稱與總數
    kpi_labels = [f"{k.replace('HL_', '')} : {row_sums[i+1]}" for i, k in enumerate(kpi_cols)]
    y_labels = [f"Total : {row_sums[0]}"] + kpi_labels
    n_rows = len(y_labels)
    from plotly.subplots import make_subplots
    fig = make_subplots(
        rows=n_rows, cols=len(modules),
        horizontal_spacing=0.01,
        vertical_spacing=0.03,
        subplot_titles=modules
    )
    # Dash 介面主色為 Bootstrap primary: #0d6efd
    bar_color = '#0d6efd'
    for col_idx, module in enumerate(modules):
        module_title = f"{module} : {col_sums[col_idx]}"
        # 修改 subplot_titles
        fig.layout.annotations[col_idx].text = module_title
        for row_idx, kpi in enumerate(y_labels):
            # 反查原始 stat_dict 的 key
            if row_idx == 0:
                kpi_key = 'Total'
            else:
                kpi_key = kpi_cols[row_idx-1]
            y_vals = [stat_dict[(kpi_key, module)][w] for w in week_list]
            fig.add_trace(
                go.Bar(
                    x=week_list,
                    y=y_vals,
                    marker_color=bar_color,
                    name=f"{kpi}-{module}",
                    width=0.2
                ),
                row=row_idx+1, col=col_idx+1
            )
    # 設定 y 軸標籤（KPI 標籤）字體大小與粗體
    yaxis_title_font = dict(size=16, family="Microsoft JhengHei, Arial, sans-serif", color="#222", weight="bold")
    for i, label in enumerate(y_labels):
        fig.update_yaxes(title_text=label, row=i+1, col=1, title_font=yaxis_title_font)

    # 設定 module 標籤（subplot_titles）字體大小與粗體
    fig.update_layout(
        # ...其他 layout 設定...
        font=dict(size=16, family="Microsoft JhengHei, Arial, sans-serif", color="#222"),
        title_font=dict(size=16, family="Microsoft JhengHei, Arial, sans-serif", color="#222", weight="bold"),
        # 下面的 annotations 是 subplot_titles
        annotations=[
            dict(
                x=ann["x"],
                y=ann["y"],
                text=ann["text"],
                xref=ann["xref"],
                yref=ann["yref"],
                showarrow=False,
                font=dict(size=16, family="Microsoft JhengHei, Arial, sans-serif", color="#222", weight="bold")
            ) if "text" in ann else ann
            for ann in fig.layout.annotations
        ]
    )
    for r in range(1, n_rows+1):
        for c in range(1, len(modules)+1):
            fig.update_xaxes(
                row=r, col=c,
                tickfont=dict(size=10),
                tickangle=0,
                automargin=True
            )
            fig.update_yaxes(
                row=r, col=c,
                tickfont=dict(size=10),
                automargin=True,
                dtick=1,
                tickformat='d',
                rangemode='tozero',
                showgrid=False  # 拿掉水平白色刻度線
            )
    fig.update_layout(
        height=130*n_rows,   # 原本120，放大到170
        width=130*len(modules),  # 原本120，放大到170
        showlegend=False,
        margin=dict(t=30, l=5, r=5, b=5),
        bargap=0.08,
        bargroupgap=0.05
    )
    return fig

def get_main_layout_content():
    today = datetime.datetime.now()
    date_str = today.strftime('%Y-%m-%d')
    week_str = f"W{today.isocalendar().week:02d}"
    # 修正週數顯示格式為 W526
    week_str = f"W{str(today.isocalendar().year)[-1]}{today.isocalendar().week:02d}"
    today_display = f"{date_str} ({week_str})"

    # 新增三個篩選 dropdown（初始 options 為空，callback 會填充）
    filter_row = html.Div([
        dbc.Row([
            dbc.Col([
                html.Label("Module", className="fw-bold mb-1"),
                dcc.Dropdown(id='module-filter', options=[], placeholder="全部")
            ], width=3),
            dbc.Col([
                html.Label("GroupName", className="fw-bold mb-1"),
                dcc.Dropdown(id='groupname-filter', options=[], placeholder="全部")
            ], width=3),
            dbc.Col([
                html.Label("ChartName", className="fw-bold mb-1"),
                dcc.Dropdown(id='chartname-filter', options=[], placeholder="全部")
            ], width=3),
            dbc.Col([
                html.Label("ChartType", className="fw-bold mb-1"),
                dcc.Dropdown(id='charttype-filter', options=[], placeholder="全部")
            ], width=3),
        ], className="mb-2")
    ])

    # --- KPI 矩陣子圖（Module為X軸，Y軸為Total/CHI/Bimode/PartRisk/Kshift/Zombie）---
    return html.Div([
        dbc.Row([
            dbc.Col([
                dbc.Card([
                    dbc.CardHeader(html.H4("FAB 選擇與操作說明", style={"fontSize": "16px"}, className="mb-0 text-white"), className="bg-dark"),
                    dbc.CardBody([
                        html.Label("請選擇廠區 (FAB):", className="fw-bold mb-2 text-primary", style={"fontSize": "16px"}),
                        dcc.Dropdown(
                            id='fac-selector',
                            placeholder="請選擇 FAB...",
                            className="mb-4 border-primary",
                            style={"fontSize": "0.95rem"}
                        ),
                        html.Hr(className="my-3"),
                        html.H5("操作說明：", className="mt-4 mb-3 text-info", style={"fontSize": "16px"}),
                        html.Ul([
                            html.Li("點擊『儲存變更』按鈕，將表格異動儲存至資料庫。", className="text-secondary small mb-1"),
                            html.Li("『備註』欄位自由輸入，按 Enter 或切換焦點即儲存。", className="text-secondary small mb-1"),
                            html.Li(html.Span(["欄位顏色說明：",
                                               html.Span("黃色", style={"color": "#CC7000", "fontWeight": "bold"}), "=有值、",
                                               html.Span("灰色", style={"color": "#6C757D", "fontWeight": "bold"}), "=無值、",
                                               html.Span("粉紅色", style={"color": "#C2185B", "fontWeight": "bold"}), "=Waive、",
                                               html.Span("綠色", style={"color": "#155724", "fontWeight": "bold"}), "=Finish。"]), className="text-secondary small")
                        ], className="list-unstyled")
                    ], style={"fontSize": "14px"})
                ], className="shadow-sm"),  # <--- 移除 h-100
            ], width=2, className="mt-2"),  # <--- mt-4 改 mt-2
            dbc.Col([
                html.Div(id="dashboard-placeholder"),
                filter_row,
                dbc.Card([
                    dbc.CardHeader([
                        html.H4(id='selected-fac-title', className="mb-0 text-white"),
                        html.Div(today_display, className="text-white-50 fw-bold fst-italic", style={"fontSize": "1rem"})
                    ], className="bg-primary d-flex justify-content-between align-items-center"),
                    dbc.CardBody([
                        dcc.Loading(
                            id="loading-kpi-table",
                            type="default",
                            children=html.Div(id='kpi-table-container', className="mb-3", style={"minHeight": "300px"}),
                        ),
                        # 這裡用 dcc.Store 控制顯示
                        dcc.Store(id='fac-selected-flag', storage_type='memory'),
                        # 儲存按鈕區塊動態顯示：只有選擇 FAB 才 render
                        html.Div(id="save-btn-block-portal")
                    ], id="main-table-body")
                ], className="shadow-sm"),  # <--- 移除 h-100
                # KPI矩陣圖放在KPI回覆下方
                dbc.Card([
                    dbc.CardHeader(
                        html.H5("KPI 狀態矩陣 (Module x KPI)", className="mb-0 text-white fw-bold"),
                        className="bg-primary"
                    ),
                    dbc.CardBody([
                        dcc.Graph(
                            id='kpi-matrix-graph',
                            style={
                                "marginBottom": "0px",
                                "paddingBottom": "0px",
                                "width": "100%",
                                "overflowX": "auto",
                                "marginLeft": "auto",
                                "marginRight": "auto",
                                "display": "block",
                                "maxWidth": "100vw"
                            },
                            config={
                                "responsive": True,
                                "displayModeBar": True,
                                "scrollZoom": True
                            }
                        )
                    ], style={
                        "background": "#fff",
                        "padding": "32px 32px 32px 32px",
                        "minHeight": "calc(170px * 7)",  # 跟圖表高度一致
                        "display": "flex",
                        "alignItems": "center",
                        "justifyContent": "center"
                    })
                ], className="shadow-sm", style={"marginBottom": "0px", "marginTop": "8px"})
            ], width=10, className="mt-2")  # <--- mt-4 改 mt-2
        ])
    ], className="mb-2")  # <--- mb-4 改 mb-2

# --- KPI 矩陣圖即時刷新 callback ---
@callback(
    Output('kpi-matrix-graph', 'figure'),
    Input('fac-selector', 'value'),
    Input('groupname-filter', 'value'),
    Input('chartname-filter', 'value'),
    Input('charttype-filter', 'value'),
    State('login-user', 'data'),
)
def update_kpi_matrix_figure(fab, groupname, chartname, charttype, user):
    if not user or not fab:
        import plotly.graph_objs as go
        return go.Figure()
    return get_kpi_matrix_figure(fab=fab, groupname=groupname, chartname=chartname, charttype=charttype)


# --- FAB 選擇時才渲染儲存按鈕區塊 ---
@callback(
    Output('save-btn-block-portal', 'children'),
    Input('fac-selector', 'value'),
    prevent_initial_call=False
)
def render_save_btn_block(fab_value):
    if fab_value:
        return html.Div([
            dbc.Button(
                [html.I(className="fas fa-save me-2"), "儲存變更"],
                id="btn-save-kpi",
                color="success",
                className="me-3 fw-bold"
            ),
            dbc.Button(
                [html.I(className="fas fa-check-double me-2"), "CHI 全部 Finish"],
                id="btn-finish-chi-all",
                color="primary",
                className="me-2 fw-bold"
            ),
            dbc.Button(
                [html.I(className="fas fa-check-double me-2"), "PartRisk 全部 Finish"],
                id="btn-finish-partrisk-all",
                color="primary",
                className="me-3 fw-bold"
            ),
            html.Div(id='save-output-message', className="flex-grow-1")
        ], className="d-flex align-items-center justify-content-end mt-3", id="save-btn-block")
    return None
# --- 一鍵 Finish 按鈕 callback ---
@callback(
    Output('dashboard-placeholder', 'children', allow_duplicate=True),
    Output('kpi-table-container', 'children', allow_duplicate=True),
    Output('groupname-filter', 'options', allow_duplicate=True),
    Output('chartname-filter', 'options', allow_duplicate=True),
    Output('charttype-filter', 'options', allow_duplicate=True),
    Output('save-output-message', 'children', allow_duplicate=True),
    Input('btn-finish-chi-all', 'n_clicks'),
    Input('btn-finish-partrisk-all', 'n_clicks'),
    State('fac-selector', 'value'),
    State('login-user', 'data'),
    State('groupname-filter', 'value'),
    State('chartname-filter', 'value'),
    State('charttype-filter', 'value'),
    prevent_initial_call=True
)
def finish_all_chi_partrisk(btn_chi, btn_partrisk, selected_fab, user, groupname, chartname, charttype):
    ctx = callback_context
    # 僅允許由 finish 按鈕觸發，且必須有 n_clicks > 0，避免 FAB 選擇時自動觸發
    if not ctx.triggered:
        raise PreventUpdate
    trigger = ctx.triggered[0]['prop_id']
    # 僅當按鈕被實際點擊才執行
    if trigger == 'btn-finish-chi-all.n_clicks' and not btn_chi:
        raise PreventUpdate
    if trigger == 'btn-finish-partrisk-all.n_clicks' and not btn_partrisk:
        raise PreventUpdate
    if trigger not in ['btn-finish-chi-all.n_clicks', 'btn-finish-partrisk-all.n_clicks']:
        raise PreventUpdate
    if not selected_fab or not user:
        raise PreventUpdate
    if trigger == 'btn-finish-chi-all.n_clicks':
        col = 'HL_CHI'
    elif trigger == 'btn-finish-partrisk-all.n_clicks':
        col = 'HL_PartRisk'
    else:
        raise PreventUpdate
    updated_count = 0
    now = datetime.datetime.now()
    with engine.connect() as connection:
        # 只更新目前 FAB 下的資料，且只針對原本有值且不是 Finish 的欄位
        df = pd.read_sql(f"SELECT KpiDefID, {col} FROM kpi_definitions WHERE FAB = :fab", connection, params={"fab": selected_fab})
        for _, row in df.iterrows():
            kpi_id = row['KpiDefID']
            old_val = str(row[col]) if pd.notna(row[col]) else ""
            # 只更新原本有值且不是 Finish 的欄位（空值不動）
            if old_val != "" and old_val != "Finish":
                update_sql = sqlalchemy.text(f"UPDATE kpi_definitions SET {col} = :val WHERE KpiDefID = :kid")
                connection.execute(update_sql, {"val": "Finish", "kid": kpi_id})
                update_status_sql = sqlalchemy.text("UPDATE kpi_status SET UpdatedBy = :by, LastUpdatedDate = :dt WHERE KpiDefID = :kid")
                connection.execute(update_status_sql, {"by": user, "dt": now, "kid": kpi_id})
                updated_count += 1
        connection.commit()
    # 更新完直接刷新表格
    dashboard, table, gopt, copt, topt = render_kpi_table(selected_fab, groupname, chartname, charttype, user)
    msg = dbc.Alert(f"已將 {col.replace('HL_', '')} 欄位全部設為 Finish，共 {updated_count} 筆！", color="success", dismissable=True, duration=3000) if updated_count > 0 else dbc.Alert("無需更新，所有欄位已是 Finish。", color="info", dismissable=True, duration=3000)
    return dashboard, table, gopt, copt, topt, msg

# --- 網頁佈局 ---
app.layout = html.Div([
    dcc.Store(id='login-user', storage_type='session'),
    dcc.Store(id='current-fab', storage_type='session'),
    html.Div(id='login-container', children=login_layout),
    html.Div(id='main-app-container', style={'display': 'none'}, children=[
        dbc.Navbar(
            children=[
                html.Div([
                    html.Div("KPI追蹤系統", style={
                        "fontWeight": "bold",
                        "fontSize": "1.5rem",
                        "letterSpacing": "0.1em",
                        "color": "white",
                        "paddingLeft": "1.2rem",
                        "marginRight": "2rem",
                        "flex": "0 0 auto"
                    }),
                    html.Div([
                        dbc.Nav(
                            [
                                dbc.NavItem(dbc.NavLink("主頁面", href="/", active="exact")),
                                dbc.NavItem(dbc.NavLink(html.Span(id='navbar-username', className="ms-2 me-2 fs-6 fw-bold text-light"), href="#")),
                                dbc.DropdownMenu(
                                    children=[
                                        dbc.DropdownMenuItem([html.I(className="fas fa-database me-2"), "查看資料庫"], id="btn-view-db"),
                                        dbc.DropdownMenuItem(divider=True),
                                        dbc.DropdownMenuItem([html.I(className="fas fa-sign-out-alt me-2"), "登出"], id="btn-logout"),
                                    ],
                                    nav=True,
                                    in_navbar=True,
                                    label="更多",
                                    align_end=True,
                                    toggle_style={"color": "white"}
                                ),
                            ],
                            className="ms-auto",
                            navbar=True
                        )
                    ], style={"display": "flex", "flex": "1 1 auto", "alignItems": "center", "justifyContent": "flex-end"})
                ], style={"display": "flex", "width": "100%", "alignItems": "center"})
            ],
            color="primary",
            dark=True,
            className="shadow-sm mb-4",
            style={"paddingLeft": "0", "paddingRight": "0", "minHeight": "56px"}
        ),
        dbc.Container(id='page-content', fluid=True)
    ]),

    dbc.Modal(
        [
            dbc.ModalHeader(dbc.ModalTitle("資料庫內容 (kpi_definitions & kpi_status)")),
            dbc.ModalBody(
                dash_table.DataTable(
                    id='db-table',
                    columns=[],
                    data=[],
                    page_size=10,
                    style_table={'overflowX': 'auto'},
                    style_cell={'textAlign': 'left', 'padding': '5px'},
                    style_header={
                        'backgroundColor': 'rgb(230, 230, 230)',
                        'fontWeight': 'bold'
                    }
                )
            ),
            dbc.ModalFooter(
                dbc.Button("關閉", id="close-db-modal", className="ms-auto", color="secondary")
            ),
        ],
        id="db-modal",
        size="xl",
        is_open=False,
        scrollable=True
    ),
        html.Div(id='dropdown-portal'),

    # --- KPI 回覆 Modal ---
    dbc.Modal([
        dbc.ModalHeader(dbc.ModalTitle("KPI 回覆")),
        dbc.ModalBody(id="reply-modal-body"),
        dbc.ModalFooter([
            dbc.Button("儲存回覆", id="modal-save-btn", color="primary", className="me-2"),
            dbc.Button("關閉", id="close-reply-modal", color="secondary")
        ]),
    ], id="reply-modal", is_open=False, size="xl", style={"maxWidth": "1600px", "width": "99vw", "minWidth": "900px"})
])

# --- Callbacks ---

# Callback 處理登入
@callback(
    Output('login-user', 'data'),
    Output('login-message', 'children'),
    Output('main-app-container', 'style'),
    Output('login-container', 'style'),
    Input('btn-login', 'n_clicks'),
    State('input-username', 'value'),
    State('input-password', 'value'),
    prevent_initial_call=True
)
def handle_authentication(login_clicks, username, password):
    if not login_clicks:
        raise PreventUpdate
    if not username or not password:
        return None, '請輸入帳號與密碼', {'display': 'none'}, {'display': 'block'}
    if username in USER_DB and USER_DB[username] == password:
        return username, '', {'display': 'block'}, {'display': 'none'}
    else:
        return None, '帳號或密碼錯誤', {'display': 'none'}, {'display': 'block'}

# Callback 處理登出
@callback(
    Output('login-user', 'data', allow_duplicate=True),
    Output('main-app-container', 'style', allow_duplicate=True),
    Output('login-container', 'style', allow_duplicate=True),
    Input('btn-logout', 'n_clicks'),
    prevent_initial_call=True
)
def handle_logout(logout_clicks):
    if not logout_clicks:
        raise PreventUpdate
    return None, {'display': 'none'}, {'display': 'block'}

# Callback 顯示導航欄使用者名稱及控制主頁面內容
@callback(
    Output('navbar-username', 'children'),
    Output('page-content', 'children'),
    Input('login-user', 'data'),
    prevent_initial_call=False
)
def update_main_content(user):
    if user:
        return f"歡迎，{user}!", get_main_layout_content()
    return "", html.Div()

# Callback 填充左側FAB選擇器
@callback(
    Output('fac-selector', 'options'),
    Input('login-user', 'data'),
    State('fac-selector', 'options')
)
def set_fac_options(user, current_options):
    if user and not current_options:
        with engine.connect() as connection:
            facs = pd.read_sql("SELECT DISTINCT FAB FROM kpi_definitions ORDER BY FAB", connection)['FAB'].tolist()
        return [{'label': f, 'value': f} for f in facs]
    raise PreventUpdate

# Callback 顯示選定的FAB標題
@callback(
    Output('selected-fac-title', 'children'),
    Input('fac-selector', 'value'),
    State('login-user', 'data')
)
def display_selected_fab(selected_fab, user):
    if not user:
        return "請先登入"
    if selected_fab:
        return f"{selected_fab} 的 KPI 回覆"
    return "請選擇一個FAB"

# 根據選定FAB產生自訂表格（含 dcc.Dropdown 和 dcc.Input）
@callback(
    Output('dashboard-placeholder', 'children'),
    Output('kpi-table-container', 'children'),
    Output('module-filter', 'options'),
    Output('groupname-filter', 'options'),
    Output('chartname-filter', 'options'),
    Output('charttype-filter', 'options'),
    Input('fac-selector', 'value'),
    Input('module-filter', 'value'),
    Input('groupname-filter', 'value'),
    Input('chartname-filter', 'value'),
    Input('charttype-filter', 'value'),
    State('login-user', 'data'),
    prevent_initial_call=True
)
def render_kpi_table(selected_fab, module, groupname, chartname, charttype, user):
    if not user or not selected_fab:
        raise PreventUpdate

    df = get_kpi_data(fab=selected_fab)
    if df.empty:
        return None, dbc.Alert(f"找不到 {selected_fab} 的 KPI 資料。", color="info", className="mt-3 text-center"), [], [], []

    # --- KPI 完成率 dashboard 計算 ---
    dashboard_cols = ['HL_CHI', 'HL_Bimode', 'HL_PartRisk', 'HL_Kshift', 'HL_Zombie']
    dashboard_stats = []
    total_need_review = 0
    total_finish = 0
    total_waive = 0
    for col in dashboard_cols:
        col_vals = df[col].fillna("")
        raw_count = (col_vals != "").sum()
        finish_count = (col_vals == "Finish").sum()
        waive_count = (col_vals == "Waive").sum()
        denominator = raw_count - waive_count  # 分母需扣掉 Waive
        finish_rate = (finish_count / denominator * 100) if denominator > 0 else 0
        dashboard_stats.append({
            "col": col.replace("HL_", ""),
            "raw": raw_count,
            "denominator": denominator,
            "finish": finish_count,
            "waive": waive_count,
            "rate": finish_rate
        })
        total_need_review += denominator
        total_finish += finish_count
        total_waive += waive_count

    # 總 KPI 完成率
    total_kpi = total_need_review
    total_finish_rate = (total_finish / total_kpi * 100) if total_kpi > 0 else 0

    # 儀表板視覺設計（總覽卡片重新設計，與五張KPI卡片高度一致）
    dashboard = html.Div([
        html.Div([
            # 重新設計的總覽卡片（由上到下排版，英文標籤）
            html.Div([
                html.Div([
                    html.I(className="fas fa-chart-pie me-2", style={"color": "#0d6efd", "fontSize": "22px", "marginRight": "6px"}),
                    html.Span("Overview", className="fw-bold text-secondary", style={"fontSize": "18px"})
                ], style={"display": "flex", "alignItems": "center", "marginBottom": "2px"}),
                html.Div([
                    html.Span("Need Review:", style={"color": "#6c757d", "fontWeight": 500, "fontSize": "15px"}),
                    html.Span(" ", style={"display": "inline-block", "width": "6px"}),
                    html.Span(f"{total_kpi}", style={"color": "#fd7e14", "fontWeight": 700, "fontSize": "15px"})
                ], style={"marginBottom": "2px", "display": "flex", "alignItems": "center"}),
                html.Div([
                    html.Span("Finish:", style={"color": "#6c757d", "fontWeight": 500, "fontSize": "15px"}),
                    html.Span(" ", style={"display": "inline-block", "width": "6px"}),
                    html.Span(f"{total_finish}", style={"color": "#198754", "fontWeight": 700, "fontSize": "15px"})
                ], style={"marginBottom": "2px", "display": "flex", "alignItems": "center"}),
                html.Div([
                    html.Span(
                        f"完成率: {'N/A' if total_kpi == 0 else f'{total_finish_rate:.1f}%'}",
                        className="fw-bold",
                        style={"fontSize": "15px", "color": "#0d6efd"}
                    )
                ], style={"marginBottom": "8px", "display": "flex", "alignItems": "center"}),
                dbc.Progress(value=100 if total_kpi == 0 else (total_kpi-total_finish)/total_kpi*100, color="primary", style={"height": "8px", "width": "60px", "marginLeft": "0px", "marginBottom": "8px", "display": "none"}, animated=True, striped=True),
                dbc.Progress(value=total_finish_rate, color="primary", style={"height": "8px", "marginTop": "0px", "width": "100%"}, animated=True, striped=True),
            ], style={
                "background": "linear-gradient(90deg, #f8fafc 60%, #e3f2fd 100%)",
                "borderRadius": "1rem",
                "boxShadow": "0 2px 8px 0 rgba(0,0,0,0.07)",
                "padding": "16px 14px 10px 18px",
                "border": "1.5px solid #e3e6e9",
                "minWidth": "193px",
                "maxWidth": "243px",
                "minHeight": "140px",
                "maxHeight": "160px",
                "display": "flex",
                "flexDirection": "column",
                "alignItems": "flex-start",
                "justifyContent": "center"
            }),
            # 五張KPI卡片
            *(html.Div([
                html.Div(stat["col"], className="fw-bold", style={"fontSize": "16px", "color": "#0d6efd", "marginBottom": "2px", "letterSpacing": "1px", "fontWeight": "bold"}),
                html.Div([
                    html.Span("Need Review:", style={"color": "#6c757d", "fontWeight": 500}),
                    html.Span(" ", style={"display": "inline-block", "width": "6px"}),
                    html.Span(f"{stat['denominator']}", style={"color": "#fd7e14", "fontWeight": 700, "fontSize": "15px"})
                ], style={"marginBottom": "2px", "display": "flex", "alignItems": "center"}),
                html.Div([
                    html.Span("Finish:", style={"color": "#6c757d", "fontWeight": 500}),
                    html.Span(" ", style={"display": "inline-block", "width": "6px"}),
                    html.Span(f"{stat['finish']}", style={"color": "#198754", "fontWeight": 700, "fontSize": "15px"})
                ], style={"marginBottom": "2px", "display": "flex", "alignItems": "center"}),
                html.Div([
                    html.Span(
                        "完成率: N/A" if stat['denominator'] == 0 else f"完成率: {stat['rate']:.1f}%",
                        className="fw-bold",
                        style={"fontSize": "15px", "color": "#0d6efd"}
                    )
                ], style={"marginTop": "2px", "marginBottom": "8px", "display": "flex", "alignItems": "center"}),
                dbc.Progress(value=stat['rate'], color="primary", style={"height": "8px", "marginTop": "0px", "width": "100%"}, animated=True, striped=True)
            ], style={
                "background": "linear-gradient(90deg, #f8fafc 60%, #e3f2fd 100%)",
                "borderRadius": "1rem",
                "boxShadow": "0 2px 8px 0 rgba(0,0,0,0.07)",
                "padding": "16px 14px 10px 18px",
                "border": "1.5px solid #e3e6e9",
                "minWidth": "193px",
                "maxWidth": "243px",
                "minHeight": "140px",
                "maxHeight": "160px",
                "display": "flex",
                "flexDirection": "column",
                "alignItems": "flex-start",
                "justifyContent": "center"
            }) for stat in dashboard_stats),
        ], style={
            "display": "flex",
            "flexDirection": "row",
            "gap": "12px",
            "width": "100%"
        })
    ], style={"marginBottom": "32px", "paddingLeft": "0px", "paddingRight": "0px", "width": "100%"})

    # 依 module 篩選
    if module:
        df = df[df['Module'] == module]
    if groupname:
        df = df[df['GroupName'] == groupname]
    if chartname:
        df = df[df['ChartName'] == chartname]
    if charttype:
        df = df[df['ChartType'] == charttype]

    # 產生 options（module_options 需用未被 module 篩選的 df）
    all_df = get_kpi_data(fab=selected_fab) if selected_fab else df
    module_options = [{'label': v, 'value': v} for v in sorted(all_df['Module'].dropna().unique())]
    groupname_options = [{'label': v, 'value': v} for v in sorted(df['GroupName'].dropna().unique())]
    chartname_options = [{'label': v, 'value': v} for v in sorted(df['ChartName'].dropna().unique())]
    charttype_options = [{'label': str(v), 'value': v} for v in sorted(df['ChartType'].dropna().unique())]

    # 查詢所有KPI的feedback/action是否有內容
    feedback_status = {}
    with engine.connect() as connection:
        for kpi_id in df['KpiDefID']:
            fb = connection.execute(
                sqlalchemy.text("SELECT feedback, action FROM kpi_feedback WHERE KpiDefID = :kid ORDER BY timestamp DESC LIMIT 1"),
                {"kid": kpi_id}
            ).fetchone()
            has_reply = (fb and ((fb[0] and str(fb[0]).strip() != "") or (fb[1] and str(fb[1]).strip() != "")))
            feedback_status[kpi_id] = has_reply

    header = html.Thead(html.Tr([
        html.Th("ID", style={"minWidth": "40px", "textAlign": "center", "backgroundColor": "#e9ecef", "fontWeight": "bold"}),
        html.Th("FAB", style={"minWidth": "60px", "textAlign": "center", "backgroundColor": "#e9ecef", "fontWeight": "bold"}),
        html.Th("回覆", style={"minWidth": "40px", "textAlign": "center", "backgroundColor": "#e9ecef", "fontWeight": "bold"}),
        html.Th("回覆操作", style={"minWidth": "60px", "textAlign": "center", "backgroundColor": "#e9ecef", "fontWeight": "bold"}),
        html.Th("Module", style={"minWidth": "80px", "textAlign": "center", "backgroundColor": "#e9ecef", "fontWeight": "bold"}),
        html.Th("GroupName", style={"minWidth": "100px", "textAlign": "center", "backgroundColor": "#e9ecef", "fontWeight": "bold"}),
        html.Th("ChartName", style={"minWidth": "120px", "textAlign": "center", "backgroundColor": "#e9ecef", "fontWeight": "bold"}),
        html.Th("ChartType", style={"minWidth": "60px", "textAlign": "center", "backgroundColor": "#e9ecef", "fontWeight": "bold"}),
        html.Th("CHI", style={"minWidth": "70px", "textAlign": "center", "backgroundColor": "#e9ecef", "fontWeight": "bold"}),
        html.Th("Bimode", style={"minWidth": "70px", "textAlign": "center", "backgroundColor": "#e9ecef", "fontWeight": "bold"}),
        html.Th("PartRisk", style={"minWidth": "70px", "textAlign": "center", "backgroundColor": "#e9ecef", "fontWeight": "bold"}),
        html.Th("Kshift", style={"minWidth": "90px", "textAlign": "center", "backgroundColor": "#e9ecef", "fontWeight": "bold"}),
        html.Th("Zombie", style={"minWidth": "70px", "textAlign": "center", "backgroundColor": "#e9ecef", "fontWeight": "bold"}),
        html.Th("備註", style={"minWidth": "70px", "maxWidth": "90px", "textAlign": "center", "backgroundColor": "#e9ecef", "fontWeight": "bold"}),
        html.Th("更新時間", style={"minWidth": "140px", "textAlign": "center", "whiteSpace": "nowrap", "backgroundColor": "#e9ecef", "fontWeight": "bold"}),
        html.Th("更新人", style={"minWidth": "80px", "textAlign": "center", "backgroundColor": "#e9ecef", "fontWeight": "bold"})
    ]))

    def dropdown_cell(row, col_name):
        val = row[col_name]
        options = []
        if val != "":
            unique_options = {val, "Finish", "Waive"}
            options = [{"label": opt, "value": opt} for opt in sorted(list(unique_options))]

        bg = ""
        fg = ""
        if val == "Waive":
            bg = "#FFD6E0" # 粉紅
            fg = "#C2185B"
        elif val == "Finish":
            bg = "#D4EDDA" # 綠色
            fg = "#155724"
        elif val != "":
            bg = "#FFF3CD" # 黃色
            fg = "#856404"
        else:
            bg = "#E9ECEF" # 灰色
            fg = "#6C757D"

        return dcc.Dropdown(
            id={"type": "hl-dropdown", "row": row['KpiDefID'], "col": col_name},
            options=options,
            value=val if val != "" else None,
            clearable=False,
            disabled=(val == ""),
            style={
                "backgroundColor": bg,
                "color": fg,
                "minWidth": "65px",
                "fontSize": "13px",
                "border": "1px solid #ced4da", # 調整邊框顏色
                "borderRadius": "0.25rem" # 圓角
            },
            className="dbc-dropdown"
        )

    # 產生rows
    rows = []
    for i, row in df.iterrows():
        reply_btn = dbc.Button(
            "回覆",
            id={"type": "reply-btn", "row": row['KpiDefID']},
            color="info",
            size="sm",
            className="me-1"
        )
        # 判斷有無回覆
        has_reply = feedback_status.get(row['KpiDefID'], False)
        reply_icon = html.I(className="fas fa-check-circle", style={"color": "#28a745", "fontSize": "18px"}) if has_reply else ""
        rows.append(html.Tr([
            html.Td(row['KpiDefID'], style={"textAlign": "center", "verticalAlign": "middle", "height": "48px"}),
            html.Td(row['FAB'], style={"textAlign": "center", "verticalAlign": "middle", "height": "48px"}),
            html.Td(reply_icon, style={"textAlign": "center", "verticalAlign": "middle", "height": "48px"}),
            html.Td(reply_btn, style={"textAlign": "center", "verticalAlign": "middle", "height": "48px"}),
            html.Td(row['Module'] if 'Module' in row else '', style={"textAlign": "center", "verticalAlign": "middle", "height": "48px"}),
            html.Td(row['GroupName'], style={"textAlign": "center", "verticalAlign": "middle", "height": "48px"}),
            html.Td(row['ChartName'], style={"textAlign": "center", "verticalAlign": "middle", "height": "48px"}),
            html.Td(row['ChartType'], style={"textAlign": "center", "verticalAlign": "middle", "height": "48px"}),
            html.Td(dropdown_cell(row, 'HL_CHI'), style={"textAlign": "center", "height": "48px"}),
            html.Td(dropdown_cell(row, 'HL_Bimode'), style={"textAlign": "center", "height": "48px"}),
            html.Td(dropdown_cell(row, 'HL_PartRisk'), style={"textAlign": "center", "height": "48px"}),
            html.Td(dropdown_cell(row, 'HL_Kshift'), style={"textAlign": "center", "height": "48px", "minWidth": "130px"}),
            html.Td(dropdown_cell(row, 'HL_Zombie'), style={"textAlign": "center", "height": "48px"}),
            html.Td(dcc.Input(
                id={'type': 'remark-input', 'row': row['KpiDefID']},
                value=row['Remark'],
                type='text',
                debounce=True,
                style={'width': '100%', 'fontSize': '11px', 'padding': '4px 6px', 'minWidth': '60px', 'maxWidth': '90px', 'overflow': 'hidden', 'textOverflow': 'ellipsis', 'border': '1px solid #ced4da', 'borderRadius': '0.25rem'}
            ), style={"textAlign": "center", "minWidth": "70px", "maxWidth": "90px", "overflow": "hidden", "textOverflow": "ellipsis", "verticalAlign": "middle", "height": "48px"}),
            html.Td(row['LastUpdatedDate'], style={"minWidth": "140px", "textAlign": "center", "whiteSpace": "nowrap", "fontSize": "13px", "verticalAlign": "middle", "color": "#6c757d", "height": "48px"}),
            html.Td(row['UpdatedBy'], style={"minWidth": "80px", "textAlign": "center", "fontSize": "13px", "verticalAlign": "middle", "color": "#6c757d", "height": "48px"}),
        ]))

    return dashboard, html.Div([
        html.Table([
            header,
            html.Tbody(rows)
        ], className="table table-bordered table-hover table-striped table-sm",
           style={"minWidth": "1200px", "fontSize": "13px", "borderCollapse": "separate", "borderSpacing": "0", "borderRadius": "0.25rem"}),
    ], style={
        "maxHeight": "calc(60vh - 40px)",  # 只佔螢幕高度約60%，再扣掉一點空間
        "overflowY": "auto",
        "overflowX": "auto",
        "border": "1px solid #e3e6e9",
        "borderRadius": "0.25rem",
        "marginBottom": "8px",
        "paddingBottom": "0",
        "height": "auto"
    }), module_options, groupname_options, chartname_options, charttype_options
# --- 儲存變更 callback ---
@callback(
    Output('save-output-message', 'children', allow_duplicate=True),
    Output('kpi-table-container', 'children', allow_duplicate=True),
    Output('module-filter', 'options', allow_duplicate=True),
    Output('groupname-filter', 'options', allow_duplicate=True),
    Output('chartname-filter', 'options', allow_duplicate=True),
    Output('charttype-filter', 'options', allow_duplicate=True),
    Input('btn-save-kpi', 'n_clicks'),
    State('fac-selector', 'value'),
    State('module-filter', 'value'),
    State({'type': 'hl-dropdown', 'row': ALL, 'col': ALL}, 'id'),
    State({'type': 'hl-dropdown', 'row': ALL, 'col': ALL}, 'value'),
    State({'type': 'remark-input', 'row': ALL}, 'id'),
    State({'type': 'remark-input', 'row': ALL}, 'value'),
    State('login-user', 'data'),
    State('groupname-filter', 'value'),
    State('chartname-filter', 'value'),
    State('charttype-filter', 'value'),
    prevent_initial_call=True
)
def save_kpi_status(n_clicks, selected_fab, module, hl_dropdown_ids, hl_dropdown_values, remark_ids, remark_values, user, groupname, chartname, charttype):
    ctx = callback_context
    if not ctx.triggered or ctx.triggered[0]['prop_id'] != 'btn-save-kpi.n_clicks':
        raise PreventUpdate
    if not n_clicks or not user:
        raise PreventUpdate
    if not selected_fab:
        dashboard, table, mopt, gopt, copt, topt = render_kpi_table(selected_fab, module, groupname, chartname, charttype, user)
        return dbc.Alert("請選擇一個FAB以儲存資料。", color="warning", dismissable=True, duration=3000), table, mopt, gopt, copt, topt

    updated_count = 0
    now = datetime.datetime.now()
    with engine.connect() as connection:
        current_def_query = f"""
        SELECT KpiDefID, HL_CHI, HL_Bimode, HL_PartRisk, HL_Kshift, HL_Zombie
        FROM kpi_definitions
        WHERE FAB = '{selected_fab}'
        """
        current_defs_df = pd.read_sql(current_def_query, connection)
        current_defs_df.set_index('KpiDefID', inplace=True)
        current_status_query = f"""
        SELECT KpiDefID, Remark
        FROM kpi_status
        WHERE KpiDefID IN (SELECT KpiDefID FROM kpi_definitions WHERE FAB = '{selected_fab}')
        """
        current_status_df = pd.read_sql(current_status_query, connection)
        current_status_df.set_index('KpiDefID', inplace=True)
        if hl_dropdown_ids and hl_dropdown_values:
            for hl_id, hl_value in zip(hl_dropdown_ids, hl_dropdown_values):
                kpi_id = hl_id['row']
                col = hl_id['col']
                if kpi_id not in current_defs_df.index:
                    continue
                old_val = str(current_defs_df.loc[kpi_id, col]) if pd.notna(current_defs_df.loc[kpi_id, col]) else ""
                new_val = str(hl_value) if hl_value is not None else ""
                if old_val != new_val:
                    update_kpi_def_sql = sqlalchemy.text(f"UPDATE kpi_definitions SET {col} = :val WHERE KpiDefID = :kid")
                    connection.execute(update_kpi_def_sql, {"val": new_val, "kid": kpi_id})
                    update_kpi_status_sql = sqlalchemy.text("UPDATE kpi_status SET UpdatedBy = :by, LastUpdatedDate = :dt WHERE KpiDefID = :kid")
                    connection.execute(update_kpi_status_sql, {"by": user, "dt": now, "kid": kpi_id})
                    updated_count += 1
        if remark_ids and remark_values:
            for remark_id, remark_value in zip(remark_ids, remark_values):
                kpi_id = remark_id['row']
                if kpi_id not in current_status_df.index:
                    continue
                old_remark = str(current_status_df.loc[kpi_id, 'Remark']) if pd.notna(current_status_df.loc[kpi_id, 'Remark']) else ""
                new_remark = str(remark_value) if remark_value is not None else ""
                if old_remark != new_remark:
                    update_remark_sql = sqlalchemy.text("UPDATE kpi_status SET Remark = :remark_val, UpdatedBy = :by, LastUpdatedDate = :dt WHERE KpiDefID = :kid")
                    connection.execute(update_remark_sql, {"remark_val": new_remark, "by": user, "dt": now, "kid": kpi_id})
                    updated_count += 1
        connection.commit()
    # 儲存後重新刷新表格
    dashboard, table, mopt, gopt, copt, topt = render_kpi_table(selected_fab, module, groupname, chartname, charttype, user)
    if updated_count > 0:
        return dbc.Alert(f"已成功儲存 {updated_count} 筆變更！", color="success", dismissable=True, duration=3000), table, mopt, gopt, copt, topt
    else:
        return dbc.Alert("無異動，未儲存。", color="info", dismissable=True, duration=3000), table, mopt, gopt, copt, topt

# --- API 端點供 Power BI 使用 ---
@app.server.route('/powerbi-data')
def get_powerbi_data_api():
    df = get_kpi_data()
    # 移除定義欄位，因為 Power BI 可能只關心狀態和備註
    df = df.drop(columns=['HL_CHI', 'HL_Bimode', 'HL_PartRisk', 'HL_Kshift', 'HL_Zombie'], errors='ignore')
    # 統一日期格式為 W526
    def format_date_with_week(dt):
        if pd.isna(dt):
            return ''
        if isinstance(dt, str):
            try:
                dt = pd.to_datetime(dt)
            except Exception:
                return dt
        date_str = dt.strftime('%Y-%m-%d')
        iso = dt.isocalendar()
        # 週數前加上年份最後一碼
        week_str = f"W{str(iso.year)[-1]}{iso.week:02d}"
        return f"{date_str} ({week_str})"
    df['LastUpdatedDate'] = df['LastUpdatedDate'].apply(format_date_with_week)
    return df.to_json(orient="records", indent=4)

@app.server.route('/all-db-data')
def get_all_db_data_api():
    with engine.connect() as connection:
        df = pd.read_sql("""
            SELECT kd.*, ks.Remark, ks.LastUpdatedDate, ks.UpdatedBy
            FROM kpi_definitions kd
            LEFT JOIN kpi_status ks ON kd.KpiDefID = ks.KpiDefID
        """, connection)
    return Response(df.to_json(orient="records", force_ascii=False), mimetype='application/json; charset=utf-8')

# --- 檢視資料庫 Modal 控制與資料填充 ---
@callback(
    Output('db-modal', 'is_open'),
    Output('db-table', 'columns'),
    Output('db-table', 'data'),
    Input('btn-view-db', 'n_clicks'),
    Input('close-db-modal', 'n_clicks'),
    State('db-modal', 'is_open'),
    prevent_initial_call=True
)
def toggle_db_modal(view_clicks, close_click, is_open):
    ctx = callback_context
    if not ctx.triggered:
        raise PreventUpdate

    trigger_id = ctx.triggered_id

    if trigger_id == 'btn-view-db' and not is_open:
        with engine.connect() as connection:
            df = pd.read_sql("""
                SELECT kd.*, ks.Remark, ks.LastUpdatedDate, ks.UpdatedBy
                FROM kpi_definitions kd
                LEFT JOIN kpi_status ks ON kd.KpiDefID = ks.KpiDefID
            """, connection)
        if 'LastUpdatedDate' in df.columns:
            def format_date_with_week(dt):
                if pd.isna(dt):
                    return ''
                if isinstance(dt, str):
                    try:
                        dt = pd.to_datetime(dt)
                    except Exception:
                        return dt
                date_str = dt.strftime('%Y-%m-%d')
                iso = dt.isocalendar()
                # 週數前加上年份最後一碼
                week_str = f"W{str(iso.year)[-1]}{iso.week:02d}"
                return f"{date_str} ({week_str})"
            df['LastUpdatedDate'] = df['LastUpdatedDate'].apply(format_date_with_week)
        columns = [{"name": col, "id": col} for col in df.columns]
        data = df.to_dict('records')
        return True, columns, data
    elif trigger_id == 'close-db-modal' and is_open:
        return False, [], []
    raise PreventUpdate

# --- 將 data.csv 新增資料自動寫入 kpi_definitions（不重複）---
def sync_csv_to_db():
    csv_file_path = r"C:\\Users\\hsa00\\Desktop\\data.csv" # 請確保路徑正確
    try:
        df_csv = pd.read_csv(csv_file_path)
        if 'FAC' in df_csv.columns:
            df_csv = df_csv.rename(columns={'FAC': 'FAB'})
        # 欄位名稱同步更名
        df_csv = df_csv.rename(columns={
            'HL_K_Defined': 'HL_CHI',
            'HL_B_Defined': 'HL_Bimode',
            'HL_P_Defined': 'HL_PartRisk',
            'HL_S_Defined': 'HL_Kshift',
            'HL_W_Defined': 'HL_Zombie'
        })
        # 確保 HL 相關欄位存在於 CSV 數據中，如果沒有則添加空字串
        for col in ['HL_CHI', 'HL_Bimode', 'HL_PartRisk', 'HL_Kshift', 'HL_Zombie']:
            if col not in df_csv.columns:
                df_csv[col] = ''
            else:
                df_csv[col] = df_csv[col].fillna('') # 填充 NaN
        with engine.connect() as connection:
            # 獲取資料庫中 kpi_definitions 的欄位資訊
            db_cols_info = pd.read_sql("PRAGMA table_info(kpi_definitions)", connection)
            db_col_names = db_cols_info['name'].tolist()

            # 檢查並添加 CSV 中存在但資料庫中沒有的欄位
            for col in df_csv.columns:
                if col not in db_col_names:
                    alter_sql = f"ALTER TABLE kpi_definitions ADD COLUMN '{col}' TEXT"
                    connection.execute(sqlalchemy.text(alter_sql))
                    connection.commit()
                    print(f"已自動新增欄位: {col}")
                    db_col_names.append(col) # 更新已存在欄位列表

            # 只選擇資料庫中已存在的欄位進行插入
            df_csv_filtered = df_csv[[col for col in df_csv.columns if col in db_col_names]]

            insert_count = 0
            for _, row in df_csv_filtered.iterrows():
                key_cols = ['FAB', 'GroupName', 'ChartName', 'ChartType']  # 加入 ChartType
                if not all(col in row for col in key_cols):
                    print(f"警告：CSV 中缺少關鍵欄位 ({key_cols})，跳過此行：{row}")
                    continue

                # 檢查資料是否已存在
                cond = ' AND '.join([f"{k} = :{k}" for k in key_cols])
                params = {k: row[k] for k in key_cols}
                exist = connection.execute(sqlalchemy.text(f"SELECT 1 FROM kpi_definitions WHERE {cond}"), params).fetchone()

                if not exist:
                    insert_cols = df_csv_filtered.columns.tolist()
                    insert_sql = f"INSERT INTO kpi_definitions ({', '.join(insert_cols)}) VALUES ({', '.join([':' + col for col in insert_cols])})"
                    insert_params = {col: row[col] for col in insert_cols}
                    connection.execute(sqlalchemy.text(insert_sql), insert_params)
                    insert_count += 1
            connection.commit()
            print(f"新增 {insert_count} 筆資料到資料庫（不覆蓋舊資料）")
    except FileNotFoundError:
        print(f"警告：找不到初始 KPI 定義檔案 '{csv_file_path}'，跳過 CSV 同步。")
    except Exception as e:
        print(f"從 CSV 同步 KPI 定義時發生錯誤: {e}")

# --- KPI 回覆 Modal callback ---
@callback(
    Output("reply-modal", "is_open"),
    Output("reply-modal-body", "children"),
    Output('kpi-table-container', 'children', allow_duplicate=True),
    Input({"type": "reply-btn", "row": ALL}, "n_clicks"),
    Input("close-reply-modal", "n_clicks"),
    State("reply-modal", "is_open"),
    State('fac-selector', 'value'),
    State('module-filter', 'value'),
    State('groupname-filter', 'value'),
    State('chartname-filter', 'value'),
    State('charttype-filter', 'value'),
    State('login-user', 'data'),
    prevent_initial_call=True
)
def toggle_reply_modal(reply_btns, close_click, is_open, selected_fab, module, groupname, chartname, charttype, user):
    ctx = callback_context
    if not ctx.triggered or ctx.triggered[0]["value"] is None:
        raise PreventUpdate
    trigger = ctx.triggered[0]["prop_id"]
    # 關閉
    if "close-reply-modal" in trigger:
        # 關閉時refresh主表格
        dashboard, table, mopt, gopt, copt, topt = render_kpi_table(selected_fab, module, groupname, chartname, charttype, user)
        return False, dash.no_update, table
    # 開啟
    if "reply-btn" in trigger:
        # 取得觸發的 row id
        triggered_id = ctx.triggered_id  # 這是 dict
        row_id = None
        if isinstance(triggered_id, dict):
            row_id = triggered_id.get("row")
        if row_id is None:
            return is_open, dash.no_update
        df = get_kpi_data(fab=selected_fab)
        # row_id 可能是 int 或 str，需轉型比對
        row = df[df["KpiDefID"].astype(str) == str(row_id)]
        if row.empty:
            return is_open, dash.no_update
        row = row.iloc[0]
        kpi_cols = ["HL_CHI", "HL_Bimode", "HL_PartRisk", "HL_Kshift", "HL_Zombie"]
        # 新增：如果五個KPI值全為空，直接關閉modal
        if all((str(row[col]).strip() == "" or pd.isna(row[col])) for col in kpi_cols):
            return False, dash.no_update
        kpi_names = ["CHI", "Bimode", "PartRisk", "Kshift", "Zombie"]
        kpi_items = []
        for col, name in zip(kpi_cols, kpi_names):
            row_id = int(row['KpiDefID']) if not isinstance(row['KpiDefID'], int) else row['KpiDefID']
            feedback_id = {"type": "modal-feedback", "kpi": col, "row": row_id}
            action_id = {"type": "modal-action", "kpi": col, "row": row_id}
            value_str = str(row[col])
            value_style = {"minWidth": "120px"}
            input_disabled = False
            input_placeholder = "Module Feedback"
            action_placeholder = "Action"
            input_style = {"width": "100%", "fontSize": "13px"}
            action_style = {"width": "100%", "fontSize": "13px"}
            # 查詢最新 feedback/action
            latest_feedback = ""
            latest_action = ""
            with engine.connect() as connection:
                fb_row = connection.execute(
                    sqlalchemy.text("SELECT feedback, action FROM kpi_feedback WHERE KpiDefID = :kid AND KPICol = :col ORDER BY timestamp DESC LIMIT 1"),
                    {"kid": row_id, "col": col}
                ).fetchone()
                if fb_row:
                    latest_feedback = fb_row[0] if fb_row[0] is not None else ""
                    latest_action = fb_row[1] if fb_row[1] is not None else ""
            if value_str.strip() == "" or value_str.strip().lower() == "nan":
                input_disabled = True
                input_placeholder = ""
                action_placeholder = ""
                input_style["backgroundColor"] = "#f1f3f6"
                action_style["backgroundColor"] = "#f1f3f6"
            if value_str.strip().lower() == "finish":
                value_style.update({"color": "#198754", "fontWeight": "bold", "backgroundColor": "#e6f4ea"})  # 綠色
            elif value_str.strip().lower() == "waive":
                value_style.update({"color": "#dc3545", "fontWeight": "bold", "backgroundColor": "#fbeaea"})  # 紅色
            kpi_items.append(html.Tr([
                html.Td(name, style={"fontWeight": "bold", "minWidth": "80px"}),
                html.Td(value_str, style=value_style),
                html.Td(dcc.Input(id=feedback_id, type="text", placeholder=input_placeholder, style=input_style, disabled=input_disabled, value=latest_feedback)),
                html.Td(dcc.Input(id=action_id, type="text", placeholder=action_placeholder, style=action_style, disabled=input_disabled, value=latest_action))
            ]))
        # 新增：左側留空區，右側為原本內容
        # 組出 minio 圖片網址，檔名格式: GroupName_ChartName_ChartType.png
        group = str(row.get('GroupName', ''))
        chart = str(row.get('ChartName', ''))
        ctype = str(row.get('ChartType', ''))
        minio_filename = f"{group}_{chart}_{ctype}.png"
        minio_url = f"http://minio-server:9000/bucket-name/{minio_filename}"
        img_component = html.Img(src=minio_url, style={
            "maxWidth": "100%", "maxHeight": "350px", "objectFit": "contain"
        }, alt="KPI 圖片")
        # 新增：圖表資訊列
        info_row = html.Div([
            html.Span(f"Module: {row.get('Module', '')}", style={"marginRight": "24px", "fontWeight": "bold"}),
            html.Span(
                f"Chart: {row.get('GroupName', '')}@{row.get('ChartName', '')}@{row.get('ChartType', '')}",
                style={"fontWeight": "bold"}
            ),
        ], style={"marginBottom": "16px", "fontSize": "16px", "background": "#f8f9fa", "padding": "8px 16px", "borderRadius": "6px"})

        body = dbc.Row([
            dbc.Col([
                html.Div(img_component, id="reply-modal-img-area", style={
                    "width": "320px", "height": "370px", "background": "#f8f9fa", "border": "1px dashed #bbb", "borderRadius": "8px",
                    "display": "flex", "alignItems": "center", "justifyContent": "center", "color": "#bbb", "fontSize": "16px"
               
                })
            ], width=5),
            dbc.Col([
                info_row,
                html.Div([
                    html.Table([
                        html.Thead(html.Tr([
                            html.Th("KPI 名稱", style={"minWidth": "80px", "textAlign": "center", "backgroundColor": "#1976d2", "color": "#fff", "fontWeight": "bold", "fontSize": "15px", "borderTopLeftRadius": "8px"}),
                            html.Th("值", style={"minWidth": "120px", "textAlign": "center", "backgroundColor": "#1976d2", "color": "#fff", "fontWeight": "bold", "fontSize": "15px"}),
                            html.Th("Module Feedback", style={"minWidth": "120px", "textAlign": "center", "backgroundColor": "#1976d2", "color": "#fff", "fontWeight": "bold", "fontSize": "15px"}),
                            html.Th("Action", style={"minWidth": "120px", "textAlign": "center", "backgroundColor": "#1976d2", "color": "#fff", "fontWeight": "bold", "fontSize": "15px", "borderTopRightRadius": "8px"})
                        ])),
                        html.Tbody([
                            html.Tr(row.children, style={
                                "backgroundColor": "#f8fafd" if i%2==0 else "#f1f3f6",
                                "padding": "10px 14px",
                                "fontSize": "14px",
                                "verticalAlign": "middle"
                            }) for i, row in enumerate(kpi_items)
                        ])
                    ], className="table table-bordered table-hover table-striped table-sm", style={
                        "minWidth": "600px",
                        "fontSize": "14px",
                        "borderCollapse": "separate",
                        "borderSpacing": "0",
                        "borderRadius": "10px",
                        "overflow": "hidden",
                        "boxShadow": "0 2px 12px 0 rgba(0,0,0,0.08)",
                        "background": "#fff"
                    }),
                ], style={
                    "padding": "18px 18px 12px 18px",
                    "background": "#fff",
                    "borderRadius": "14px",
                    "boxShadow": "0 2px 16px 0 rgba(0,0,0,0.10)",
                    "marginBottom": "12px"
                }),
                html.Div(id="modal-save-msg", className="mt-2"),
            ], width=7)
        ], align="center", className="gy-2")
        # 將儲存按鈕固定在右下角
        # 儲存按鈕必須在 modal 內容內部，不能用 fixed，否則 dash callback 會失效
        # 不再自訂右下角儲存按鈕，僅回傳 body，保留 ModalFooter 的按鈕
        return True, body, dash.no_update
# --- Modal Feedback 儲存 callback ---
from dash import ALL
@callback(
    Output("modal-save-msg", "children"),
    Output("reply-modal", "is_open", allow_duplicate=True),
    Input("modal-save-btn", "n_clicks"),
    State({"type": "modal-feedback", "kpi": ALL, "row": ALL}, "id"),
    State({"type": "modal-feedback", "kpi": ALL, "row": ALL}, "value"),
    State({"type": "modal-action", "kpi": ALL, "row": ALL}, "id"),
    State({"type": "modal-action", "kpi": ALL, "row": ALL}, "value"),
    State('login-user', 'data'),
    prevent_initial_call=True
)
def save_modal_feedback(n_clicks, feedback_ids, feedback_values, action_ids, action_values, user):
    if not n_clicks or not user:
        return dash.no_update, dash.no_update
    import datetime
    now = datetime.datetime.now()
    count = 0
    with engine.connect() as connection:
        # 儲存所有 feedback/action，合併同一 KPICol 一起寫入
        for i, fid in enumerate(feedback_ids):
            val_fb = feedback_values[i] if feedback_values and i < len(feedback_values) else None
            # 找到對應的 action id/value
            # action_ids/action_values 與 feedback_ids 順序一致（同一 row/kpi）
            val_act = None
            for j, aid in enumerate(action_ids):
                if aid and fid and aid['row'] == fid['row'] and aid['kpi'] == fid['kpi']:
                    val_act = action_values[j] if action_values and j < len(action_values) else None
                    break
            if fid and (val_fb or val_act):
                connection.execute(
                    sqlalchemy.text("INSERT INTO kpi_feedback (KpiDefID, KPICol, feedback, action, user, timestamp) VALUES (:kid, :col, :fb, :act, :user, :ts)"),
                    {"kid": fid['row'], "col": fid['kpi'], "fb": val_fb, "act": val_act, "user": user, "ts": now}
                )
                count += 1
        connection.commit()
    msg = dbc.Alert(f"已成功儲存 {count} 筆回覆！", color="success", duration=3000, dismissable=True) if count > 0 else dbc.Alert("無資料可儲存。", color="info", duration=3000, dismissable=True)
    return msg, dash.no_update
    return is_open, dash.no_update
# --- 主程式區塊 ---
if __name__ == '__main__':
    print("啟動 Dash 伺服器...")
    initialize_db()
    sync_csv_to_db()
    app.run(debug=True, port=8051)
