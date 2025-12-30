"""
EEZO 2026 タスクダッシュボード
============================

北海道食材EC（EEZO）の年間タスクを可視化するガントチャートダッシュボード。

使用方法:
    python src/app.py

ブラウザでアクセス:
    http://127.0.0.1:8050
"""

import pandas as pd
from dash import Dash, html, dcc, callback, Output, Input, State
import dash_bootstrap_components as dbc
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import os

# =============================================================================
# 定数定義
# =============================================================================

DATA_PATH = "data/raw/eezo_2026_weekly_tasks.csv"
OUTPUT_DIR = "output"

# カテゴリ色設定
CATEGORY_COLORS = {
    "プラットフォーム実装": "#3498db",  # 青
    "UX動線": "#9b59b6",                # 紫
    "商品コンテンツ": "#e74c3c",        # 赤
    "集客販促": "#f39c12",              # オレンジ
    "データ活用": "#2ecc71",            # 緑
}

# 担当者色設定
ASSIGNEE_COLORS = {
    "松永": "#3498db",  # 青系
    "高山": "#e74c3c",  # 赤系
}

# =============================================================================
# データ読み込み
# =============================================================================

def load_data(filepath: str) -> pd.DataFrame:
    """
    CSVファイルを読み込み、日付をパースする。
    
    Args:
        filepath: CSVファイルのパス
        
    Returns:
        pd.DataFrame: 読み込んだデータフレーム
    """
    df = pd.read_csv(filepath, encoding="utf-8")
    df["開始日"] = pd.to_datetime(df["開始日"], format="%Y/%m/%d")
    df["終了日"] = pd.to_datetime(df["終了日"], format="%Y/%m/%d")
    
    # マイルストーンフラグを追加（★マーク付き）
    df["is_milestone"] = df["成果物/マイルストーン"].str.contains("★", na=False)
    
    # タスクIDを追加
    df["task_id"] = range(1, len(df) + 1)
    
    return df

# =============================================================================
# ガントチャート生成
# =============================================================================

def create_gantt_chart(
    df: pd.DataFrame,
    color_by: str = "カテゴリ",
    granularity: str = "week"
) -> go.Figure:
    """
    ガントチャートを生成する。
    
    Args:
        df: タスクデータフレーム
        color_by: 色分けの基準（"カテゴリ" or "担当者"）
        granularity: 時間粒度（"day", "week", "month"）
        
    Returns:
        go.Figure: Plotlyのガントチャート
    """
    color_map = CATEGORY_COLORS if color_by == "カテゴリ" else ASSIGNEE_COLORS
    
    fig = px.timeline(
        df,
        x_start="開始日",
        x_end="終了日",
        y="タスク",
        color=color_by,
        color_discrete_map=color_map,
        hover_data=["四半期", "週番号", "担当者", "カテゴリ", "成果物/マイルストーン"],
        title="EEZO 2026 タスクガントチャート"
    )
    
    # マイルストーン強調（★付きタスク）
    milestones = df[df["is_milestone"]]
    if not milestones.empty:
        fig.add_trace(go.Scatter(
            x=milestones["終了日"],
            y=milestones["タスク"],
            mode="markers",
            marker=dict(symbol="star", size=12, color="gold", line=dict(color="black", width=1)),
            name="★マイルストーン",
            hoverinfo="skip"
        ))
    
    # レイアウト調整
    fig.update_layout(
        height=max(400, len(df) * 25),
        xaxis_title="日付",
        yaxis_title="",
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(l=200, r=50, t=80, b=50),
    )
    
    # X軸の粒度設定
    if granularity == "day":
        fig.update_xaxes(dtick="D1", tickformat="%m/%d")
    elif granularity == "week":
        fig.update_xaxes(dtick="D7", tickformat="%m/%d")
    else:  # month
        fig.update_xaxes(dtick="M1", tickformat="%Y/%m")
    
    return fig

# =============================================================================
# Dashアプリケーション
# =============================================================================

# アプリ初期化
app = Dash(
    __name__,
    external_stylesheets=[dbc.themes.FLATLY],
    suppress_callback_exceptions=True
)
app.title = "EEZO 2026 タスクダッシュボード"

# データ読み込み
df = load_data(DATA_PATH)

# =============================================================================
# レイアウト
# =============================================================================

# フィルターカード
filter_card = dbc.Card([
    dbc.CardHeader("フィルター"),
    dbc.CardBody([
        dbc.Row([
            dbc.Col([
                html.Label("四半期"),
                dcc.Dropdown(
                    id="quarter-filter",
                    options=[{"label": q, "value": q} for q in df["四半期"].unique()],
                    value=df["四半期"].unique().tolist(),
                    multi=True,
                    placeholder="四半期を選択..."
                )
            ], md=3),
            dbc.Col([
                html.Label("担当者"),
                dcc.Dropdown(
                    id="assignee-filter",
                    options=[{"label": a, "value": a} for a in df["担当者"].unique()],
                    value=df["担当者"].unique().tolist(),
                    multi=True,
                    placeholder="担当者を選択..."
                )
            ], md=3),
            dbc.Col([
                html.Label("カテゴリ"),
                dcc.Dropdown(
                    id="category-filter",
                    options=[{"label": c, "value": c} for c in df["カテゴリ"].unique()],
                    value=df["カテゴリ"].unique().tolist(),
                    multi=True,
                    placeholder="カテゴリを選択..."
                )
            ], md=3),
            dbc.Col([
                html.Label("色分け"),
                dcc.RadioItems(
                    id="color-by",
                    options=[
                        {"label": "カテゴリ", "value": "カテゴリ"},
                        {"label": "担当者", "value": "担当者"}
                    ],
                    value="カテゴリ",
                    inline=True
                )
            ], md=3),
        ]),
        html.Hr(),
        dbc.Row([
            dbc.Col([
                html.Label("期間粒度"),
                dcc.RadioItems(
                    id="granularity",
                    options=[
                        {"label": "日", "value": "day"},
                        {"label": "週", "value": "week"},
                        {"label": "月", "value": "month"}
                    ],
                    value="week",
                    inline=True
                )
            ], md=4),
            dbc.Col([
                dbc.Button("Excelダウンロード", id="download-btn", color="success", className="mt-3"),
                dcc.Download(id="download-excel")
            ], md=4, className="text-end"),
        ])
    ])
], className="mb-3")

# サマリーカード
summary_card = dbc.Card([
    dbc.CardBody([
        dbc.Row([
            dbc.Col([html.H5("タスク数", className="text-muted"), html.H3(id="total-tasks")], md=3),
            dbc.Col([html.H5("期間", className="text-muted"), html.H4(id="date-range")], md=5),
            dbc.Col([html.H5("担当者別", className="text-muted"), html.Div(id="assignee-summary")], md=4),
        ])
    ])
], className="mb-3")

# メインレイアウト
app.layout = dbc.Container([
    html.H1("EEZO 2026 タスクダッシュボード", className="my-4 text-center"),
    filter_card,
    summary_card,
    dbc.Card([
        dbc.CardBody([
            dcc.Loading(
                dcc.Graph(id="gantt-chart", config={"displayModeBar": True}),
                type="circle"
            )
        ])
    ])
], fluid=True)

# =============================================================================
# コールバック
# =============================================================================

@callback(
    [Output("gantt-chart", "figure"),
     Output("total-tasks", "children"),
     Output("date-range", "children"),
     Output("assignee-summary", "children")],
    [Input("quarter-filter", "value"),
     Input("assignee-filter", "value"),
     Input("category-filter", "value"),
     Input("color-by", "value"),
     Input("granularity", "value")]
)
def update_dashboard(quarters, assignees, categories, color_by, granularity):
    """フィルター変更時にダッシュボードを更新"""
    
    # フィルタリング
    filtered_df = df[
        (df["四半期"].isin(quarters or [])) &
        (df["担当者"].isin(assignees or [])) &
        (df["カテゴリ"].isin(categories or []))
    ]
    
    # ガントチャート生成
    if filtered_df.empty:
        fig = go.Figure()
        fig.update_layout(title="データがありません", height=200)
    else:
        fig = create_gantt_chart(filtered_df, color_by, granularity)
    
    # サマリー計算
    total = len(filtered_df)
    
    if not filtered_df.empty:
        date_range = f"{filtered_df['開始日'].min().strftime('%Y/%m/%d')} 〜 {filtered_df['終了日'].max().strftime('%Y/%m/%d')}"
        assignee_counts = filtered_df.groupby("担当者").size()
        assignee_summary = " / ".join([f"{k}: {v}" for k, v in assignee_counts.items()])
    else:
        date_range = "-"
        assignee_summary = "-"
    
    return fig, f"{total} タスク", date_range, assignee_summary


@callback(
    Output("download-excel", "data"),
    Input("download-btn", "n_clicks"),
    [State("quarter-filter", "value"),
     State("assignee-filter", "value"),
     State("category-filter", "value")],
    prevent_initial_call=True
)
def download_excel(n_clicks, quarters, assignees, categories):
    """フィルター適用後のデータをExcelでダウンロード"""
    
    filtered_df = df[
        (df["四半期"].isin(quarters or [])) &
        (df["担当者"].isin(assignees or [])) &
        (df["カテゴリ"].isin(categories or []))
    ]
    
    # 出力用に整形
    export_df = filtered_df[[
        "四半期", "週番号", "開始日", "終了日", 
        "担当者", "カテゴリ", "タスク", "成果物/マイルストーン"
    ]].copy()
    
    export_df["開始日"] = export_df["開始日"].dt.strftime("%Y/%m/%d")
    export_df["終了日"] = export_df["終了日"].dt.strftime("%Y/%m/%d")
    
    filename = f"eezo_tasks_{datetime.now().strftime('%Y%m%d')}.xlsx"
    
    return dcc.send_data_frame(export_df.to_excel, filename, index=False)


# =============================================================================
# メイン実行
# =============================================================================

if __name__ == "__main__":
    print("=" * 50)
    print("EEZO 2026 タスクダッシュボード")
    print("=" * 50)
    print(f"データ: {len(df)} タスク")
    print(f"期間: {df['開始日'].min().strftime('%Y/%m/%d')} 〜 {df['終了日'].max().strftime('%Y/%m/%d')}")
    print("=" * 50)
    print("アクセス: http://127.0.0.1:8050")
    print("=" * 50)
    
    app.run(debug=True, host="0.0.0.0", port=8050)
