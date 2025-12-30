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
from dash import Dash, html, dcc, callback, Output, Input, State, no_update, ctx
import dash_bootstrap_components as dbc
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
from io import BytesIO
import os
from typing import Optional

# =============================================================================
# 定数定義
# =============================================================================

DATA_PATH = "data/raw/eezo_2026_weekly_tasks.csv"
OUTPUT_DIR = "output"

# カテゴリ色設定（CLAUDE.md準拠）
CATEGORY_COLORS: dict[str, str] = {
    "プラットフォーム実装": "#3498db",  # 青
    "UX動線": "#9b59b6",                # 紫
    "商品コンテンツ": "#e74c3c",        # 赤
    "集客販促": "#f39c12",              # オレンジ
    "データ活用": "#2ecc71",            # 緑
}

# 担当者色設定
ASSIGNEE_COLORS: dict[str, str] = {
    "松永": "#2980b9",  # 青系
    "高山": "#c0392b",  # 赤系
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

    # 期間（日数）を計算
    df["期間"] = (df["終了日"] - df["開始日"]).dt.days + 1

    return df


# =============================================================================
# ガントチャート生成
# =============================================================================

def create_gantt_chart(
    df: pd.DataFrame,
    color_by: str = "カテゴリ",
    granularity: str = "week",
    group_by: str = "none",
    show_today_line: bool = True
) -> go.Figure:
    """
    ガントチャートを生成する。

    Args:
        df: タスクデータフレーム
        color_by: 色分けの基準（"カテゴリ" or "担当者"）
        granularity: 時間粒度（"day", "week", "month"）
        group_by: グループ化（"none", "担当者", "カテゴリ"）
        show_today_line: 今日線を表示するか

    Returns:
        go.Figure: Plotlyのガントチャート
    """
    if df.empty:
        fig = go.Figure()
        fig.update_layout(
            title="データがありません",
            height=200,
            paper_bgcolor="white",
            plot_bgcolor="white"
        )
        return fig

    color_map = CATEGORY_COLORS if color_by == "カテゴリ" else ASSIGNEE_COLORS

    # グループ化に応じてY軸のラベルを調整
    df_chart = df.copy()
    if group_by == "担当者":
        df_chart["y_label"] = df_chart["担当者"] + " | " + df_chart["タスク"]
        df_chart = df_chart.sort_values(["担当者", "開始日"])
    elif group_by == "カテゴリ":
        df_chart["y_label"] = df_chart["カテゴリ"] + " | " + df_chart["タスク"]
        df_chart = df_chart.sort_values(["カテゴリ", "開始日"])
    else:
        df_chart["y_label"] = df_chart["タスク"]

    # ガントチャート作成
    fig = px.timeline(
        df_chart,
        x_start="開始日",
        x_end="終了日",
        y="y_label",
        color=color_by,
        color_discrete_map=color_map,
        custom_data=["四半期", "週番号", "担当者", "カテゴリ", "成果物/マイルストーン", "期間"],
    )

    # ホバーテンプレート設定
    fig.update_traces(
        hovertemplate=(
            "<b>%{y}</b><br>"
            "期間: %{x|%Y/%m/%d} 〜 %{customdata[5]}日間<br>"
            "四半期: %{customdata[0]} / %{customdata[1]}<br>"
            "担当者: %{customdata[2]}<br>"
            "カテゴリ: %{customdata[3]}<br>"
            "成果物: %{customdata[4]}<br>"
            "<extra></extra>"
        )
    )

    # マイルストーン強調（★付きタスク）
    milestones = df_chart[df_chart["is_milestone"]]
    if not milestones.empty:
        fig.add_trace(go.Scatter(
            x=milestones["終了日"],
            y=milestones["y_label"],
            mode="markers",
            marker=dict(
                symbol="star",
                size=14,
                color="gold",
                line=dict(color="black", width=1)
            ),
            name="★マイルストーン",
            hoverinfo="skip"
        ))

    # 今日線を追加
    if show_today_line:
        today = datetime.now()
        min_date = df_chart["開始日"].min()
        max_date = df_chart["終了日"].max()

        if min_date <= today <= max_date:
            fig.add_vline(
                x=today,
                line_dash="dash",
                line_color="red",
                line_width=2,
                annotation_text="今日",
                annotation_position="top"
            )

    # レイアウト調整
    chart_height = max(500, len(df_chart) * 28)
    fig.update_layout(
        height=chart_height,
        xaxis_title="日付",
        yaxis_title="",
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            bgcolor="rgba(255,255,255,0.8)"
        ),
        margin=dict(l=300, r=50, t=80, b=50),
        paper_bgcolor="white",
        plot_bgcolor="#fafafa",
        yaxis=dict(
            categoryorder="array",
            categoryarray=df_chart["y_label"].tolist()[::-1]
        ),
        font=dict(family="Noto Sans JP, sans-serif"),
    )

    # X軸の粒度設定
    if granularity == "day":
        fig.update_xaxes(dtick="D1", tickformat="%m/%d", tickangle=45)
    elif granularity == "week":
        fig.update_xaxes(dtick="D7", tickformat="%m/%d")
    else:  # month
        fig.update_xaxes(dtick="M1", tickformat="%Y/%m")

    # グリッド線
    fig.update_xaxes(showgrid=True, gridwidth=1, gridcolor="rgba(0,0,0,0.1)")
    fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor="rgba(0,0,0,0.05)")

    return fig


def create_excel_export(df: pd.DataFrame) -> BytesIO:
    """
    フィルター適用後のデータをExcelファイルとしてエクスポートする。
    タスク一覧とサマリー統計の2シート構成。

    Args:
        df: エクスポートするデータフレーム

    Returns:
        BytesIO: Excelファイルのバイトストリーム
    """
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # シート1: タスク一覧
        export_df = df[[
            "四半期", "週番号", "開始日", "終了日",
            "担当者", "カテゴリ", "タスク", "成果物/マイルストーン", "期間"
        ]].copy()
        export_df["開始日"] = export_df["開始日"].dt.strftime("%Y/%m/%d")
        export_df["終了日"] = export_df["終了日"].dt.strftime("%Y/%m/%d")
        export_df.to_excel(writer, sheet_name="タスク一覧", index=False)

        # シート2: サマリー統計
        summary_data = []

        # 全体統計
        summary_data.append({"項目": "総タスク数", "値": len(df)})
        summary_data.append({"項目": "期間開始", "値": df["開始日"].min().strftime("%Y/%m/%d")})
        summary_data.append({"項目": "期間終了", "値": df["終了日"].max().strftime("%Y/%m/%d")})
        summary_data.append({"項目": "マイルストーン数", "値": df["is_milestone"].sum()})
        summary_data.append({"項目": "", "値": ""})

        # 担当者別
        summary_data.append({"項目": "【担当者別タスク数】", "値": ""})
        for assignee, count in df.groupby("担当者").size().items():
            summary_data.append({"項目": f"  {assignee}", "値": count})
        summary_data.append({"項目": "", "値": ""})

        # カテゴリ別
        summary_data.append({"項目": "【カテゴリ別タスク数】", "値": ""})
        for category, count in df.groupby("カテゴリ").size().items():
            summary_data.append({"項目": f"  {category}", "値": count})
        summary_data.append({"項目": "", "値": ""})

        # 四半期別
        summary_data.append({"項目": "【四半期別タスク数】", "値": ""})
        for quarter, count in df.groupby("四半期").size().items():
            summary_data.append({"項目": f"  {quarter}", "値": count})

        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name="サマリー統計", index=False)

    output.seek(0)
    return output


# =============================================================================
# Dashアプリケーション
# =============================================================================

# アプリ初期化
app = Dash(
    __name__,
    external_stylesheets=[
        dbc.themes.FLATLY,
        "https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&display=swap"
    ],
    suppress_callback_exceptions=True
)
app.title = "EEZO 2026 タスクダッシュボード"

# データ読み込み
df = load_data(DATA_PATH)

# 日付範囲の計算
min_date = df["開始日"].min()
max_date = df["終了日"].max()
date_range_days = (max_date - min_date).days


# =============================================================================
# レイアウト
# =============================================================================

# ヘッダー
header = dbc.Navbar(
    dbc.Container([
        dbc.Row([
            dbc.Col([
                html.H3("EEZO 2026 タスクダッシュボード", className="text-white mb-0"),
                html.Small("北海道食材EC 年間プロジェクト管理", className="text-light")
            ]),
        ], align="center", className="flex-grow-1"),
        dbc.Row([
            dbc.Col([
                dbc.Button(
                    [html.I(className="fas fa-file-excel me-2"), "Excel出力"],
                    id="download-btn",
                    color="success",
                    className="me-2"
                ),
                dcc.Download(id="download-excel")
            ])
        ], align="center"),
    ], fluid=True),
    color="primary",
    dark=True,
    className="mb-3"
)

# フィルターカード
filter_card = dbc.Card([
    dbc.CardHeader([
        html.I(className="fas fa-filter me-2"),
        "フィルター・表示設定"
    ], className="fw-bold"),
    dbc.CardBody([
        # 第1行: フィルター
        dbc.Row([
            dbc.Col([
                dbc.Label("四半期", className="fw-bold text-muted small"),
                dcc.Dropdown(
                    id="quarter-filter",
                    options=[{"label": q, "value": q} for q in sorted(df["四半期"].unique())],
                    value=sorted(df["四半期"].unique().tolist()),
                    multi=True,
                    placeholder="四半期を選択..."
                )
            ], md=3),
            dbc.Col([
                dbc.Label("担当者", className="fw-bold text-muted small"),
                dcc.Dropdown(
                    id="assignee-filter",
                    options=[{"label": a, "value": a} for a in df["担当者"].unique()],
                    value=df["担当者"].unique().tolist(),
                    multi=True,
                    placeholder="担当者を選択..."
                )
            ], md=3),
            dbc.Col([
                dbc.Label("カテゴリ", className="fw-bold text-muted small"),
                dcc.Dropdown(
                    id="category-filter",
                    options=[{"label": c, "value": c} for c in df["カテゴリ"].unique()],
                    value=df["カテゴリ"].unique().tolist(),
                    multi=True,
                    placeholder="カテゴリを選択..."
                )
            ], md=6),
        ], className="mb-3"),

        # 第2行: 表示設定
        dbc.Row([
            dbc.Col([
                dbc.Label("期間粒度", className="fw-bold text-muted small"),
                dbc.RadioItems(
                    id="granularity",
                    options=[
                        {"label": "日", "value": "day"},
                        {"label": "週", "value": "week"},
                        {"label": "月", "value": "month"}
                    ],
                    value="week",
                    inline=True,
                    className="mt-1"
                )
            ], md=2),
            dbc.Col([
                dbc.Label("グループ化", className="fw-bold text-muted small"),
                dbc.RadioItems(
                    id="group-by",
                    options=[
                        {"label": "なし", "value": "none"},
                        {"label": "担当者別", "value": "担当者"},
                        {"label": "カテゴリ別", "value": "カテゴリ"}
                    ],
                    value="none",
                    inline=True,
                    className="mt-1"
                )
            ], md=3),
            dbc.Col([
                dbc.Label("ソート順", className="fw-bold text-muted small"),
                dbc.RadioItems(
                    id="sort-by",
                    options=[
                        {"label": "開始日", "value": "開始日"},
                        {"label": "担当者", "value": "担当者"},
                        {"label": "カテゴリ", "value": "カテゴリ"}
                    ],
                    value="開始日",
                    inline=True,
                    className="mt-1"
                )
            ], md=3),
            dbc.Col([
                dbc.Label("色分け", className="fw-bold text-muted small"),
                dbc.RadioItems(
                    id="color-by",
                    options=[
                        {"label": "カテゴリ", "value": "カテゴリ"},
                        {"label": "担当者", "value": "担当者"}
                    ],
                    value="カテゴリ",
                    inline=True,
                    className="mt-1"
                )
            ], md=2),
            dbc.Col([
                dbc.Label("今日線", className="fw-bold text-muted small"),
                dbc.Checklist(
                    id="show-today-line",
                    options=[{"label": "表示", "value": True}],
                    value=[True],
                    inline=True,
                    className="mt-1"
                )
            ], md=2),
        ], className="mb-3"),

        # 第3行: 日付範囲スライダー
        dbc.Row([
            dbc.Col([
                dbc.Label("日付範囲", className="fw-bold text-muted small"),
                dcc.RangeSlider(
                    id="date-range-slider",
                    min=0,
                    max=date_range_days,
                    step=7,
                    value=[0, date_range_days],
                    marks={
                        0: min_date.strftime("%Y/%m"),
                        date_range_days // 4: (min_date + pd.Timedelta(days=date_range_days // 4)).strftime("%Y/%m"),
                        date_range_days // 2: (min_date + pd.Timedelta(days=date_range_days // 2)).strftime("%Y/%m"),
                        date_range_days * 3 // 4: (min_date + pd.Timedelta(days=date_range_days * 3 // 4)).strftime("%Y/%m"),
                        date_range_days: max_date.strftime("%Y/%m"),
                    },
                    tooltip={"placement": "bottom", "always_visible": False}
                )
            ], md=12),
        ]),
    ])
], className="mb-3 shadow-sm")

# サマリーカード
summary_card = dbc.Card([
    dbc.CardBody([
        dbc.Row([
            dbc.Col([
                html.Div([
                    html.Span("📊 ", style={"fontSize": "1.5em"}),
                    html.Span("タスク数", className="text-muted small d-block"),
                    html.Span(id="total-tasks", className="h4 fw-bold text-primary")
                ], className="text-center")
            ], md=2),
            dbc.Col([
                html.Div([
                    html.Span("📅 ", style={"fontSize": "1.5em"}),
                    html.Span("期間", className="text-muted small d-block"),
                    html.Span(id="date-range", className="h6")
                ], className="text-center")
            ], md=3),
            dbc.Col([
                html.Div([
                    html.Span("👥 ", style={"fontSize": "1.5em"}),
                    html.Span("担当者別", className="text-muted small d-block"),
                    html.Div(id="assignee-summary")
                ], className="text-center")
            ], md=3),
            dbc.Col([
                html.Div([
                    html.Span("📁 ", style={"fontSize": "1.5em"}),
                    html.Span("カテゴリ別", className="text-muted small d-block"),
                    html.Div(id="category-summary")
                ], className="text-center")
            ], md=4),
        ], align="center")
    ])
], className="mb-3 shadow-sm")

# タスク並び替え用ストア（セッション内保持）
task_order_store = dcc.Store(id="task-order-store", storage_type="session")

# メインレイアウト
app.layout = dbc.Container([
    task_order_store,
    header,
    filter_card,
    summary_card,
    dbc.Card([
        dbc.CardBody([
            dcc.Loading(
                dcc.Graph(
                    id="gantt-chart",
                    config={
                        "displayModeBar": True,
                        "displaylogo": False,
                        "modeBarButtonsToRemove": ["lasso2d", "select2d"],
                        "toImageButtonOptions": {
                            "format": "png",
                            "filename": "eezo_gantt_chart",
                            "height": 1200,
                            "width": 1800,
                            "scale": 2
                        }
                    }
                ),
                type="circle",
                color="#3498db"
            )
        ])
    ], className="shadow-sm")
], fluid=True, className="pb-4")


# =============================================================================
# コールバック
# =============================================================================

@callback(
    [Output("gantt-chart", "figure"),
     Output("total-tasks", "children"),
     Output("date-range", "children"),
     Output("assignee-summary", "children"),
     Output("category-summary", "children")],
    [Input("quarter-filter", "value"),
     Input("assignee-filter", "value"),
     Input("category-filter", "value"),
     Input("color-by", "value"),
     Input("granularity", "value"),
     Input("group-by", "value"),
     Input("sort-by", "value"),
     Input("date-range-slider", "value"),
     Input("show-today-line", "value")]
)
def update_dashboard(
    quarters: list[str],
    assignees: list[str],
    categories: list[str],
    color_by: str,
    granularity: str,
    group_by: str,
    sort_by: str,
    date_range: list[int],
    show_today_line: list
) -> tuple:
    """フィルター変更時にダッシュボードを更新"""

    # 日付範囲の計算
    start_date = min_date + pd.Timedelta(days=date_range[0])
    end_date = min_date + pd.Timedelta(days=date_range[1])

    # フィルタリング
    filtered_df = df[
        (df["四半期"].isin(quarters or [])) &
        (df["担当者"].isin(assignees or [])) &
        (df["カテゴリ"].isin(categories or [])) &
        (df["開始日"] >= start_date) &
        (df["終了日"] <= end_date)
    ].copy()

    # ソート
    if sort_by == "開始日":
        filtered_df = filtered_df.sort_values(["開始日", "担当者"])
    elif sort_by == "担当者":
        filtered_df = filtered_df.sort_values(["担当者", "開始日"])
    else:  # カテゴリ
        filtered_df = filtered_df.sort_values(["カテゴリ", "開始日"])

    # ガントチャート生成
    show_today = True in (show_today_line or [])
    fig = create_gantt_chart(
        filtered_df,
        color_by=color_by,
        granularity=granularity,
        group_by=group_by,
        show_today_line=show_today
    )

    # サマリー計算
    total = len(filtered_df)

    if not filtered_df.empty:
        date_range_str = f"{filtered_df['開始日'].min().strftime('%Y/%m/%d')} 〜 {filtered_df['終了日'].max().strftime('%Y/%m/%d')}"

        # 担当者別サマリー
        assignee_counts = filtered_df.groupby("担当者").size()
        assignee_badges = [
            dbc.Badge(
                f"{k}: {v}",
                color="primary" if k == "松永" else "danger",
                className="me-1"
            )
            for k, v in assignee_counts.items()
        ]

        # カテゴリ別サマリー
        category_counts = filtered_df.groupby("カテゴリ").size()
        category_badges = [
            dbc.Badge(
                f"{k[:4]}…: {v}" if len(k) > 5 else f"{k}: {v}",
                style={"backgroundColor": CATEGORY_COLORS.get(k, "#999")},
                className="me-1 mb-1"
            )
            for k, v in category_counts.items()
        ]
    else:
        date_range_str = "-"
        assignee_badges = "-"
        category_badges = "-"

    return (
        fig,
        f"{total}",
        date_range_str,
        assignee_badges,
        category_badges
    )


@callback(
    Output("download-excel", "data"),
    Input("download-btn", "n_clicks"),
    [State("quarter-filter", "value"),
     State("assignee-filter", "value"),
     State("category-filter", "value"),
     State("date-range-slider", "value")],
    prevent_initial_call=True
)
def download_excel(
    n_clicks: int,
    quarters: list[str],
    assignees: list[str],
    categories: list[str],
    date_range: list[int]
) -> dict:
    """フィルター適用後のデータをExcelでダウンロード"""

    # 日付範囲の計算
    start_date = min_date + pd.Timedelta(days=date_range[0])
    end_date = min_date + pd.Timedelta(days=date_range[1])

    filtered_df = df[
        (df["四半期"].isin(quarters or [])) &
        (df["担当者"].isin(assignees or [])) &
        (df["カテゴリ"].isin(categories or [])) &
        (df["開始日"] >= start_date) &
        (df["終了日"] <= end_date)
    ]

    excel_data = create_excel_export(filtered_df)
    filename = f"eezo_tasks_{datetime.now().strftime('%Y%m%d')}.xlsx"

    return dcc.send_bytes(excel_data.getvalue(), filename)


# =============================================================================
# メイン実行
# =============================================================================

if __name__ == "__main__":
    # 出力ディレクトリの作成
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print("=" * 60)
    print("  EEZO 2026 タスクダッシュボード")
    print("=" * 60)
    print(f"  データ: {len(df)} タスク")
    print(f"  期間: {df['開始日'].min().strftime('%Y/%m/%d')} 〜 {df['終了日'].max().strftime('%Y/%m/%d')}")
    print(f"  担当者: {', '.join(df['担当者'].unique())}")
    print(f"  カテゴリ: {len(df['カテゴリ'].unique())} 種類")
    print("=" * 60)
    print("  アクセス: http://127.0.0.1:8050")
    print("=" * 60)

    app.run(debug=True, host="0.0.0.0", port=8050)
