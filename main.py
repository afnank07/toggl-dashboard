import pandas as pd
import plotly.graph_objects as go
from dash import Dash, dcc, html, Input, Output
import dash_bootstrap_components as dbc
import datetime
import glob

def find_data_start_row(file_path):
    preview = pd.read_excel(file_path, nrows=30, header=None)
    for idx, row in preview.iterrows():
        row_str = ' '.join([str(val).lower() for val in row.values if pd.notna(val)])
        if 'description' in row_str and 'duration' in row_str and 'project' in row_str:
            return idx
    return 0

def load_data(file_path):
    header_row = find_data_start_row(file_path)
    df = pd.read_excel(file_path, header=header_row, engine="openpyxl")
    df.columns = [str(col).strip() for col in df.columns]
    df = df.dropna(how='all')
    df = df.reset_index(drop=True)
    df['Project'] = df['Project'].fillna('Unspecified')
    df['Tags'] = df['Tags'].fillna('Untagged')
    # Parse duration to minutes
    def parse_duration(duration_str):
        if pd.isna(duration_str) or str(duration_str).strip() in ['-', '', 'nan', '#', '@']:
            return 0
        try:
            s = str(duration_str).strip()
            is_negative = s.startswith('-')
            s = s.replace('-', '').replace('days', '').strip()
            time_part = s.split()[-1] if ' ' in s else s
            h, m, sec = 0, 0, 0
            if ':' in time_part:
                parts = time_part.split(':')
                if len(parts) == 3:
                    h, m, sec = map(int, parts)
                elif len(parts) == 2:
                    h, m = map(int, parts)
            minutes = h * 60 + m + sec / 60
            return -minutes if is_negative else minutes
        except Exception:
            return 0
    df['Duration_Minutes'] = df['Duration'].apply(parse_duration)
    # Parse date
    df['Start date'] = pd.to_datetime(df['Start date'], errors='coerce')
    return df

def get_project_pie(df):
    project_time = df.groupby('Project')['Duration_Minutes'].sum()
    project_hours = project_time / 60
    fig = go.Figure(go.Pie(
        labels=project_hours.index,
        values=project_hours.values,
        hole=0.4,
        hovertemplate='<b>%{label}</b><br>Hours: %{value:.2f}<br>Percent: %{percent}<br>Count: %{customdata}<extra></extra>',
        customdata=[df[df['Project']==proj].shape[0] for proj in project_hours.index],
        textinfo='label+percent',
        textfont_size=14
    ))
    fig.update_layout(title_text='Project Distribution', margin=dict(l=20, r=20, t=40, b=20), autosize=True)
    fig.update_layout(legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5))
    fig.update_traces(textposition='inside')
    return fig

def get_tags_pie(df, project):
    sub_df = df[df['Project'] == project]
    tag_time = sub_df.groupby('Tags')['Duration_Minutes'].sum()
    tag_hours = tag_time / 60
    if tag_hours.empty:
        fig = go.Figure()
        fig.add_annotation(text=f"No tag data for project '{project}'",
                          x=0.5, y=0.5, showarrow=False, font=dict(size=18))
        fig.update_layout(title_text=f'Tags Distribution for Project: {project}', margin=dict(l=20, r=20, t=40, b=20), autosize=True)
        fig.update_xaxes(visible=False)
        fig.update_yaxes(visible=False)
        return fig
    fig = go.Figure(go.Pie(
        labels=tag_hours.index,
        values=tag_hours.values,
        hole=0.4,
        hovertemplate='<b>%{label}</b><br>Hours: %{value:.2f}<br>Percent: %{percent}<br>Count: %{customdata}<extra></extra>',
        customdata=[sub_df[sub_df['Tags']==tag].shape[0] for tag in tag_hours.index],
        textinfo='label+percent',
        textfont_size=14
    ))
    fig.update_layout(title_text=f'Tags Distribution for Project: {project}', margin=dict(l=20, r=20, t=40, b=20), autosize=True)
    fig.update_layout(legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5))
    fig.update_traces(textposition='inside')
    return fig

def filter_df_by_date(df, start_date, end_date):
    mask = (df['Start date'] >= start_date) & (df['Start date'] <= end_date)
    return df[mask]

def serve_dashboard(df):
    app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
    min_date = df['Start date'].min()
    max_date = df['Start date'].max()
    projects = df['Project'].unique()
    app.layout = dbc.Container([
        html.H2("Time Tracking Pie Charts Dashboard", className="text-center mb-4"),
        dbc.Row([
            dbc.Col([
                html.Label("Select Date Range:"),
                dcc.DatePickerRange(
                    id='date-range',
                    min_date_allowed=min_date,
                    max_date_allowed=max_date,
                    start_date=min_date,
                    end_date=max_date,
                    style={"width": "100%"}
                ),
                html.Div(id='total-hours', className='mt-3', style={"fontSize": "20px", "fontWeight": "bold", "color": "#007bff"}),
            ], width=12, className="d-flex flex-column justify-content-center align-items-center", style={"paddingBottom": "20px"})
        ], className='mb-2'),
        dbc.Row([
            dbc.Col([
                html.Div([
                    dcc.Graph(id='project-pie', figure=get_project_pie(df), style={"width": "100%", "height": "100%"}),
                ], className="d-flex flex-column justify-content-center align-items-center", style={"height": "100%", "padding": "20px", "minHeight": "420px"}),
                html.Label("Select Project for Tags Pie Chart:", style={"marginTop": "30px"}),
                dcc.Dropdown(
                    id='project-dropdown',
                    options=[{'label': p, 'value': p} for p in projects],
                    value=projects[0],
                    clearable=False,
                    style={"width": "100%", "minWidth": "300px", "fontSize": "16px", "marginBottom": "20px"}
                ),
                html.Div([
                    dcc.Graph(id='tags-pie', figure=get_tags_pie(df, projects[0]), style={"width": "100%", "height": "100%"})
                ], className="d-flex flex-column justify-content-center align-items-center", style={"height": "100%", "padding": "20px", "minHeight": "420px"}),
            ], width=8, className="mx-auto")
        ], className="align-items-center justify-content-center", style={"minHeight": "900px"})
    ], fluid=True)

    @app.callback(
        Output('project-pie', 'figure'),
        Output('tags-pie', 'figure'),
        Output('total-hours', 'children'),
        Input('date-range', 'start_date'),
        Input('date-range', 'end_date'),
        Input('project-dropdown', 'value')
    )
    def update_pies_and_hours(start_date, end_date, selected_project):
        dff = filter_df_by_date(df, pd.to_datetime(start_date), pd.to_datetime(end_date))
        proj_fig = get_project_pie(dff)
        tags_fig = get_tags_pie(dff, selected_project)
        total_hours = dff['Duration_Minutes'].sum() / 60
        total_hours_text = f"Total Hours in Selected Range: {total_hours:.2f}"
        return proj_fig, tags_fig, total_hours_text

    app.run(debug=True)

if __name__ == "__main__":
    excel_files = glob.glob("./files/current/*.xlsx")
    
    if not excel_files:
        raise FileNotFoundError("No .xlsx files found in ./files/ directory")
    
    # Pick the first one (or adjust logic if you want latest by date, etc.)
    file_path = excel_files[0]
    # file_path = "./files/TogglTrack_Report_Detailed_report__from_08_28_2025_to_09_15_2025_.xlsx"
    file_path = "./files/current/TogglTrack_Report_Detailed_report__from_08_27_2025_to_09_20_2025_.xlsx"
    df = load_data(file_path)
    serve_dashboard(df)
