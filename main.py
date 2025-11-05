import pandas as pd
import plotly.graph_objects as go
from dash import Dash, dcc, html, Input, Output
import dash_bootstrap_components as dbc
import datetime
import glob

def get_tags_pie_main(df):
    tag_time = df.groupby('Tags')['Duration_Minutes'].sum()
    tag_hours = tag_time / 60
    fig = go.Figure(go.Pie(
        labels=tag_hours.index,
        values=tag_hours.values,
        hole=0.4,
        hovertemplate='<b>%{label}</b><br>Hours: %{value:.2f}<br>Percent: %{percent}<br>Count: %{customdata}<extra></extra>',
        customdata=[df[df['Tags']==tag].shape[0] for tag in tag_hours.index],
        textinfo='label+percent',
        textfont_size=14
    ))
    fig.update_layout(title_text='Tags Distribution', margin=dict(l=20, r=20, t=40, b=20), autosize=True)
    fig.update_layout(legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.05, font=dict(size=10)))
    fig.update_traces(textposition='inside')
    return fig

def get_projects_pie_for_tag(df, tag):
    sub_df = df[df['Tags'] == tag]
    project_time = sub_df.groupby('Project')['Duration_Minutes'].sum()
    project_hours = project_time / 60
    if project_hours.empty:
        fig = go.Figure()
        fig.add_annotation(text=f"No project data for tag '{tag}'",
                          x=0.5, y=0.5, showarrow=False, font=dict(size=18))
        fig.update_layout(title_text=f'Project Distribution for Tag: {tag}', margin=dict(l=20, r=20, t=40, b=20), autosize=True)
        fig.update_xaxes(visible=False)
        fig.update_yaxes(visible=False)
        return fig
    fig = go.Figure(go.Pie(
        labels=project_hours.index,
        values=project_hours.values,
        hole=0.4,
        hovertemplate='<b>%{label}</b><br>Hours: %{value:.2f}<br>Percent: %{percent}<br>Count: %{customdata}<extra></extra>',
        customdata=[sub_df[sub_df['Project']==proj].shape[0] for proj in project_hours.index],
        textinfo='label+percent',
        textfont_size=14
    ))
    fig.update_layout(title_text=f'Project Distribution for Tag: {tag}', margin=dict(l=20, r=20, t=40, b=20), autosize=True)
    fig.update_layout(legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.05, font=dict(size=10)))
    fig.update_traces(textposition='inside')
    return fig



def load_data(file_path):
    df = pd.read_excel(file_path, header=0, engine="openpyxl")
    df.columns = [str(col).strip() for col in df.columns]
    df = df.dropna(how='all').reset_index(drop=True)
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
    fig.update_layout(legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.05, font=dict(size=10)))
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
    fig.update_layout(legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.05, font=dict(size=10)))
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
    tags = df['Tags'].unique()
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
                html.Div([
                    html.Div(id='total-hours', className='mt-3', style={"fontSize": "20px", "fontWeight": "bold", "color": "#007bff"}),
                    html.Div(id='total-days', className='mt-1', style={"fontSize": "16px", "fontWeight": "bold", "color": "#333"})
                ]),
            ], width=12, className="d-flex flex-column justify-content-center align-items-center", style={"paddingBottom": "20px"})
        ], className='mb-2'),
        dbc.Row([
            dbc.Col([
                html.Div([
                    html.Label("Select Main Distribution:"),
                    dcc.Dropdown(
                        id='main-distribution-dropdown',
                        options=[
                            {'label': 'Project Distribution', 'value': 'Project'},
                            {'label': 'Tags Distribution', 'value': 'Tags'}
                        ],
                        value='Project',
                        clearable=False,
                        style={"width": "100%", "minWidth": "300px", "fontSize": "16px", "marginBottom": "20px"}
                    ),
                    dcc.Graph(id='main-pie', style={"width": "100%", "height": "500px", "minHeight": "400px"}),
                ], className="d-flex flex-column justify-content-center align-items-center", style={"height": "100%", "padding": "20px", "minHeight": "420px"}),
                html.Label(id='breakdown-label', style={"marginTop": "30px"}),
                dcc.Dropdown(
                    id='breakdown-dropdown',
                    options=[{'label': p, 'value': p} for p in projects],
                    value=projects[0],
                    clearable=False,
                    style={"width": "100%", "minWidth": "300px", "fontSize": "16px", "marginBottom": "20px"}
                ),
                html.Div([
                    dcc.Graph(id='breakdown-pie', style={"width": "100%", "height": "500px", "minHeight": "400px"})
                ], className="d-flex flex-column justify-content-center align-items-center", style={"height": "100%", "padding": "20px", "minHeight": "420px"}),
            ], width=8, className="mx-auto")
        ], className="align-items-center justify-content-center", style={"minHeight": "900px"})
    ], fluid=True)

    @app.callback(
        Output('main-pie', 'figure'),
        Output('breakdown-pie', 'figure'),
        Output('breakdown-dropdown', 'options'),
        Output('breakdown-dropdown', 'value'),
        Output('breakdown-label', 'children'),
        Output('total-hours', 'children'),
        Output('total-days', 'children'),
        Input('date-range', 'start_date'),
        Input('date-range', 'end_date'),
        Input('main-distribution-dropdown', 'value'),
        Input('breakdown-dropdown', 'value')
    )
    def update_pies_and_breakdown(start_date, end_date, main_dist, breakdown_value):
        dff = filter_df_by_date(df, pd.to_datetime(start_date), pd.to_datetime(end_date))
        total_hours = dff['Duration_Minutes'].sum() / 60
        total_hours_text = f"Total Hours in Selected Range: {total_hours:.2f}"
        # Calculate total days selected
        try:
            start = pd.to_datetime(start_date)
            end = pd.to_datetime(end_date)
            total_days = (end - start).days + 1 if pd.notnull(start) and pd.notnull(end) else 0
        except Exception:
            total_days = 0
        total_days_text = f"Total days selected: {total_days}"
        if main_dist == 'Project':
            main_fig = get_project_pie(dff)
            breakdown_options = [{'label': p, 'value': p} for p in dff['Project'].unique()]
            if breakdown_value not in dff['Project'].unique():
                breakdown_value = dff['Project'].unique()[0] if len(dff['Project'].unique()) > 0 else None
            breakdown_fig = get_tags_pie(dff, breakdown_value) if breakdown_value else go.Figure()
            breakdown_label = "Select Project for Tags Pie Chart:"
        else:
            main_fig = get_tags_pie_main(dff)
            breakdown_options = [{'label': t, 'value': t} for t in dff['Tags'].unique()]
            if breakdown_value not in dff['Tags'].unique():
                breakdown_value = dff['Tags'].unique()[0] if len(dff['Tags'].unique()) > 0 else None
            breakdown_fig = get_projects_pie_for_tag(dff, breakdown_value) if breakdown_value else go.Figure()
            breakdown_label = "Select Tag for Project Pie Chart:"
        return main_fig, breakdown_fig, breakdown_options, breakdown_value, breakdown_label, total_hours_text, total_days_text

    app.run(debug=True)

if __name__ == "__main__":
    # Use the final_transformed.xlsx file as input
    file_path = "final_transformed.xlsx"
    df = load_data(file_path)
    serve_dashboard(df)
