import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime

# --- Data Loading and Preprocessing ---
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
    df['Start date'] = pd.to_datetime(df['Start date'], errors='coerce')
    return df

# --- Pie Chart Functions ---
def get_project_pie(df):
    project_time = df.groupby('Project')['Duration_Minutes'].sum()
    project_hours = project_time / 60
    fig = go.Figure(go.Pie(
        labels=project_hours.index,
        values=project_hours.values,
        hole=0.4,
        hovertemplate='<b>%{label}</b><br>Hours: %{value:.2f}<br>Percent: %{percent}<br>Count: %{customdata}<extra></extra>',
        customdata=[df[df['Project']==proj].shape[0] for proj in project_hours.index],
        textinfo='label+percent+value',
        texttemplate='%{label}<br>%{percent} (%{value:.2f}h)',
        textfont_size=14
    ))
    fig.update_layout(title_text='Project Distribution', margin=dict(l=20, r=20, t=40, b=20), autosize=True)
    fig.update_layout(legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.05, font=dict(size=10)))
    fig.update_traces(textposition='inside')
    return fig

def get_tags_pie_main(df):
    tag_time = df.groupby('Tags')['Duration_Minutes'].sum()
    tag_hours = tag_time / 60
    fig = go.Figure(go.Pie(
        labels=tag_hours.index,
        values=tag_hours.values,
        hole=0.4,
        hovertemplate='<b>%{label}</b><br>Hours: %{value:.2f}<br>Percent: %{percent}<br>Count: %{customdata}<extra></extra>',
        customdata=[df[df['Tags']==tag].shape[0] for tag in tag_hours.index],
        textinfo='label+percent+value',
        texttemplate='%{label}<br>%{percent} (%{value:.2f}h)',
        textfont_size=14
    ))
    fig.update_layout(title_text='Tags Distribution', margin=dict(l=20, r=20, t=40, b=20), autosize=True)
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
        textinfo='label+percent+value',
        texttemplate='%{label}<br>%{percent} (%{value:.2f}h)',
        textfont_size=14
    ))
    fig.update_layout(title_text=f'Tags Distribution for Project: {project}', margin=dict(l=20, r=20, t=40, b=20), autosize=True)
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
        textinfo='label+percent+value',
        texttemplate='%{label}<br>%{percent} (%{value:.2f}h)',
        textfont_size=14
    ))
    fig.update_layout(title_text=f'Project Distribution for Tag: {tag}', margin=dict(l=20, r=20, t=40, b=20), autosize=True)
    fig.update_layout(legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.05, font=dict(size=10)))
    fig.update_traces(textposition='inside')
    return fig

# --- Main Streamlit App ---
def main():
    st.set_page_config(page_title="Time Tracking Dashboard", layout="wide")
    st.title("Time Tracking Pie Charts Dashboard")
    file_path = "files\\output\\final_transformed.xlsx"
    df = load_data(file_path)
    min_date = df['Start date'].min()
    max_date = df['Start date'].max()
    # --- Sidebar Controls ---
    st.sidebar.header("Filters")
    start_date = st.sidebar.date_input("Start date", min_value=min_date, max_value=max_date, value=min_date)
    end_date = st.sidebar.date_input("End date", min_value=min_date, max_value=max_date, value=max_date)
    main_dist = st.sidebar.selectbox("Main Distribution", ["Project", "Tags"], index=0)
    day_filter = st.sidebar.radio(
        "Show data for:",
        ["Both Weekdays & Weekends", "Weekdays Only", "Weekends Only"],
        index=0
    )
    # --- Filter Data ---
    mask = (df['Start date'] >= pd.to_datetime(start_date)) & (df['Start date'] <= pd.to_datetime(end_date))
    dff = df[mask].copy()
    # Add weekday/weekend column
    dff['DayType'] = dff['Start date'].dt.dayofweek.apply(lambda x: 'Weekend' if x >= 5 else 'Weekday')
    if day_filter == "Weekdays Only":
        dff = dff[dff['DayType'] == 'Weekday']
    elif day_filter == "Weekends Only":
        dff = dff[dff['DayType'] == 'Weekend']
    # Recalculate total_days based on unique days in filtered data
    unique_days = dff['Start date'].dt.date.nunique()
    total_days = unique_days if unique_days > 0 else 1
    total_hours = dff['Duration_Minutes'].sum() / 60
    st.markdown(f"**Total Hours in Selected Range:** {total_hours:.2f}")
    st.markdown(f"**Total days selected:** {unique_days}")
    # --- Main Pie Chart and Daily Average Pie Chart ---
    col1, col2 = st.columns([2,2])
    with col1:
        if main_dist == "Project":
            st.subheader("Project Distribution")
            main_fig = get_project_pie(dff)
            st.plotly_chart(main_fig, use_container_width=True)
            breakdown_options = dff['Project'].dropna().unique().tolist()
            breakdown_label = "Select Project for Tags Pie Chart:"
        else:
            st.subheader("Tags Distribution")
            main_fig = get_tags_pie_main(dff)
            st.plotly_chart(main_fig, use_container_width=True)
            breakdown_options = dff['Tags'].dropna().unique().tolist()
            breakdown_label = "Select Tag for Project Pie Chart:"
    with col2:
        if breakdown_options:
            breakdown_value = st.selectbox(breakdown_label, breakdown_options, key="breakdown")
            if main_dist == "Project":
                breakdown_fig = get_tags_pie(dff, breakdown_value)
            else:
                breakdown_fig = get_projects_pie_for_tag(dff, breakdown_value)
            st.plotly_chart(breakdown_fig, use_container_width=True)
        else:
            st.info("No breakdown options available for the selected range.")

    # --- Daily Average Pie Chart ---
    st.markdown("---")
    st.subheader("Daily Average Hours per Activity")
    avg_group = st.radio("Group by", ["Project", "Tags"], horizontal=True, key="avg_group")
    # Calculate number of days in selected range (avoid division by zero)
    num_days = max(total_days, 1)
    if avg_group == "Project":
        avg_series = dff.groupby('Project')['Duration_Minutes'].sum() / 60 / num_days
        labels = avg_series.index
        title = "Daily Average Hours per Project"
    else:
        avg_series = dff.groupby('Tags')['Duration_Minutes'].sum() / 60 / num_days
        labels = avg_series.index
        title = "Daily Average Hours per Tag"
    avg_fig = go.Figure(go.Pie(
        labels=labels,
        values=avg_series.values,
        hole=0.4,
        hovertemplate='<b>%{label}</b><br>Avg Hours/Day: %{value:.2f}<br>Percent: %{percent}<extra></extra>',
        textinfo='label+percent+value',
        texttemplate='%{label}<br>%{percent} (%{value:.2f}h)',
        textfont_size=14
    ))
    avg_fig.update_layout(title_text=title, margin=dict(l=20, r=20, t=40, b=20), autosize=True)
    avg_fig.update_layout(legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.05, font=dict(size=10)))
    avg_fig.update_traces(textposition='inside')
    st.plotly_chart(avg_fig, use_container_width=True)

if __name__ == "__main__":
    main()
    # streamlit run main.py