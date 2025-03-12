import gspread
from oauth2client.service_account import ServiceAccountCredentials
from google.oauth2 import service_account
import smtplib
from email.mime.text import MIMEText
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import seaborn as sns
import numpy as np
import scipy.stats as stats
from wordcloud import WordCloud
import squarify
import plotly.graph_objects as go
import networkx as nx
from docx import Document
from io import BytesIO
import base64
import warnings
import io
import zipfile
import time
import google.generativeai as genai
import re 
import hashlib
import json
from PIL import Image

# STREAMLIT CSS AND CONFIGURATIONS ==============================================================================================================

logo = Image.open("insightfox_logo.jpeg")

st.set_page_config(
    page_title="InsightFox",
    page_icon=logo,
    layout="wide",  # or "wide" if you prefer
    initial_sidebar_state="auto"
)

st.set_option('client.showErrorDetails', False)

matplotlib.use('Agg')
warnings.filterwarnings("ignore", category=DeprecationWarning, message=".*pyplot.*")

st.markdown(
    """
    <style>
    section[data-testid="stMain"] > div[data-testid="stMainBlockContainer"] {
         padding-top: 0px;  # Remove padding completely
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Hide Streamlit style elements
hide_st_style = """
                <style>
                #MainMenu {visibility: hidden;}
                footer {visibility: hidden;}
                header {visibility: hidden;}
                </style>
                """
st.markdown(hide_st_style, unsafe_allow_html=True)

st.markdown("""
    <style>
    [data-testid="stTextArea"] {
        color: #FFFFFF;
    }
    </style>
    """, unsafe_allow_html=True)

# Set Montserrat font
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Montserrat', sans-serif;
    }
    </style>
    """, unsafe_allow_html=True)

# Change color of specific Streamlit elements
st.markdown("""
    <style>
    .st-emotion-cache-1o6s5t7 {
        color: #ababab !important;
    }
    </style>
    """, unsafe_allow_html=True)

st.markdown("""
    <style>
    .stExpander {
        background-color: #F3F7EC;
        border-radius: 10px;
    }
    
    .stExpander > details {
        background-color: #F3F7EC;
        border-radius: 10px;
    }
    
    .stExpander > details > summary {
        background-color: #F3F7EC;
        border-radius: 10px 10px 0 0;
        padding: 10px;
    }
    
    .stExpander > details > div {
        background-color: #F3F7EC;
        border-radius: 0 0 10px 10px;
        padding: 10px;
    }
    
    .stCheckbox {
        background-color: #F3F7EC;
        border-radius: 5px;
        padding: 5px;
    }
    </style>
    """, unsafe_allow_html=True)

st.markdown("""
    <style>
    .stButton > button {
        color: #F3F7EC;
        background-color: #bd5c34;
        border-radius: 10px;
    }
    </style>
    """, unsafe_allow_html=True)

st.markdown(
    """
    <style>
    .streamlit-expanderHeader {
        font-size: 20px;
    }
    .streamlit-expanderContent {
        max-height: 400px;
        overflow-y: auto;
    }
    </style>
    """,
    unsafe_allow_html=True
)

def get_base64_of_bin_file(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

def set_png_as_page_bg(png_file):
    bin_str = get_base64_of_bin_file(png_file)
    page_bg_img = '''
    <style>
    .stApp {
        background-image: url("data:image/png;base64,%s");
        background-size: cover;
    }
    </style>
    ''' % bin_str
    
    st.markdown(page_bg_img, unsafe_allow_html=True)

# Set the background image
set_png_as_page_bg("background.png")

data_str = {
    'City': ['Tokyo', 'Osaka', 'Kyoto', 'Yokohama', 'Nagoya', 'Sapporo', 'Kobe', 'Fukuoka'],
}

data_int = {
    'Population': [13929286, 2691185, 1464890, 3757630, 2332000, 1970895, 1545873, 1612392],
}

df_sample_str = pd.DataFrame(data_str)
df_sample_int = pd.DataFrame(data_int)
# ================================================ SESSION STATE ============================================================================================================

# Initialize session state
if "page" not in st.session_state:
    st.session_state.page = 0  # 0 for register, 1 for login

if "login" not in st.session_state:
    st.session_state.login = 0  # 0 for not logged in, 1 for logged in

jsonhuhuhaha = "1PlgbJOiKKmOMfn2UH34eNKhuWJRFw7iKveAzuQuKo_w"
url_huhuhaha = f'https://docs.google.com/spreadsheet/ccc?key={jsonhuhuhaha}&output=csv'

if 'credentials_huhuhaha' not in st.session_state:
    df_huhuhaha = pd.read_csv(url_huhuhaha)
    credentials_huhuhaha = df_huhuhaha.iloc[0, 0]
    st.session_state['credentials_huhuhaha'] = credentials_huhuhaha

# Retrieve credentials from session state
credentials_huhuhaha = st.session_state['credentials_huhuhaha']

# Parse the JSON string into a Python dictionary
credentials_dict = json.loads(credentials_huhuhaha)

# Create credentials from the dictionary
creds = service_account.Credentials.from_service_account_info(credentials_dict)

# Add the scopes
scoped_creds = creds.with_scopes(['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive'])

# Authorize the client
client = gspread.authorize(scoped_creds)

# Google Sheets key
sheetskey = "1oJVg7AyQKShgPfCnb_uxfpg1K9nwP0gftv0SILS8BgY"

# ================================================= LOGIN FUNCTIONS =================================================================================================================

def get_login_df():
    """Retrieves login data from Google Sheets."""
    url = f'https://docs.google.com/spreadsheet/ccc?key={sheetskey}&output=csv'
    df_login = pd.read_csv(url)
    return df_login

def send_email(email, message):
    """Sends an email notification."""
    try:
        sender_email = "insightfoxa@gmail.com"  # Replace with your email
        sender_password = "txsr udsn fhpe pfns"  # Replace with your APP PASSWORD

        msg = MIMEText(message)
        msg['Subject'] = 'Password Reminder'
        msg['From'] = sender_email
        msg['To'] = email

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, email, msg.as_string())

        st.success("Reminder email sent, and don't forget to check spam!")
    except Exception as e:
        st.error(f"Failed to send email: {e}")

# ============================================ CLEAN FUNCTIONS =======================================================================================================

def remove_outliers_iqr(df, column):
    """Removes outliers from a DataFrame column using the IQR method."""
    Q1 = df[column].quantile(0.25)
    Q3 = df[column].quantile(0.75)
    IQR = Q3 - Q1
    lower_bound = Q1 - 1.5 * IQR
    upper_bound = Q3 + 1.5 * IQR
    df_filtered = df[(df[column] >= lower_bound) & (df[column] <= upper_bound)]
    return df_filtered

# =============================================== EDA FUNCTIONS ====================================================================================================================

def get_filename_from_title(title):
    """Converts a plot title into a safe file name by replacing spaces and colons."""
    filename = title.replace(" ", "_").replace(":", "") + ".png"
    return filename

# ============================================= Plotting for INT Data ==========================================================================================================================================================================

def get_histogram_figures(df, theme="Blue"):
    """Generates histogram figures for all numerical columns.
    Returns a list of tuples: (filename, figure)."""
    try:
        themes = {
            "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"],
            "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"],
            "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"],
            "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"],
            "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"],
            "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],
            "Pastel": ["#F0E68C", "#E6E096", "#Dcd9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"]
        }
        if theme not in themes:
            st.write(f"Theme '{theme}' not found. Using Blue theme.")
            theme = "Blue"
        colors = themes[theme]
        numerical_cols = df.select_dtypes(include=['number']).columns
        figures = []
        for i, col in enumerate(numerical_cols):
            fig, ax = plt.subplots(figsize=(8, 6))
            sns.histplot(df[col], color=colors[i % len(colors)], kde=True, ax=ax)
            title_str = f"Distribution of {col}"
            ax.set_title(title_str)
            ax.set_xlabel(f"{col} Values")
            ax.set_ylabel("Frequency")
            ax.grid(True, linestyle='--', alpha=0.6)
            filename = get_filename_from_title(title_str)
            figures.append((filename, fig))
        return figures
    except Exception as e:
        print(f"Error generating histogram figures: {e}")
        return None

def get_boxplot_figures(df, theme="Blue"):
    """Generates boxplot figures for all numerical columns.
    Returns a list of tuples: (filename, figure)."""
    try:
        themes = {
            "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"],
            "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"],
            "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"],
            "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"],
            "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"],
            "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],
            "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"]
        }
        if theme not in themes:
            st.write(f"Theme '{theme}' not found. Using Blue theme.")
            theme = "Blue"
        colors = themes[theme]
        numerical_cols = df.select_dtypes(include=['number']).columns
        figures = []
        for i, col in enumerate(numerical_cols):
            fig, ax = plt.subplots(figsize=(8, 6))
            sns.boxplot(x=df[col], color=colors[i % len(colors)], ax=ax)
            title_str = f"Box Plot of {col}"
            ax.set_title(title_str)
            ax.set_xlabel(f"{col} Values")
            ax.grid(True, linestyle='--', alpha=0.6)
            filename = get_filename_from_title(title_str)
            figures.append((filename, fig))
        return figures
    except Exception as e:
        print(f"Error generating boxplot figures: {e}")
        return None

def get_scatterplot_figures(df, theme="Blue"):
    """Generates scatterplot figures for unique pairs of numerical columns.
    Returns a list of tuples: (filename, figure)."""
    try:
        themes = {
            "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"],
            "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"],
            "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"],
            "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"],
            "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"],
            "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],
            "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"]
        }
        if theme not in themes:
            st.write(f"Theme '{theme}' not found. Using Blue theme.")
            theme = "Blue"
        colors = themes[theme]
        numerical_cols = df.select_dtypes(include=['number']).columns
        figures = []
        for i, col1 in enumerate(numerical_cols):
            for j, col2 in enumerate(numerical_cols):
                if i < j:  # Plot only unique pairs
                    fig, ax = plt.subplots(figsize=(8, 6))
                    sns.scatterplot(x=df[col1], y=df[col2], color=colors[(i + j) % len(colors)], ax=ax)
                    title_str = f"Scatter Plot: {col1} vs {col2}"
                    ax.set_title(title_str)
                    ax.set_xlabel(f"{col1} Values")
                    ax.set_ylabel(f"{col2} Values")
                    ax.grid(True, linestyle='--', alpha=0.6)
                    filename = get_filename_from_title(title_str)
                    figures.append((filename, fig))
        return figures
    except Exception as e:
        print(f"Error generating scatterplot figures: {e}")
        return None

def get_lineplot_figures(df, theme="Blue"):
    """Generates lineplot figures for all numerical columns.
    Returns a list of tuples: (filename, figure)."""
    try:
        themes = {
            "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"],
            "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"],
            "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"],
            "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"],
            "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"],
            "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],
            "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"]
        }
        if theme not in themes:
            st.write(f"Theme '{theme}' not found. Using Blue theme.")
            theme = "Blue"
        colors = themes[theme]
        numerical_cols = df.select_dtypes(include=['number']).columns
        figures = []
        for i, col in enumerate(numerical_cols):
            fig, ax = plt.subplots(figsize=(8, 6))
            sns.lineplot(x=df.index, y=df[col], color=colors[i % len(colors)], ax=ax)
            title_str = f"Line Plot of {col}"
            ax.set_title(title_str)
            ax.set_xlabel("Index")
            ax.set_ylabel(f"{col} Values")
            ax.grid(True, linestyle='--', alpha=0.6)
            filename = get_filename_from_title(title_str)
            figures.append((filename, fig))
        return figures
    except Exception as e:
        print(f"Error generating lineplot figures: {e}")
        return None

def get_areaplot_figures(df, theme="Blue"):
    """Generates areaplot figures for all numerical columns.
    Returns a list of tuples: (filename, figure)."""
    try:
        themes = {
            "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"],
            "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"],
            "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"],
            "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"],
            "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"],
            "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],
            "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"]
        }
        if theme not in themes:
            st.write(f"Theme '{theme}' not found. Using Blue theme.")
            theme = "Blue"
        colors = themes[theme]
        numerical_cols = df.select_dtypes(include=['number']).columns
        figures = []
        for i, col in enumerate(numerical_cols):
            fig, ax = plt.subplots(figsize=(8, 6))
            ax.fill_between(df.index, df[col], color=colors[i % len(colors)], alpha=0.6)
            title_str = f"Area Plot of {col}"
            ax.set_title(title_str)
            ax.set_xlabel("Index")
            ax.set_ylabel(f"{col} Values")
            ax.grid(True, linestyle='--', alpha=0.6)
            filename = get_filename_from_title(title_str)
            figures.append((filename, fig))
        return figures
    except Exception as e:
        print(f"Error generating areaplot figures: {e}")
        return None

def get_violinplot_figures(df, theme="Blue"):
    """Generates violinplot figures for all numerical columns.
    Returns a list of tuples: (filename, figure)."""
    try:
        themes = {
            "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"],
            "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"],
            "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"],
            "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"],
            "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"],
            "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],
            "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"]
        }
        if theme not in themes:
            st.write(f"Theme '{theme}' not found. Using Blue theme.")
            theme = "Blue"
        colors = themes[theme]
        numerical_cols = df.select_dtypes(include=['number']).columns
        figures = []
        for i, col in enumerate(numerical_cols):
            fig, ax = plt.subplots(figsize=(8, 6))
            sns.violinplot(y=df[col], color=colors[i % len(colors)], ax=ax)
            title_str = f"Violin Plot of {col}"
            ax.set_title(title_str)
            ax.set_ylabel(f"{col} Values")
            ax.grid(True, linestyle='--', alpha=0.6)
            filename = get_filename_from_title(title_str)
            figures.append((filename, fig))
        return figures
    except Exception as e:
        print(f"Error generating violinplot figures: {e}")
        return None

def get_correlation_heatmap_figure(df, theme="Blue"):
    """Generates a correlation heatmap figure for numerical columns.
    Returns a list of tuples: (filename, figure)."""
    try:
        themes = {
            "Blue": ["#E0F2FE", "#B3E5FC", "#81D4FA", "#4FC3F7", "#29B6F6", "#03A9F4", "#039BE5", "#0288D1", "#01579B"],
            "Green": ["#E8F5E9", "#C8E6C9", "#A5D6A7", "#81C784", "#66BB6A", "#4CAF50", "#43A047", "#388E3C", "#2E7D32"],
            "Red": ["#FFE0E0", "#FFCDD2", "#EF9A9A", "#E57373", "#F44336", "#D32F2F", "#C62828", "#B71C1C", "#8B0000"],
            "Purple": ["#F3E5F5", "#E1BEE7", "#CE93D8", "#BA68C8", "#9C27B0", "#8E24AA", "#7B1FA2", "#6A1B9A", "#4A148C"],
            "Orange": ["#FFF3E0", "#FFE0B2", "#FFCC80", "#FFB366", "#FFA726", "#FF9800", "#FB8C00", "#F57C00", "#E65100"],
            "Gray": ["#F5F5F5", "#EEEEEE", "#E0E0E0", "#BDBDBD", "#9E9E9E", "#757575", "#616161", "#424242", "#212121"],
            "Pastel": ["#FFF9C4", "#FFECB3", "#FFE082", "#FFD54F", "#FFCA28", "#FFC107", "#FFB300", "#FFA000", "#FF6F00"]
        }

        if theme not in themes:
            st.write(f"Theme '{theme}' not found. Using Blue theme.")
            theme = "Blue"

        colors = themes[theme]

        numerical_cols = df.select_dtypes(include=['number']).columns
        corr_matrix = df[numerical_cols].corr()

        fig, ax = plt.subplots(figsize=(10, 8))
        cmap = sns.color_palette(colors, as_cmap=True)
        sns.heatmap(corr_matrix, annot=True, cmap=cmap, linewidths=.5, ax=ax)
        title_str = "Correlation Heatmap"
        ax.set_title(title_str)
        filename = get_filename_from_title(title_str)
        return [(filename, fig)]
    except Exception as e:
        print(f"Error generating correlation heatmap figure: {e}")
        return None

def get_cdf_figures(df, theme="Blue"):
    """Generates CDF figures for all numerical columns.
    Returns a list of tuples: (filename, figure)."""
    try:
        themes = {
            "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"],
            "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"],
            "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"],
            "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"],
            "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"],
            "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],
            "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"]
        }
        if theme not in themes:
            st.write(f"Theme '{theme}' not found. Using Blue theme.")
            theme = "Blue"
        colors = themes[theme]
        numerical_cols = df.select_dtypes(include=['number']).columns
        figures = []
        for i, col in enumerate(numerical_cols):
            fig, ax = plt.subplots(figsize=(8, 6))
            sorted_data = np.sort(df[col])
            yvals = np.arange(len(sorted_data)) / float(len(sorted_data) - 1)
            ax.plot(sorted_data, yvals, color=colors[i % len(colors)])
            title_str = f"CDF of {col}"
            ax.set_title(title_str)
            ax.set_xlabel(f"{col} Values")
            ax.set_ylabel("Cumulative Probability")
            ax.grid(True, linestyle='--', alpha=0.6)
            filename = get_filename_from_title(title_str)
            figures.append((filename, fig))
        return figures
    except Exception as e:
        print(f"Error generating CDF figures: {e}")
        return None

# ============================================= Plotting for STR Data ==========================================================================================================================================================================

def get_categorical_barplot_figures(df, theme="Blue"):
    """Generates barplot figures for all categorical columns.
    Returns a list of tuples: (filename, figure)."""
    try:
        themes = {
            "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"],
            "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"],
            "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"],
            "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"],
            "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"],
            "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],
            "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"]
        }
        if theme not in themes:
            st.write(f"Theme '{theme}' not found. Using Blue theme.")
            theme = "Blue"
        colors = themes[theme]
        categorical_cols = df.select_dtypes(include=['object']).columns
        figures = []
        for i, col in enumerate(categorical_cols):
            fig, ax = plt.subplots(figsize=(8, 6))
            sns.countplot(y=df[col], color=colors[i % len(colors)], ax=ax)
            title_str = f"Bar Plot of {col}"
            ax.set_title(title_str)
            ax.set_xlabel("Count")
            ax.set_ylabel(f"{col} Categories")
            ax.grid(axis='x', linestyle='--', alpha=0.6)
            filename = get_filename_from_title(title_str)
            figures.append((filename, fig))
        return figures
    except Exception as e:
        print(f"Error generating categorical barplot figures: {e}")
        return None

def get_piechart_figures(df, theme="Blue"):
    """Generates pie chart figures for all categorical columns.
    Returns a list of tuples: (filename, figure)."""
    try:
        themes = {
            "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"],
            "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"],
            "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"],
            "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"],
            "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"],
            "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],
            "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"]
        }
        if theme not in themes:
            st.write(f"Theme '{theme}' not found. Using Blue theme.")
            theme = "Blue"
        colors = themes[theme]
        categorical_cols = df.select_dtypes(include=['object']).columns
        figures = []
        for i, col in enumerate(categorical_cols):
            fig, ax = plt.subplots(figsize=(8, 6))
            counts = df[col].value_counts()
            ax.pie(counts, labels=counts.index, autopct='%1.1f%%', colors=colors)
            title_str = f"Pie Chart of {col}"
            ax.set_title(title_str)
            filename = get_filename_from_title(title_str)
            figures.append((filename, fig))
        return figures
    except Exception as e:
        print(f"Error generating piechart figures: {e}")
        return None

def get_stacked_barplot_figures(df, theme="Blue"):
    """Generates stacked bar plot figures for pairs of categorical columns.
    Returns a list of tuples: (filename, figure)."""
    try:
        themes = {
            "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"],
            "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"],
            "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"],
            "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"],
            "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"],
            "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],
            "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"]
        }
        if theme not in themes:
            st.write(f"Theme '{theme}' not found. Using Blue theme.")
            theme = "Blue"
        colors = themes[theme]
        categorical_cols = df.select_dtypes(include=['object']).columns
        figures = []
        if len(categorical_cols) < 2:
            st.write("At least two categorical columns are required for a stacked bar plot.")
            return figures
        for i, col1 in enumerate(categorical_cols):
            for j, col2 in enumerate(categorical_cols):
                if i < j:
                    fig, ax = plt.subplots(figsize=(8, 6))
                    ct = pd.crosstab(df[col1], df[col2])
                    ct.plot(kind='bar', stacked=True, color=colors, ax=ax)
                    title_str = f"Stacked Bar Plot: {col1} vs {col2}"
                    ax.set_title(title_str)
                    ax.set_xlabel(col1)
                    ax.set_ylabel("Count")
                    ax.tick_params(axis='x', rotation=45)
                    plt.tight_layout()
                    filename = get_filename_from_title(title_str)
                    figures.append((filename, fig))
        return figures
    except Exception as e:
        print(f"Error generating stacked barplot figures: {e}")
        return None

def get_grouped_barplot_figures(df, theme="Blue"):
    """Generates grouped bar plot figures for pairs of categorical columns.
    Returns a list of tuples: (filename, figure)."""
    try:
        themes = {
            "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"],
            "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"],
            "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"],
            "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"],
            "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"],
            "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],
            "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"]
        }
        if theme not in themes:
            st.write(f"Theme '{theme}' not found. Using Blue theme.")
            theme = "Blue"
        colors = themes[theme]
        categorical_cols = df.select_dtypes(include=['object']).columns
        figures = []
        if len(categorical_cols) < 2:
            st.write("At least two categorical columns are required for a grouped bar plot.")
            return figures
        for i, col1 in enumerate(categorical_cols):
            for j, col2 in enumerate(categorical_cols):
                if i < j:
                    fig, ax = plt.subplots(figsize=(8, 6))
                    ct = pd.crosstab(df[col1], df[col2])
                    ct.plot(kind='bar', color=colors, ax=ax)
                    title_str = f"Grouped Bar Plot: {col1} vs {col2}"
                    ax.set_title(title_str)
                    ax.set_xlabel(col1)
                    ax.set_ylabel("Count")
                    plt.xticks(rotation=45, ha='right')
                    plt.tight_layout()
                    filename = get_filename_from_title(title_str)
                    figures.append((filename, fig))
        return figures
    except Exception as e:
        print(f"Error generating grouped barplot figures: {e}")
        return None

def get_wordcloud_figures(df, theme="Blue"):
    """Generates word cloud figures for all categorical columns.
    Returns a list of tuples: (filename, figure)."""
    try:
        themes = {
            "Blue": "Blues",
            "Green": "Greens",
            "Red": "Reds",
            "Purple": "Purples",
            "Orange": "Oranges",
            "Gray": "Greys",
            "Pastel": "Pastel1"
        }
        if theme not in themes:
            st.write(f"Theme '{theme}' not found. Using Blue theme.")
            theme = "Blue"
        colormap = themes[theme]
        categorical_cols = df.select_dtypes(include=['object']).columns
        figures = []
        for i, col in enumerate(categorical_cols):
            fig, ax = plt.subplots(figsize=(8, 6))
            text = ' '.join(df[col].dropna().astype(str))
            if not text:
                st.write(f"Column '{col}' contains no valid text data.")
                continue
            wordcloud = WordCloud(width=800, height=400, background_color='white', colormap=colormap).generate(text)
            ax.imshow(wordcloud, interpolation='bilinear')
            ax.axis('off')
            title_str = f"Word Cloud of {col}"
            ax.set_title(title_str)
            filename = get_filename_from_title(title_str)
            figures.append((filename, fig))
        return figures
    except Exception as e:
        print(f"Error generating wordcloud figures: {e}")
        return None

def get_countplot_figures(df, theme="Blue"):
    """Generates count plot figures for all categorical columns.
    Returns a list of tuples: (filename, figure)."""
    try:
        themes = {
            "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"],
            "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"],
            "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"],
            "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"],
            "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"],
            "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],
            "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"]
        }
        if theme not in themes:
            st.write(f"Theme '{theme}' not found. Using Blue theme.")
            theme = "Blue"
        colors = themes[theme]
        categorical_cols = df.select_dtypes(include=['object']).columns
        figures = []
        for i, col in enumerate(categorical_cols):
            fig, ax = plt.subplots(figsize=(8, 6))
            sns.countplot(data=df, x=col, color=colors[i % len(colors)], ax=ax)
            title_str = f"Count Plot of {col}"
            ax.set_title(title_str)
            ax.set_xlabel(f"{col} Categories")
            ax.set_ylabel("Count")
            ax.tick_params(axis='x', rotation=45)
            plt.tight_layout()
            filename = get_filename_from_title(title_str)
            figures.append((filename, fig))
        return figures
    except Exception as e:
        print(f"Error generating countplot figures: {e}")
        return None

def get_treemap_figures(df, theme="Blue"):
    """Generates treemap figures for all categorical columns.
    Returns a list of tuples: (filename, figure)."""
    try:
        themes = {
            "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"],
            "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"],
            "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"],
            "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"],
            "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"],
            "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],
            "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"]
        }
        if theme not in themes:
            st.write(f"Theme '{theme}' not found. Using Blue theme.")
            theme = "Blue"
        colors = themes[theme]
        categorical_cols = df.select_dtypes(include=['object']).columns
        figures = []
        for i, col in enumerate(categorical_cols):
            fig, ax = plt.subplots(figsize=(10, 8))
            counts = df[col].value_counts()
            labels = counts.index
            sizes = counts.values
            total = sizes.sum()
            percentages = [f'{size / total * 100:.2f}%' for size in sizes]
            labels_with_percentage = [f'{label}\n({percentage})' for label, percentage in zip(labels, percentages)]
            squarify.plot(sizes=sizes, label=labels_with_percentage, color=colors, alpha=0.8, ax=ax)
            ax.axis('off')
            title_str = f"Treemap of {col}"
            ax.set_title(title_str)
            filename = get_filename_from_title(title_str)
            figures.append((filename, fig))
        return figures
    except Exception as e:
        print(f"Error generating treemap figures: {e}")
        return None

# ============================================= PLOTTING PLOT ONLY ===================================================================================================================================

def plot_histograms(df, theme="Blue"):

    themes = {
        "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"], # Medium-high saturation blues
        "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"], # Medium-high saturation greens
        "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"], # Medium-high saturation reds
        "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"], # Medium-high saturation purples
        "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"], # Medium-high saturation oranges
        "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],  # Medium-high saturation grays
        "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"] # Medium-high saturation pastels
    }

    if theme not in themes:
        print(f"Theme '{theme}' not found. Using Blue theme.")
        theme = "Blue"

    colors = themes[theme]

    numerical_cols = df.select_dtypes(include=['number']).columns

    for i, col in enumerate(numerical_cols):
        fig, ax = plt.subplots(figsize=(8, 6))  # Create fig and ax
        sns.histplot(df[col], color=colors[i % len(colors)], kde=True, ax=ax)  # Use ax
        ax.set_title(f"Distribution of {col}")
        ax.set_xlabel(f"{col} Values")
        ax.set_ylabel("Frequency")
        ax.grid(True, linestyle='--', alpha=0.6)
        st.pyplot(fig, clear_figure=True)  # Pass fig to st.pyplot

def plot_boxplots(df, theme="Blue"):

    themes = {
        "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"], # Medium-high saturation blues
        "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"], # Medium-high saturation greens
        "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"], # Medium-high saturation reds
        "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"], # Medium-high saturation purples
        "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"], # Medium-high saturation oranges
        "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],  # Medium-high saturation grays
        "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"] # Medium-high saturation pastels
    }

    if theme not in themes:
        print(f"Theme '{theme}' not found. Using Blue theme.")
        theme = "Blue"

    colors = themes[theme]

    numerical_cols = df.select_dtypes(include=['number']).columns

    for i, col in enumerate(numerical_cols):
        fig, ax = plt.subplots(figsize=(8, 6))  # Create fig and ax
        sns.boxplot(x=df[col], color=colors[i % len(colors)], ax=ax)  # Use ax
        ax.set_title(f"Box Plot of {col}")
        ax.set_xlabel(f"{col} Values")
        ax.grid(True, linestyle='--', alpha=0.6)
        st.pyplot(fig, clear_figure=True)  # Pass fig to st.pyplot

def plot_scatterplots(df, theme="Blue"):

    themes = {
        "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"], # Medium-high saturation blues
        "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"], # Medium-high saturation greens
        "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"], # Medium-high saturation reds
        "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"], # Medium-high saturation purples
        "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"], # Medium-high saturation oranges
        "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],  # Medium-high saturation grays
        "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"] # Medium-high saturation pastels
    }

    if theme not in themes:
        print(f"Theme '{theme}' not found. Using Blue theme.")
        theme = "Blue"

    colors = themes[theme]

    numerical_cols = df.select_dtypes(include=['number']).columns

    for i, col1 in enumerate(numerical_cols):
        for j, col2 in enumerate(numerical_cols):
            if i < j:  # Plot only unique pairs
                fig, ax = plt.subplots(figsize=(8, 6))  # Create fig and ax
                sns.scatterplot(x=df[col1], y=df[col2], color=colors[(i + j) % len(colors)], ax=ax)  # Use ax
                ax.set_title(f"Scatter Plot: {col1} vs {col2}")
                ax.set_xlabel(f"{col1} Values")
                ax.set_ylabel(f"{col2} Values")
                ax.grid(True, linestyle='--', alpha=0.6)
                st.pyplot(fig, clear_figure=True)  # Pass fig to st.pyplot

def plot_lineplots(df, theme="Blue"):

    themes = {
        "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"], # Medium-high saturation blues
        "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"], # Medium-high saturation greens
        "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"], # Medium-high saturation reds
        "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"], # Medium-high saturation purples
        "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"], # Medium-high saturation oranges
        "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],  # Medium-high saturation grays
        "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"] # Medium-high saturation pastels
    }

    if theme not in themes:
        print(f"Theme '{theme}' not found. Using Blue theme.")
        theme = "Blue"

    colors = themes[theme]

    numerical_cols = df.select_dtypes(include=['number']).columns

    for i, col in enumerate(numerical_cols):
        fig, ax = plt.subplots(figsize=(8, 6))  # Create fig and ax
        sns.lineplot(x=df.index, y=df[col], color=colors[i % len(colors)], ax=ax)  # Use ax
        ax.set_title(f"Line Plot of {col}")
        ax.set_xlabel("Index")
        ax.set_ylabel(f"{col} Values")
        ax.grid(True, linestyle='--', alpha=0.6)
        st.pyplot(fig, clear_figure=True)  # Pass fig to st.pyplot

def plot_areaplots(df, theme="Blue"):

    themes = {
        "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"], # Medium-high saturation blues
        "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"], # Medium-high saturation greens
        "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"], # Medium-high saturation reds
        "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"], # Medium-high saturation purples
        "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"], # Medium-high saturation oranges
        "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],  # Medium-high saturation grays
        "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"] # Medium-high saturation pastels
    }

    if theme not in themes:
        print(f"Theme '{theme}' not found. Using Blue theme.")
        theme = "Blue"

    colors = themes[theme]

    numerical_cols = df.select_dtypes(include=['number']).columns

    for i, col in enumerate(numerical_cols):
        fig, ax = plt.subplots(figsize=(8, 6))  # Create fig and ax
        ax.fill_between(df.index, df[col], color=colors[i % len(colors)], alpha=0.6)  # Use ax for plotting
        ax.set_title(f"Area Plot of {col}")
        ax.set_xlabel("Index")
        ax.set_ylabel(f"{col} Values")
        ax.grid(True, linestyle='--', alpha=0.6)
        st.pyplot(fig, clear_figure=True)  # Pass fig to st.pyplot

def plot_violinplots(df, theme="Blue"):

    themes = {
        "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"], # Medium-high saturation blues
        "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"], # Medium-high saturation greens
        "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"], # Medium-high saturation reds
        "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"], # Medium-high saturation purples
        "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"], # Medium-high saturation oranges
        "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],  # Medium-high saturation grays
        "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"] # Medium-high saturation pastels
    }

    if theme not in themes:
        print(f"Theme '{theme}' not found. Using Blue theme.")
        theme = "Blue"

    colors = themes[theme]

    numerical_cols = df.select_dtypes(include=['number']).columns

    for i, col in enumerate(numerical_cols):
        fig, ax = plt.subplots(figsize=(8, 6))  # Create fig and ax
        sns.violinplot(y=df[col], color=colors[i % len(colors)], ax=ax)  # Use ax for plotting
        ax.set_title(f"Violin Plot of {col}")
        ax.set_ylabel(f"{col} Values")
        ax.grid(True, linestyle='--', alpha=0.6)
        st.pyplot(fig, clear_figure=True)  # Pass fig to st.pyplot

def plot_correlation_heatmap(df, cmap_option="Blue"):
    """
    Plots a correlation heatmap for numerical columns in a Pandas DataFrame using color palettes derived from previous themes.

    Args:
        df (pd.DataFrame): The DataFrame to plot the correlation heatmap from.
        cmap_option (str, optional): Color palette option. Defaults to "Blue".
    """

    cmap_options = {
        "Blue": ["#E0F2FE", "#B3E5FC", "#81D4FA", "#4FC3F7", "#29B6F6", "#03A9F4", "#039BE5", "#0288D1", "#01579B"],
        "Green": ["#E8F5E9", "#C8E6C9", "#A5D6A7", "#81C784", "#66BB6A", "#4CAF50", "#43A047", "#388E3C", "#2E7D32"],
        "Red": ["#FFE0E0", "#FFCDD2", "#EF9A9A", "#E57373", "#F44336", "#D32F2F", "#C62828", "#B71C1C", "#8B0000"],
        "Purple": ["#F3E5F5", "#E1BEE7", "#CE93D8", "#BA68C8", "#9C27B0", "#8E24AA", "#7B1FA2", "#6A1B9A", "#4A148C"],
        "Orange": ["#FFF3E0", "#FFE0B2", "#FFCC80", "#FFB366", "#FFA726", "#FF9800", "#FB8C00", "#F57C00", "#E65100"],
        "Gray": ["#F5F5F5", "#EEEEEE", "#E0E0E0", "#BDBDBD", "#9E9E9E", "#757575", "#616161", "#424242", "#212121"],
        "Pastel": ["#FFF9C4", "#FFECB3", "#FFE082", "#FFD54F", "#FFCA28", "#FFC107", "#FFB300", "#FFA000", "#FF6F00"]
    }

    if cmap_option not in cmap_options:
        print(f"cmap_option '{cmap_option}' not found. Using Blue cmap.")
        cmap_option = "Blue"

    colors = cmap_options[cmap_option]

    numerical_cols = df.select_dtypes(include=['number']).columns
    corr_matrix = df[numerical_cols].corr()

    fig, ax = plt.subplots(figsize=(10, 8))  # Create fig and ax
    cmap = sns.color_palette(colors, as_cmap=True)
    sns.heatmap(corr_matrix, annot=True, cmap=cmap, linewidths=.5, ax=ax)  # Use ax
    ax.set_title("Correlation Heatmap")
    st.pyplot(fig, clear_figure=True)  # Pass fig to st.pyplot

def plot_cdf(df, theme="Blue"):

    themes = {
        "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"], # Medium-high saturation blues
        "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"], # Medium-high saturation greens
        "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"], # Medium-high saturation reds
        "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"], # Medium-high saturation purples
        "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"], # Medium-high saturation oranges
        "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],  # Medium-high saturation grays
        "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"] # Medium-high saturation pastels
    }

    if theme not in themes:
        print(f"Theme '{theme}' not found. Using Blue theme.")
        theme = "Blue"

    colors = themes[theme]

    numerical_cols = df.select_dtypes(include=['number']).columns

    for i, col in enumerate(numerical_cols):
        fig, ax = plt.subplots(figsize=(8, 6))  # Create fig, ax

        # Calculate CDF
        sorted_data = np.sort(df[col])
        yvals = np.arange(len(sorted_data)) / float(len(sorted_data) - 1)

        ax.plot(sorted_data, yvals, color=colors[i % len(colors)])  # Use ax for plotting
        ax.set_title(f"CDF of {col}")
        ax.set_xlabel(f"{col} Values")
        ax.set_ylabel("Cumulative Probability")
        ax.grid(True, linestyle='--', alpha=0.6)
        st.pyplot(fig, clear_figure=True)  # Pass the figure to st.pyplot

# Plotting for STR Data ==========================================================================================================================================================================

def plot_categorical_barplots(df, theme="Blue"):

    themes = {
        "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"], # Medium-high saturation blues
        "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"], # Medium-high saturation greens
        "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"], # Medium-high saturation reds
        "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"], # Medium-high saturation purples
        "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"], # Medium-high saturation oranges
        "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],  # Medium-high saturation grays
        "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"] # Medium-high saturation pastels
    }

    if theme not in themes:
        print(f"Theme '{theme}' not found. Using Blue theme.")
        theme = "Blue"

    colors = themes[theme]

    categorical_cols = df.select_dtypes(include=['object']).columns

    for i, col in enumerate(categorical_cols):
        fig, ax = plt.subplots(figsize=(8, 6))  # Create fig, ax
        sns.countplot(y=df[col], color=colors[i % len(colors)], ax=ax)  # Plot on ax
        ax.set_title(f"Bar Plot of {col}")
        ax.set_xlabel("Count")
        ax.set_ylabel(f"{col} Categories")
        ax.grid(axis='x', linestyle='--', alpha=0.6)
        st.pyplot(fig, clear_figure=True)  # Pass fig to st.pyplot

def plot_piecharts(df, theme="Blue"):
    """
    Plots pie charts for all categorical (string) columns in a Pandas DataFrame with light/medium high-saturated color themes.

    Args:
        df (pd.DataFrame): The DataFrame to plot pie charts from.
        theme (str, optional): The color theme for the plots. Defaults to "Blue".
    """

    themes = {
        "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"], # Medium-high saturation blues
        "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"], # Medium-high saturation greens
        "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"], # Medium-high saturation reds
        "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"], # Medium-high saturation purples
        "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"], # Medium-high saturation oranges
        "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],  # Medium-high saturation grays
        "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"] # Medium-high saturation pastels
    }

    if theme not in themes:
        print(f"Theme '{theme}' not found. Using Blue theme.")
        theme = "Blue"

    colors = themes[theme]

    categorical_cols = df.select_dtypes(include=['object']).columns

    for i, col in enumerate(categorical_cols):
        fig, ax = plt.subplots(figsize=(8, 6))  # Create fig, ax
        counts = df[col].value_counts()
        ax.pie(counts, labels=counts.index, autopct='%1.1f%%', colors=colors)  # Plot on ax
        ax.set_title(f"Pie Chart of {col}")
        st.pyplot(fig, clear_figure=True)  # Pass fig to st.pyplot

def plot_stacked_barplots(df, theme="Blue"):
    """
    Plots stacked bar plots for pairs of categorical (string) columns in a Pandas DataFrame with light/medium high-saturated color themes.

    Args:
        df (pd.DataFrame): The DataFrame to plot stacked bar plots from.
        theme (str, optional): The color theme for the plots. Defaults to "Blue".
    """

    themes = {
        "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"], # Medium-high saturation blues
        "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"], # Medium-high saturation greens
        "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"], # Medium-high saturation reds
        "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"], # Medium-high saturation purples
        "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"], # Medium-high saturation oranges
        "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],  # Medium-high saturation grays
        "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"] # Medium-high saturation pastels
    }

    if theme not in themes:
        print(f"Theme '{theme}' not found. Using Blue theme.")
        theme = "Blue"

    colors = themes[theme]

    categorical_cols = df.select_dtypes(include=['object']).columns

    if len(categorical_cols) < 2:
        print("At least two categorical columns are required for a stacked bar plot.")
        return

    for i, col1 in enumerate(categorical_cols):
        for j, col2 in enumerate(categorical_cols):
            if i < j:
                fig, ax = plt.subplots(figsize=(8, 6))  # Create fig, ax
                ct = pd.crosstab(df[col1], df[col2])
                ct.plot(kind='bar', stacked=True, color=colors, ax=ax)  # Plot on ax
                ax.set_title(f"Stacked Bar Plot: {col1} vs {col2}")
                ax.set_xlabel(col1)
                ax.set_ylabel("Count")
                ax.tick_params(axis='x', rotation=45)
                plt.tight_layout()
                st.pyplot(fig, clear_figure=True)  # Pass fig to st.pyplot

def plot_grouped_barplots(df, theme="Blue"):
    """
    Plots grouped bar plots (clustered bar plots) for pairs of categorical (string) columns in a Pandas DataFrame with light/medium high-saturated color themes.

    Args:
        df (pd.DataFrame): The DataFrame to plot grouped bar plots from.
        theme (str, optional): The color theme for the plots. Defaults to "Blue".
    """

    themes = {
        "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"], # Medium-high saturation blues
        "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"], # Medium-high saturation greens
        "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"], # Medium-high saturation reds
        "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"], # Medium-high saturation purples
        "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"], # Medium-high saturation oranges
        "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],  # Medium-high saturation grays
        "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"] # Medium-high saturation pastels
    }

    if theme not in themes:
        print(f"Theme '{theme}' not found. Using Blue theme.")
        theme = "Blue"

    colors = themes[theme]

    categorical_cols = df.select_dtypes(include=['object']).columns

    if len(categorical_cols) < 2:
        print("At least two categorical columns are required for a grouped bar plot.")
        return

    for i, col1 in enumerate(categorical_cols):
        for j, col2 in enumerate(categorical_cols):
            if i < j:
                fig, ax = plt.subplots(figsize=(8, 6))  # Create fig, ax
                ct = pd.crosstab(df[col1], df[col2])
                ct.plot(kind='bar', color=colors, ax=ax)  # Plot on the ax
                ax.set_title(f"Grouped Bar Plot: {col1} vs {col2}")
                ax.set_xlabel(col1)
                ax.set_ylabel("Count")
                plt.xticks(rotation=45, ha='right')  # Rotate x-axis labels
                plt.tight_layout()
                st.pyplot(fig, clear_figure=True)  # Pass the fig to st.pyplot

def plot_wordclouds(df, theme="Blue"):
    """
    Plots word clouds for all categorical (string) columns in a Pandas DataFrame with color themes.

    Args:
        df (pd.DataFrame): The DataFrame to plot word clouds from.
        theme (str, optional): The color theme for the word clouds. Defaults to "Blue".
    """

    themes = {
        "Blue": "Blues",
        "Green": "Greens",
        "Red": "Reds",
        "Purple": "Purples",
        "Orange": "Oranges",
        "Gray": "Greys",
        "Pastel": "Pastel1"
    }

    if theme not in themes:
        print(f"Theme '{theme}' not found. Using Blue theme.")
        theme = "Blue"

    colormap = themes[theme]

    categorical_cols = df.select_dtypes(include=['object']).columns

    for i, col in enumerate(categorical_cols):
        fig, ax = plt.subplots(figsize=(8, 6))  # Create fig, ax
        text = ' '.join(df[col].dropna().astype(str))  # Combine all strings into one text
        if not text:
            print(f"Column '{col}' contains no valid text data.")
            continue

        wordcloud = WordCloud(width=800, height=400, background_color='white', colormap=colormap).generate(text)
        ax.imshow(wordcloud, interpolation='bilinear')  # Plot on ax
        ax.axis('off')
        ax.set_title(f"Word Cloud of {col}")
        st.pyplot(fig, clear_figure=True)  # Pass fig to st.pyplot

def plot_countplots(df, theme="Blue"):
    """
    Plots count plots for all categorical (string) columns in a Pandas DataFrame with light/medium high-saturated color themes.

    Args:
        df (pd.DataFrame): The DataFrame to plot count plots from.
        theme (str, optional): The color theme for the plots. Defaults to "Blue".
    """

    themes = {
        "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"], # Medium-high saturation blues
        "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"], # Medium-high saturation greens
        "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"], # Medium-high saturation reds
        "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"], # Medium-high saturation purples
        "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"], # Medium-high saturation oranges
        "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],  # Medium-high saturation grays
        "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"] # Medium-high saturation pastels
    }

    if theme not in themes:
        print(f"Theme '{theme}' not found. Using Blue theme.")
        theme = "Blue"

    colors = themes[theme]

    categorical_cols = df.select_dtypes(include=['object']).columns

    for i, col in enumerate(categorical_cols):
        fig, ax = plt.subplots(figsize=(8, 6))  # Create fig, ax
        sns.countplot(data=df, x=col, color=colors[i % len(colors)], ax=ax)  # Use ax
        ax.set_title(f"Count Plot of {col}")
        ax.set_xlabel(f"{col} Categories")
        ax.set_ylabel("Count")
        ax.tick_params(axis='x', rotation=45)
        plt.tight_layout()
        st.pyplot(fig, clear_figure=True)  # Pass fig to st.pyplot

def plot_treemaps(df, theme="Blue"):
    """
    Plots treemaps for categorical (string) columns in a Pandas DataFrame with light/medium high-saturated color themes, including percentage values on each block.

    Args:
        df (pd.DataFrame): The DataFrame to plot treemaps from.
        theme (str, optional): The color theme for the plots. Defaults to "Blue".
    """

    themes = {
        "Blue": ["#6699CC", "#7A99CC", "#80B3FF", "#7AA8D1", "#667FCC", "#5673B3", "#4660A6", "#384D99", "#283980"], # Medium-high saturation blues
        "Green": ["#70D4A2", "#70C189", "#5CA66B", "#4A8C50", "#3D773C", "#315C33", "#2A4E2E", "#223E26", "#192E1B"], # Medium-high saturation greens
        "Red": ["#E87975", "#E86260", "#D9504E", "#C94140", "#B63234", "#A2262C", "#8E1B24", "#7A111C", "#600A14"], # Medium-high saturation reds
        "Purple": ["#AA94BF", "#957DBF", "#806BBF", "#6F57B3", "#604A9E", "#503C88", "#433077", "#3A256A", "#281854"], # Medium-high saturation purples
        "Orange": ["#FFA500", "#FFB347", "#FFC170", "#FFCC99", "#FFD9C2", "#F2BA9C", "#E6A582", "#D99067", "#CC7A4E"], # Medium-high saturation oranges
        "Gray": ["#A8A8A8", "#B3B3B3", "#BEBEBE", "#C8C8C8", "#D3D3D3", "#DADADA", "#E3E3E3", "#EDEDED", "#F5F5F5"],  # Medium-high saturation grays
        "Pastel": ["#F0E68C", "#E6E096", "#DCD9A0", "#D2D3AA", "#C9CDBC", "#BFC7D1", "#B0C1D9", "#A2BBDF", "#8FB5E4"] # Medium-high saturation pastels
    }

    if theme not in themes:
        print(f"Theme '{theme}' not found. Using Blue theme.")
        theme = "Blue"

    colors = themes[theme]

    categorical_cols = df.select_dtypes(include=['object']).columns

    for i, col in enumerate(categorical_cols):
        fig, ax = plt.subplots(figsize=(10, 8))  # Create fig, ax
        counts = df[col].value_counts()
        labels = counts.index
        sizes = counts.values

        # Calculate percentages
        total = sizes.sum()
        percentages = [f'{size / total * 100:.2f}%' for size in sizes]

        # Combine labels with percentages
        labels_with_percentage = [f'{label}\n({percentage})' for label, percentage in zip(labels, percentages)]

        squarify.plot(sizes=sizes, label=labels_with_percentage, color=colors, alpha=0.8, ax=ax)  # Use ax for plotting
        ax.axis('off')
        ax.set_title(f"Treemap of {col}")
        st.pyplot(fig, clear_figure=True)  # Pass fig to st.pyplot

# ============================================== DOC REPORTING ===========================================================================================================

def generate_doc_report_en(df):
    """
    Generates a comprehensive DOCX report from EDA results with overall implications.

    Args:
        df (pd.DataFrame): The input DataFrame.
        output_filename (str): The name of the output DOCX file.
    """

    document = Document()

    # Basic Statistics & Overall Implications (Part 1)
    document.add_heading("Comprehensive Exploratory Data Analysis Report", level=1)
    basic_stats = f"This report provides a detailed overview of the dataset. The dataset comprises {df.shape[0]} observations (rows) and {df.shape[1]} variables (columns). "
    basic_stats += f"Notably, no duplicate rows were found, indicating data uniqueness. However, a significant number of missing cells were identified, totaling {df.isnull().sum().sum()}. This highlights potential data completeness issues that may require further investigation."
    document.add_paragraph(basic_stats)

    implications = "Overall, the dataset presents a combination of strengths and challenges. The absence of duplicate rows suggests a well-curated dataset, but the presence of missing values necessitates careful handling during analysis. "
    if df.isnull().sum().sum() / (df.shape[0] * df.shape[1]) > 0.1:
        implications += "The high percentage of missing data (over 10%) could significantly impact the reliability of statistical analyses and model building. "
    else:
        implications += "While missing values are present, their volume is relatively manageable, and suitable imputation or deletion strategies can be employed. "

    document.add_paragraph(implications)

    # Variable Types & Implications
    numeric_cols = df.select_dtypes(include=np.number).columns.tolist()
    text_cols = df.select_dtypes(include='object').columns.tolist()
    document.add_heading("Variable Types", level=2)
    var_types = f"The dataset features {len(numeric_cols)} numeric variables, including '{', '.join(numeric_cols[:5])}', and {len(text_cols)} text-based variables: '{', '.join(text_cols)}'. "
    var_types += "No categorical variables were identified in this dataset."
    document.add_paragraph(var_types)

    type_implications = "The composition of numeric and text variables suggests a dataset suitable for quantitative and qualitative analyses. "
    if len(numeric_cols) > len(text_cols):
        type_implications += "The dominance of numeric variables indicates a dataset primarily designed for statistical modeling and quantitative analysis. "
    elif len(text_cols) > len(numeric_cols):
        type_implications += "The prevalence of text-based variables suggests a dataset rich in textual information, potentially suitable for natural language processing (NLP) tasks or qualitative content analysis. "

    document.add_paragraph(type_implications)

    # Highly Correlated Variables & Implications
    if numeric_cols:
        document.add_heading("Correlation Analysis", level=2)
        corr_matrix = df[numeric_cols].corr().abs()
        for col in numeric_cols:
            if col in corr_matrix.columns:
                corr_series = corr_matrix[col].sort_values(ascending=False)
                document.add_heading(f"Correlation Analysis for '{col}'", level=3)

                try:
                    # Top 5 Mostly Correlated
                    top_5_mostly = ""
                    num_mostly = min(5, len(corr_series) - 1)
                    if len(corr_series) > 1:
                        for i in range(1, num_mostly + 1):
                            top_5_mostly += f"'{corr_series.index.iloc[i]}' with {corr_series.iloc[i]:.4f} correlation, "
                        if top_5_mostly:
                            top_5_mostly = top_5_mostly[:-2]
                            document.add_paragraph(f"The top {num_mostly} mostly correlated variables are: {top_5_mostly}. High correlation indicates a strong relationship, suggesting that these variables move together. It might be useful to examine these pairs more closely.")

                    # Top 5 Least Correlated
                    top_5_least = ""
                    num_least = min(5, len(corr_series) - 1)
                    start_least = max(len(corr_series) - num_least, 1)
                    if len(corr_series) > 1:
                        for i in range(start_least, len(corr_series)):
                            top_5_least += f"'{corr_series.index.iloc[i]}' with {corr_series.iloc[i]:.4f} correlation, "
                        if top_5_least:
                            top_5_least = top_5_least[:-2]
                            document.add_paragraph(f"The top {num_least} least correlated variables are: {top_5_least}. Low correlation suggests that these variables are relatively independent. This can be important for building models where independence is assumed.")

                    correlation_insight = f"For '{col}', the high correlations suggest that these variables might be used interchangeably or that they are driven by a common underlying factor. Low correlations indicate variables that contribute unique information. "
                    document.add_paragraph(correlation_insight)
                except IndexError:
                    document.add_paragraph("I'm sorry, this type of Data cannot be correlated. Please use data that is able to be correlated.")
                except Exception as e:
                    document.add_paragraph(f"An error occurred during correlation analysis: {e}. Please use data that is able to be correlated.")

    else:
        document.add_paragraph("Correlation data is not available as there are no numerical columns.")
        document.add_paragraph("Without numerical data, correlation implications cannot be provided.")

    # Variables with Unique Values & Implications
    unique_counts = df.nunique()
    document.add_heading("Variables with Unique Values", level=2)
    document.add_paragraph("The dataset displays a wide range of uniqueness across its variables.")
    for col in df.columns:
        document.add_paragraph(f"'{col}' contains {unique_counts[col]} unique values.")
    document.add_paragraph("This variability in uniqueness can provide insights into the nature and distribution of the data. High unique value counts can indicate identifiers or detailed categorical variables, while low counts suggest broad categories or limited variability.")

    unique_implications = "The distribution of unique values influences the choice of analytical methods. Columns with very high unique values might need special treatment, especially if they are identifiers that don't contribute to statistical modeling. "
    if any(unique_counts / df.shape[0] > 0.8):
        unique_implications += "Columns with very high cardinality (unique values approaching the number of rows) might be considered as identifiers and excluded from certain analyses. "
    document.add_paragraph(unique_implications)

    # Uniform Distribution (Simplified) & Implications
    document.add_heading("Uniform Distribution (Simplified)", level=2)
    document.add_paragraph("A simplified check for uniform distribution was conducted.")
    uniform_cols = []
    for col in df.columns:
        if df[col].nunique() > 10:
            if (df[col].value_counts(normalize=True).std() < 0.05):
                uniform_cols.append(col)
    if uniform_cols:
        uniform_str = f"Variables such as '{', '.join(uniform_cols)}' might exhibit a uniform distribution. "
        uniform_str += "A uniform distribution suggests that all values are equally likely, which can be important for certain statistical tests or modeling assumptions. "
    else:
        uniform_str = "No variables were identified as potentially having a uniform distribution based on this simplified check. "
    uniform_str += "Variables with fewer than 10 unique values were excluded from this check."
    document.add_paragraph(uniform_str)

    uniform_implications = "The presence or absence of uniform distributions impacts the choice of statistical tests. Uniform distributions can be important for hypothesis testing and simulation studies. "
    if uniform_cols:
        uniform_implications += "The detected potential uniform distributions might simplify certain modeling or hypothesis testing procedures. "
    else:
        uniform_implications += "The absence of strong indications of uniform distributions suggests that data transformations or alternative tests might be necessary. "
    document.add_paragraph(uniform_implications)

    # Missing Values & Implications
    document.add_heading("Missing Values", level=2)
    document.add_paragraph("The dataset contains missing values across multiple variables.")
    missing_values = df.isnull().sum()
    for col in missing_values.index:
        document.add_paragraph(f"'{col}' has {missing_values[col]} missing values.")
    missing_str = f"The variable '{missing_values.index[missing_values.argmax()]}' has the highest number of missing values, with {missing_values.max()} missing values. Addressing these missing values is crucial for accurate analysis. High missing value counts can bias results and reduce the reliability of conclusions."
    document.add_paragraph(missing_str)

    missing_implications = "The handling of missing values is critical. High missing value counts can lead to biased or unreliable results. "
    if missing_values.max() / df.shape[0] > 0.3:
        missing_implications += "Columns with over 30% missing values might be considered for removal or require advanced imputation techniques. "
    else:
        missing_implications += "The missing values can be addressed using standard imputation or deletion methods. "
    document.add_paragraph(missing_implications)

    # Top 5 Mostly and Least Correlated Variables & Implications
    if numeric_cols:
        document.add_heading("Correlation Analysis", level=2)
        corr_matrix = df[numeric_cols].corr().abs()
        for col in numeric_cols:
            if col in corr_matrix.columns:
                corr_series = corr_matrix[col].sort_values(ascending=False)
                document.add_heading(f"Correlation Analysis for '{col}'", level=3)
                try:
                    top_5_mostly = ""
                    for i in range(1, 6):
                        top_5_mostly += f"'{corr_series.index[i]}' with {corr_series[i]:.4f} correlation, "
                    top_5_mostly = top_5_mostly[:-2]
                    document.add_paragraph(f"The top 5 mostly correlated variables are: {top_5_mostly}. High correlation indicates a strong relationship, suggesting that these variables move together. It might be useful to examine these pairs more closely.")

                    top_5_least = ""
                    for i in range(len(corr_series) - 5, len(corr_series)):
                        top_5_least += f"'{corr_series.index[i]}' with {corr_series[i]:.4f} correlation, "
                    top_5_least = top_5_least[:-2]
                    document.add_paragraph(f"The top 5 least correlated variables are: {top_5_least}. Low correlation suggests that these variables are relatively independent. This can be important for building models where independence is assumed.")

                    correlation_insight = f"For '{col}', the high correlations suggest that these variables might be used interchangeably or that they are driven by a common underlying factor. Low correlations indicate variables that contribute unique information. "
                    document.add_paragraph(correlation_insight)
                except IndexError:
                    document.add_paragraph("I'm sorry, this type of Data cannot be correlated. Please use data that is able to be correlated.")
                except Exception as e:
                    document.add_paragraph(f"An error occurred during correlation analysis: {e}. Please use data that is able to be correlated.")

    else:
        document.add_paragraph("Correlation data is not available as there are no numerical columns.")
        document.add_paragraph("Without numerical data, correlation implications cannot be provided.")

    # Variables Insight & Overall Implications (Part 2)
    document.add_heading("Variables Insight", level=2)
    for column_name in df.columns:
        col = df[column_name]
        document.add_paragraph(f"Analysis of Column: {column_name}")
        insight_paragraph = f"The column '{column_name}' has a data type of {col.dtype}, with {col.nunique()} unique values and {col.isnull().sum()} missing values. "

        if pd.api.types.is_numeric_dtype(col):
            mean = col.mean()
            std = col.std()
            min_val = col.min()
            q25 = col.quantile(0.25)
            median = col.median()
            q75 = col.quantile(0.75)
            max_val = col.max()
            skew = col.skew()
            kurt = col.kurt()
            zeros = (col == 0).sum()

            insight_paragraph += f"Its mean is {mean:.4f}, standard deviation is {std:.4f}, minimum value is {min_val:.4f}, 25th percentile is {q25:.4f}, median is {median:.4f}, 75th percentile is {q75:.4f}, maximum value is {max_val:.4f}, skewness is {skew:.4f}, kurtosis is {kurt:.4f}, and the number of zeros is {zeros}. "

            if std > 0:
                insight_paragraph += f"A standard deviation of {std:.4f} indicates the spread of the data around the mean. "
            if skew > 1 or skew < -1:
                insight_paragraph += f"A skewness of {skew:.4f} suggests that the data is highly skewed. "
            elif skew > 0.5 or skew < -0.5:
                insight_paragraph += f"A skewness of {skew:.4f} suggests moderate skewness. "
            if kurt > 3:
                insight_paragraph += f"A kurtosis of {kurt:.4f} indicates a leptokurtic distribution (heavy tails). "
            elif kurt < 3:
                insight_paragraph += f"A kurtosis of {kurt:.4f} indicates a platykurtic distribution (light tails). "

        elif pd.api.types.is_string_dtype(col) or pd.api.types.is_object_dtype(col):
            most_frequent = col.mode()[0]
            insight_paragraph += f"The most frequent value is '{most_frequent}', which appears {(col == most_frequent).sum()} times. "
            if col.nunique() / len(col) > 0.5:
                insight_paragraph += "This column has a high cardinality, meaning many unique values relative to the total number of entries. "
            if col.isnull().sum() / len(col) > 0.5:
                insight_paragraph += "This column has a high percentage of missing values. "

        document.add_paragraph(insight_paragraph)

    overall_variable_implications = "The individual variable insights provide a granular understanding of the data's characteristics. Skewness and kurtosis values highlight potential deviations from normal distributions, which can impact the choice of statistical tests. High cardinality in text variables might require feature engineering or dimensionality reduction. "
    document.add_paragraph(overall_variable_implications)

    return document 

def generate_doc_report_id(df, output_filename="eda_report_indonesian.docx"):
    """
    Menghasilkan laporan DOCX yang komprehensif dari hasil EDA dengan implikasi keseluruhan.

    Args:
        df (pd.DataFrame): DataFrame input.
        output_filename (str): Nama file DOCX output.
    """

    document = Document()

    # Statistik Dasar & Implikasi Keseluruhan (Bagian 1)
    document.add_heading("Laporan Analisis Data Eksplorasi Komprehensif", level=1)
    basic_stats = f"Laporan ini memberikan tinjauan rinci tentang kumpulan data. Kumpulan data terdiri dari {df.shape[0]} observasi (baris) dan {df.shape[1]} variabel (kolom). "
    basic_stats += f"Khususnya, tidak ditemukan baris duplikat, yang menunjukkan keunikan data. Namun, sejumlah besar sel yang hilang teridentifikasi, dengan total {df.isnull().sum().sum()}. Ini menyoroti potensi masalah kelengkapan data yang mungkin memerlukan penyelidikan lebih lanjut."
    document.add_paragraph(basic_stats)

    implications = "Secara keseluruhan, kumpulan data menyajikan kombinasi kekuatan dan tantangan. Tidak adanya baris duplikat menunjukkan kumpulan data yang terkurasi dengan baik, tetapi keberadaan nilai yang hilang memerlukan penanganan yang cermat selama analisis. "
    if df.isnull().sum().sum() / (df.shape[0] * df.shape[1]) > 0.1:
        implications += "Persentase data yang hilang yang tinggi (lebih dari 10%) dapat secara signifikan memengaruhi keandalan analisis statistik dan pembuatan model. "
    else:
        implications += "Meskipun nilai yang hilang ada, volumenya relatif dapat dikelola, dan strategi imputasi atau penghapusan yang sesuai dapat diterapkan. "

    document.add_paragraph(implications)

    # Jenis Variabel & Implikasi
    numeric_cols = df.select_dtypes(include=np.number).columns.tolist()
    text_cols = df.select_dtypes(include='object').columns.tolist()
    document.add_heading("Jenis Variabel", level=2)
    var_types = f"Kumpulan data menampilkan {len(numeric_cols)} variabel numerik, termasuk '{', '.join(numeric_cols[:5])}', dan {len(text_cols)} variabel berbasis teks: '{', '.join(text_cols)}'. "
    var_types += "Tidak ada variabel kategorikal yang teridentifikasi dalam kumpulan data ini."
    document.add_paragraph(var_types)

    type_implications = "Komposisi variabel numerik dan teks menunjukkan kumpulan data yang cocok untuk analisis kuantitatif dan kualitatif. "
    if len(numeric_cols) > len(text_cols):
        type_implications += "Dominasi variabel numerik menunjukkan kumpulan data yang terutama dirancang untuk pemodelan statistik dan analisis kuantitatif. "
    elif len(text_cols) > len(numeric_cols):
        type_implications += "Prevalensi variabel berbasis teks menunjukkan kumpulan data yang kaya akan informasi tekstual, yang berpotensi cocok untuk tugas pemrosesan bahasa alami (NLP) atau analisis konten kualitatif. "

    document.add_paragraph(type_implications)

    # Variabel dengan Korelasi Tinggi & Implikasi
    if numeric_cols:
        corr_matrix = df[numeric_cols].corr().abs()
        upper = corr_matrix.where(np.triu(np.ones(corr_matrix.shape), k=1).astype(bool))
        highly_correlated = [column for column in upper.columns if any(upper[column] > 0.8)]
        document.add_heading("Variabel dengan Korelasi Tinggi (x > 0.8)", level=2)
        try:
            if highly_correlated:
                high_corr = f"Beberapa variabel menunjukkan korelasi tinggi (di atas 0.8), yang menunjukkan hubungan yang kuat atau potensi redundansi. Ini termasuk '{', '.join(highly_correlated)}'. "
                high_corr += "Korelasi tinggi seperti itu mungkin memerlukan pemeriksaan lebih lanjut untuk memahami ketergantungan yang mendasarinya. Korelasi di atas 0.8 umumnya menunjukkan hubungan linear positif atau negatif yang kuat. Ini dapat menyiratkan bahwa satu variabel adalah proksi untuk variabel lain, atau bahwa keduanya dipengaruhi oleh faktor umum."
                document.add_paragraph(high_corr)

                corr_implications = "Keberadaan variabel dengan korelasi tinggi menunjukkan potensi masalah multikolinearitas, yang dapat memengaruhi stabilitas dan interpretasi model regresi. "
                corr_implications += "Mungkin perlu dilakukan pemilihan fitur atau teknik reduksi dimensi untuk mengatasi hal ini. "
                document.add_paragraph(corr_implications)

            else:
                document.add_paragraph("Tidak ada variabel dengan korelasi di atas 0.8 yang ditemukan. Ini menunjukkan bahwa variabel numerik dalam kumpulan data tidak menunjukkan hubungan linear yang kuat satu sama lain.")
                document.add_paragraph("Tidak adanya korelasi yang kuat menyederhanakan pembuatan model karena multikolinearitas bukan masalah utama.")
        except Exception as e:
            document.add_paragraph(f"Terjadi kesalahan selama analisis korelasi tinggi: {e}. Harap gunakan data yang dapat dikorelasikan.")

    else:
        document.add_paragraph("Tidak ada kolom numerik yang ditemukan, jadi analisis korelasi tidak dapat dilakukan.")
        document.add_paragraph("Tanpa kolom numerik, penilaian korelasi variabel tidak dapat dilakukan.")

    # Variabel dengan Nilai Unik & Implikasi
    unique_counts = df.nunique()
    document.add_heading("Variabel dengan Nilai Unik", level=2)
    document.add_paragraph("Kumpulan data menampilkan berbagai keunikan di seluruh variabelnya.")
    for col in df.columns:
        document.add_paragraph(f"'{col}' berisi {unique_counts[col]} nilai unik.")
    document.add_paragraph("Variabilitas keunikan ini dapat memberikan wawasan tentang sifat dan distribusi data. Jumlah nilai unik yang tinggi dapat menunjukkan pengidentifikasi atau variabel kategorikal terperinci, sementara jumlah yang rendah menunjukkan kategori luas atau variabilitas terbatas.")

    unique_implications = "Distribusi nilai unik memengaruhi pilihan metode analisis. Kolom dengan nilai unik yang sangat tinggi mungkin memerlukan perlakuan khusus, terutama jika itu adalah pengidentifikasi yang tidak berkontribusi pada pemodelan statistik. "
    if any(unique_counts / df.shape[0] > 0.8):
        unique_implications += "Kolom dengan kardinalitas yang sangat tinggi (nilai unik mendekati jumlah baris) mungkin dianggap sebagai pengidentifikasi dan dikecualikan dari analisis tertentu. "
    document.add_paragraph(unique_implications)

    # Distribusi Seragam (Sederhana) & Implikasi
    document.add_heading("Distribusi Seragam (Sederhana)", level=2)
    document.add_paragraph("Pemeriksaan sederhana untuk distribusi seragam dilakukan.")
    uniform_cols = []
    for col in df.columns:
        if df[col].nunique() > 10:
            if (df[col].value_counts(normalize=True).std() < 0.05):
                uniform_cols.append(col)
    if uniform_cols:
        uniform_str = f"Variabel seperti '{', '.join(uniform_cols)}' mungkin menunjukkan distribusi seragam. "
        uniform_str += "Distribusi seragam menunjukkan bahwa semua nilai sama-sama mungkin, yang penting untuk tes statistik atau asumsi pemodelan tertentu. "
    else:
        uniform_str = "Tidak ada variabel yang teridentifikasi berpotensi memiliki distribusi seragam berdasarkan pemeriksaan sederhana ini. "
    uniform_str += "Variabel dengan kurang dari 10 nilai unik dikecualikan dari pemeriksaan ini."
    document.add_paragraph(uniform_str)

    uniform_implications = "Keberadaan atau tidak adanya distribusi seragam memengaruhi pilihan tes statistik. Distribusi seragam dapat penting untuk pengujian hipotesis dan studi simulasi. "
    if uniform_cols:
        uniform_implications += "Potensi distribusi seragam yang terdeteksi mungkin menyederhanakan prosedur pemodelan atau pengujian hipotesis tertentu. "
    else:
        uniform_implications += "Tidak adanya indikasi kuat distribusi seragam menunjukkan bahwa transformasi data atau tes alternatif mungkin diperlukan. "
    document.add_paragraph(uniform_implications)

    # Nilai yang Hilang & Implikasi
    document.add_heading("Nilai yang Hilang", level=2)
    document.add_paragraph("Kumpulan data berisi nilai yang hilang di beberapa variabel.")
    missing_values = df.isnull().sum()
    for col in missing_values.index:
        document.add_paragraph(f"'{col}' memiliki {missing_values[col]} nilai yang hilang.")
    missing_str = f"Variabel '{missing_values.index[missing_values.argmax()]}' memiliki jumlah nilai yang hilang tertinggi, dengan {missing_values.max()} nilai yang hilang. Mengatasi nilai yang hilang ini sangat penting untuk analisis"
    document.add_paragraph(missing_str)

    missing_implications = "Penanganan nilai yang hilang sangat penting. Jumlah nilai yang hilang yang tinggi dapat menyebabkan hasil yang bias atau tidak dapat diandalkan. "
    if missing_values.max() / df.shape[0] > 0.3:
        missing_implications += "Kolom dengan lebih dari 30% nilai yang hilang mungkin dipertimbangkan untuk dihapus atau memerlukan teknik imputasi lanjutan. "
    else:
        missing_implications += "Nilai yang hilang dapat diatasi menggunakan metode imputasi atau penghapusan standar. "
    document.add_paragraph(missing_implications)

    # 5 Variabel dengan Korelasi Tertinggi dan Terendah & Implikasi
    if numeric_cols:
        document.add_heading("Analisis Korelasi", level=2)
        corr_matrix = df[numeric_cols].corr().abs()
        for col in numeric_cols:
            if col in corr_matrix.columns:
                corr_series = corr_matrix[col].sort_values(ascending=False)
                document.add_heading(f"Analisis Korelasi untuk '{col}'", level=3)
                try:
                    top_5_mostly = ""
                    for i in range(1, 6):
                        top_5_mostly += f"'{corr_series.index[i]}' dengan korelasi {corr_series[i]:.4f}, "
                    top_5_mostly = top_5_mostly[:-2]
                    document.add_paragraph(f"5 variabel dengan korelasi tertinggi adalah: {top_5_mostly}. Korelasi tinggi menunjukkan hubungan yang kuat, menunjukkan bahwa variabel-variabel ini bergerak bersamaan. Mungkin berguna untuk memeriksa pasangan-pasangan ini lebih dekat.")

                    top_5_least = ""
                    for i in range(len(corr_series) - 5, len(corr_series)):
                        top_5_least += f"'{corr_series.index[i]}' dengan korelasi {corr_series[i]:.4f}, "
                    top_5_least = top_5_least[:-2]
                    document.add_paragraph(f"5 variabel dengan korelasi terendah adalah: {top_5_least}. Korelasi rendah menunjukkan bahwa variabel-variabel ini relatif independen. Ini dapat penting untuk membangun model di mana independensi diasumsikan.")

                    correlation_insight = f"Untuk '{col}', korelasi tinggi menunjukkan bahwa variabel-variabel ini mungkin digunakan secara bergantian atau bahwa mereka didorong oleh faktor mendasar yang sama. Korelasi rendah menunjukkan variabel-variabel yang memberikan informasi unik. "
                    document.add_paragraph(correlation_insight)
                except IndexError:
                    document.add_paragraph("Mohon maaf, tipe Data ini tidak dapat dikorelasikan. Silakan gunakan data yang dapat dikorelasikan.")
                except Exception as e:
                    document.add_paragraph(f"Terjadi kesalahan selama analisis korelasi: {e}. Silakan gunakan data yang dapat dikorelasikan.")
    else:
        document.add_paragraph("Data korelasi tidak tersedia karena tidak ada kolom numerik.")
        document.add_paragraph("Tanpa data numerik, implikasi korelasi tidak dapat diberikan.")

    # Wawasan Variabel & Implikasi Keseluruhan (Bagian 2)
    document.add_heading("Wawasan Variabel", level=2)
    for column_name in df.columns:
        col = df[column_name]
        document.add_paragraph(f"Analisis Kolom: {column_name}")
        insight_paragraph = f"Kolom '{column_name}' memiliki tipe data {col.dtype}, dengan {col.nunique()} nilai unik dan {col.isnull().sum()} nilai yang hilang. "

        if pd.api.types.is_numeric_dtype(col):
            mean = col.mean()
            std = col.std()
            min_val = col.min()
            q25 = col.quantile(0.25)
            median = col.median()
            q75 = col.quantile(0.75)
            max_val = col.max()
            skew = col.skew()
            kurt = col.kurt()
            zeros = (col == 0).sum()

            insight_paragraph += f"Rata-ratanya adalah {mean:.4f}, standar deviasi adalah {std:.4f}, nilai minimum adalah {min_val:.4f}, persentil ke-25 adalah {q25:.4f}, median adalah {median:.4f}, persentil ke-75 adalah {q75:.4f}, nilai maksimum adalah {max_val:.4f}, skewness adalah {skew:.4f}, kurtosis adalah {kurt:.4f}, dan jumlah nol adalah {zeros}. "

            if std > 0:
                insight_paragraph += f"Standar deviasi sebesar {std:.4f} menunjukkan penyebaran data di sekitar rata-rata. "
            if skew > 1 or skew < -1:
                insight_paragraph += f"Skewness sebesar {skew:.4f} menunjukkan bahwa data sangat miring. "
            elif skew > 0.5 or skew < -0.5:
                insight_paragraph += f"Skewness sebesar {skew:.4f} menunjukkan kemiringan sedang. "
            if kurt > 3:
                insight_paragraph += f"Kurtosis sebesar {kurt:.4f} menunjukkan distribusi leptokurtik (ekor berat). "
            elif kurt < 3:
                insight_paragraph += f"Kurtosis sebesar {kurt:.4f} menunjukkan distribusi platikurtik (ekor ringan). "

        elif pd.api.types.is_string_dtype(col) or pd.api.types.is_object_dtype(col):
            most_frequent = col.mode()[0]
            insight_paragraph += f"Nilai paling sering adalah '{most_frequent}', yang muncul {(col == most_frequent).sum()} kali. "
            if col.nunique() / len(col) > 0.5:
                insight_paragraph += "Kolom ini memiliki kardinalitas tinggi, yang berarti banyak nilai unik relatif terhadap jumlah total entri. "
            if col.isnull().sum() / len(col) > 0.5:
                insight_paragraph += "Kolom ini memiliki persentase nilai yang hilang yang tinggi. "

        document.add_paragraph(insight_paragraph)

    overall_variable_implications = "Wawasan variabel individu memberikan pemahaman terperinci tentang karakteristik data. Nilai skewness dan kurtosis menyoroti potensi penyimpangan dari distribusi normal, yang dapat memengaruhi pilihan tes statistik. Kardinalitas tinggi dalam variabel teks mungkin memerlukan rekayasa fitur atau reduksi dimensi. "
    document.add_paragraph(overall_variable_implications)

    return document 

# ============================================== Download Functions =========================================================================================

def aggregate_and_download_plots(df, selected_plot_types, theme="Blue", cmap_option="Blue"):
    """Aggregates figures from the selected plot types and returns a zip buffer."""
    all_figures = []

    # Normalize plot type names for comparison
    selected_plot_types_lower = [plot_type.lower() for plot_type in selected_plot_types]

    if "histograms" in selected_plot_types_lower:
        figs = get_histogram_figures(df, theme)
        if figs:
            all_figures.extend(figs)
    if "boxplots" in selected_plot_types_lower:
        figs = get_boxplot_figures(df, theme)
        if figs:
            all_figures.extend(figs)
    if "scatterplots" in selected_plot_types_lower:
        figs = get_scatterplot_figures(df, theme)
        if figs:
            all_figures.extend(figs)
    if "lineplots" in selected_plot_types_lower:
        figs = get_lineplot_figures(df, theme)
        if figs:
            all_figures.extend(figs)
    if "areaplots" in selected_plot_types_lower:
        figs = get_areaplot_figures(df, theme)
        if figs:
            all_figures.extend(figs)
    if "violinplots" in selected_plot_types_lower:
        figs = get_violinplot_figures(df, theme)
        if figs:
            all_figures.extend(figs)
    if "correlation heatmap" in selected_plot_types_lower:
        figs = get_correlation_heatmap_figure(df, cmap_option)
        if figs:
            all_figures.extend(figs)
    if "cdf" in selected_plot_types_lower:
        figs = get_cdf_figures(df, theme)
        if figs:
            all_figures.extend(figs)
    if "categorical barplots" in selected_plot_types_lower:
        figs = get_categorical_barplot_figures(df, theme)
        if figs:
            all_figures.extend(figs)
    if "piecharts" in selected_plot_types_lower:
        figs = get_piechart_figures(df, theme)
        if figs:
            all_figures.extend(figs)
    if "stacked barplots" in selected_plot_types_lower:
        figs = get_stacked_barplot_figures(df, theme)
        if figs:
            all_figures.extend(figs)
    if "grouped barplots" in selected_plot_types_lower:
        figs = get_grouped_barplot_figures(df, theme)
        if figs:
            all_figures.extend(figs)
    if "wordclouds" in selected_plot_types_lower:
        figs = get_wordcloud_figures(df, theme)
        if figs:
            all_figures.extend(figs)
    if "countplots" in selected_plot_types_lower:
        figs = get_countplot_figures(df, theme)
        if figs:
            all_figures.extend(figs)
    if "treemaps" in selected_plot_types_lower:
        figs = get_treemap_figures(df, theme)
        if figs:
            all_figures.extend(figs)

    if not all_figures:
        st.write("No plot types selected or no figures were generated.")
        return None  # Return None if no figures

    # Create a zip file in memory
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
        for filename, fig in all_figures:
            buf = io.BytesIO()
            fig.savefig(buf, format="png", bbox_inches="tight")
            buf.seek(0)
            zip_file.writestr(filename, buf.getvalue())
            plt.close(fig)  # Close figure to avoid memory warning.

    zip_buffer.seek(0)
    return zip_buffer  # return the zip buffer.

# Plotting for Report Data ================================================================================================================================================================================

def eda_dataframe_to_docx_en(df):
    """
    Performs Exploratory Data Analysis (EDA) on a DataFrame and creates a docx document.
    """
    document = Document()

    document.add_heading("Exploratory Data Analysis", 0)

    document.add_heading("Basic Statistics", level=1)
    document.add_paragraph(f"Number of Observations (Rows): {df.shape[0]}")
    document.add_paragraph(f"Number of Variables (Columns): {df.shape[1]}")
    document.add_paragraph(f"Duplicate Rows: {df.duplicated().sum()}")
    document.add_paragraph(f"Missing Cells: {df.isnull().sum().sum()}")

    document.add_heading("Variable Types", level=1)
    numeric_cols = df.select_dtypes(include=np.number).columns.tolist()
    categorical_cols = df.select_dtypes(include='category').columns.tolist()
    text_cols = df.select_dtypes(include='object').columns.tolist()

    document.add_paragraph(f"Numeric: {len(numeric_cols)} ({', '.join(map(str, numeric_cols))})")
    document.add_paragraph(f"Categorical: {len(categorical_cols)} ({', '.join(map(str, categorical_cols))})")
    document.add_paragraph(f"Text (Object): {len(text_cols)} ({', '.join(map(str, text_cols))})")

    document.add_heading("Highly Correlated Variables (x > 0.8)", level=1)
    if numeric_cols:
        corr_matrix = df[numeric_cols].corr().abs()
        upper = corr_matrix.where(np.triu(np.ones(corr_matrix.shape), k=1).astype(bool))
        highly_correlated = [column for column in upper.columns if any(upper[column] > 0.8)]
        if highly_correlated:
            document.add_paragraph(f"{', '.join(map(str, highly_correlated))}")
        else:
            document.add_paragraph("No highly correlated variables found.")
    else:
        document.add_paragraph("No numeric columns found to calculate correlation.")

    document.add_heading("Variables with Unique Values", level=1)
    unique_counts = df.nunique()
    document.add_paragraph(str(unique_counts))

    document.add_heading("Variables with Uniform Distribution (Simplified)", level=1)
    for col in df.columns:
        if df[col].nunique() > 10:
            if (df[col].value_counts(normalize=True).std() < 0.05):
                document.add_paragraph(f"Column: {col} might have a uniform distribution")
        else:
            document.add_paragraph(f"Column: {col} has less than 10 unique values, cannot check for uniform distribution")

    document.add_heading("Missing Values per Variable", level=1)
    document.add_paragraph(str(df.isnull().sum()))

    document.add_heading("Top 5 Mostly and Least Correlated Variables", level=1)
    if len(numeric_cols) > 0:
        corr_matrix = df[numeric_cols].corr().abs()
        for col in numeric_cols:
            if col in corr_matrix.columns:
                corr_series = corr_matrix[col].sort_values(ascending=False)
                document.add_paragraph(f"\nCorrelation with {col}:")
                document.add_paragraph("Top 5 Mostly Correlated:")
                document.add_paragraph(str(corr_series.head(6).tail(5)))
                document.add_paragraph("Top 5 Least Correlated:")
                document.add_paragraph(str(corr_series.tail(5)))

    return document

def eda_dataframe_to_docx_id(df):
    """
    Melakukan Analisis Data Eksplorasi (EDA) pada DataFrame dan membuat dokumen docx dalam bahasa Indonesia.
    """
    document = Document()

    document.add_heading("Analisis Data Eksplorasi", 0)

    document.add_heading("Statistik Dasar", level=1)
    document.add_paragraph(f"Jumlah Observasi (Baris): {df.shape[0]}")
    document.add_paragraph(f"Jumlah Variabel (Kolom): {df.shape[1]}")
    document.add_paragraph(f"Baris Duplikat: {df.duplicated().sum()}")
    document.add_paragraph(f"Sel yang Hilang: {df.isnull().sum().sum()}")

    document.add_heading("Tipe Variabel", level=1)
    numeric_cols = df.select_dtypes(include=np.number).columns.tolist()
    categorical_cols = df.select_dtypes(include='category').columns.tolist()
    text_cols = df.select_dtypes(include='object').columns.tolist()

    document.add_paragraph(f"Numerik: {len(numeric_cols)} ({', '.join(map(str, numeric_cols))})")
    document.add_paragraph(f"Kategorikal: {len(categorical_cols)} ({', '.join(map(str, categorical_cols))})")
    document.add_paragraph(f"Teks (Objek): {len(text_cols)} ({', '.join(map(str, text_cols))})")

    document.add_heading("Variabel dengan Korelasi Tinggi (x > 0.8)", level=1)
    if numeric_cols:
        corr_matrix = df[numeric_cols].corr().abs()
        upper = corr_matrix.where(np.triu(np.ones(corr_matrix.shape), k=1).astype(bool))
        highly_correlated = [column for column in upper.columns if any(upper[column] > 0.8)]
        if highly_correlated:
            document.add_paragraph(f"{', '.join(map(str, highly_correlated))}")
        else:
            document.add_paragraph("Tidak ditemukan variabel dengan korelasi tinggi.")
    else:
        document.add_paragraph("Tidak ditemukan kolom numerik untuk menghitung korelasi.")

    document.add_heading("Variabel dengan Nilai Unik", level=1)
    unique_counts = df.nunique()
    document.add_paragraph(str(unique_counts))

    document.add_heading("Variabel dengan Distribusi Seragam (Sederhana)", level=1)
    for col in df.columns:
        if df[col].nunique() > 10:
            if (df[col].value_counts(normalize=True).std() < 0.05):
                document.add_paragraph(f"Kolom: {col} mungkin memiliki distribusi seragam")
        else:
            document.add_paragraph(f"Kolom: {col} memiliki kurang dari 10 nilai unik, tidak dapat diperiksa distribusi seragam")

    document.add_heading("Nilai yang Hilang per Variabel", level=1)
    document.add_paragraph(str(df.isnull().sum()))

    document.add_heading("5 Variabel dengan Korelasi Tertinggi dan Terendah", level=1)
    if len(numeric_cols) > 0:
        corr_matrix = df[numeric_cols].corr().abs()
        for col in numeric_cols:
            if col in corr_matrix.columns:
                corr_series = corr_matrix[col].sort_values(ascending=False)
                document.add_paragraph(f"\nKorelasi dengan {col}:")
                document.add_paragraph("5 Korelasi Tertinggi:")
                document.add_paragraph(str(corr_series.head(6).tail(5)))
                document.add_paragraph("5 Korelasi Terendah:")
                document.add_paragraph(str(corr_series.tail(5)))

    return document

def eda_all_columns_to_docx_en(df):
    """
    Performs Exploratory Data Analysis on all columns iteratively and creates a docx document.
    """
    document = Document()
    document.add_heading("Column-Wise Exploratory Data Analysis", 0)

    for column_name in df.columns:
        col = df[column_name]
        document.add_heading(f"Analysis of Column: {column_name}", level=1)
        document.add_paragraph(f"Data Type: {col.dtype}")
        document.add_paragraph(f"Number of Unique Values: {col.nunique()}")
        document.add_paragraph(f"Number of Missing Values: {col.isnull().sum()}")

        if pd.api.types.is_numeric_dtype(col):
            document.add_paragraph(f"Mean: {col.mean()}")
            document.add_paragraph(f"Standard Deviation: {col.std()}")
            document.add_paragraph(f"Minimum: {col.min()}")
            document.add_paragraph(f"25th Percentile: {col.quantile(0.25)}")
            document.add_paragraph(f"Median: {col.median()}")
            document.add_paragraph(f"75th Percentile: {col.quantile(0.75)}")
            document.add_paragraph(f"Maximum: {col.max()}")
            document.add_paragraph(f"Skewness: {col.skew()}")
            document.add_paragraph(f"Kurtosis: {col.kurt()}")
            document.add_paragraph(f"Number of Zeros: {(col == 0).sum()}")

        elif pd.api.types.is_string_dtype(col) or pd.api.types.is_object_dtype(col):
            most_frequent = col.mode()[0]
            document.add_paragraph(f"Most Frequent Value: {most_frequent}")
            document.add_paragraph(f"Frequency of Most Frequent Value: {(col == most_frequent).sum()}")

    return document

def eda_all_columns_to_docx_id(df):
    """
    Melakukan Analisis Data Eksplorasi pada semua kolom secara iteratif dan membuat dokumen docx.
    """
    document = Document()
    document.add_heading("Analisis Data Eksplorasi Per Kolom", 0)

    for column_name in df.columns:
        col = df[column_name]
        document.add_heading(f"Analisis Kolom: {column_name}", level=1)
        document.add_paragraph(f"Tipe Data: {col.dtype}")
        document.add_paragraph(f"Jumlah Nilai Unik: {col.nunique()}")
        document.add_paragraph(f"Jumlah Nilai Hilang: {col.isnull().sum()}")

        if pd.api.types.is_numeric_dtype(col):
            document.add_paragraph(f"Rata-rata: {col.mean()}")
            document.add_paragraph(f"Standar Deviasi: {col.std()}")
            document.add_paragraph(f"Minimum: {col.min()}")
            document.add_paragraph(f"Persentil ke-25: {col.quantile(0.25)}")
            document.add_paragraph(f"Median: {col.median()}")
            document.add_paragraph(f"Persentil ke-75: {col.quantile(0.75)}")
            document.add_paragraph(f"Maksimum: {col.max()}")
            document.add_paragraph(f"Kemiringan (Skewness): {col.skew()}")
            document.add_paragraph(f"Kurtosis: {col.kurt()}")
            document.add_paragraph(f"Jumlah Nol: {(col == 0).sum()}")

        elif pd.api.types.is_string_dtype(col) or pd.api.types.is_object_dtype(col):
            most_frequent = col.mode()[0]
            document.add_paragraph(f"Nilai Paling Sering Muncul: {most_frequent}")
            document.add_paragraph(f"Frekuensi Nilai Paling Sering Muncul: {(col == most_frequent).sum()}")

    return document

# ============================================ QUOTA HANDLING =============================================================================

def get_login_df():
    """Retrieves login data from the 'Accounts' sheet."""
    url = f'https://docs.google.com/spreadsheet/ccc?key={sheetskey}&output=csv'
    df_login = pd.read_csv(url)
    return df_login

def send_email(email, message):
    """Sends an email notification."""
    try:
        sender_email = "insightfoxa@gmail.com"  # Replace with your sender email
        sender_password = "txsr udsn fhpe pfns"  # Replace with your app password

        msg = MIMEText(message)
        msg['Subject'] = 'Password Reminder'
        msg['From'] = sender_email
        msg['To'] = email

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, email, msg.as_string())

        st.success("Reminder email sent. Please check your inbox (and spam folder)!")
    except Exception as e:
        st.error(f"Failed to send email: {e}")

def get_quota_df():
    """Retrieves quota data from the 'Quotas' sheet."""
    sheet = client.open_by_key(sheetskey).worksheet("Quotas")
    data = sheet.get_all_records()
    return pd.DataFrame(data)

def update_quota(email, column):
    """
    Decrements the quota in the given column for the specified email.
    If quota is >0, subtracts 1 and updates the Google Sheet.
    """
    sheet = client.open_by_key(sheetskey).worksheet("Quotas")
    df = get_quota_df()
    user_idx = df.index[df['email'] == email].tolist()
    if user_idx:
        row_index = user_idx[0] + 2  # Adjust for header row and 1-indexing in Sheets
        current_quota = int(df.loc[user_idx[0], column])
        if current_quota > 0:
            new_quota = current_quota - 1
            # Compute the column letter (assumes columns are in order)
            col_idx = df.columns.get_loc(column)  # 0-based index
            col_letter = chr(65 + col_idx)  # 0 -> A, 1 -> B, etc.
            sheet.update_acell(f"{col_letter}{row_index}", new_quota)
        else:
            st.error("No quota remaining for this action.")
    else:
        st.error("User not found in quota records.")

# ========================================== PAYMENT PROCESSING =====================================================================================

def colnum_to_letters(n):
    """Convert a 1-indexed column number to Excel-style column letters."""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result

def get_gemini_repsonse(input_prompt, image_parts):
    model = genai.GenerativeModel('gemini-1.5-flash')
    response = model.generate_content([input_prompt, image_parts[0]])
    return response.text

def input_image_setup(uploaded_file):
    # Read the uploaded file into bytes
    bytes_data = uploaded_file.read()
    image_parts = [
        {
            "mime_type": uploaded_file.type,
            "data": bytes_data
        }
    ]
    return image_parts

def extract_amount(text):
    """Extracts currency and amount from the Gemini response."""
    try:
        pattern = r"([A-Z]{3})\s*-\s*([\d.]+)"
        matches = re.findall(pattern, text)
        if matches:
            return matches
        else:
            return None
    except Exception as e:
        st.error(f"Error extracting amount: {e}")
        return None

def get_quota_df():
    """Retrieves quota data from the 'Quotas' sheet."""
    sheet = client.open_by_key(sheetskey).worksheet("Quotas")
    data = sheet.get_all_records()
    return pd.DataFrame(data)

def add_quota(email, column, amount_to_add=1):
    """
    Updates the quota in the given column for the specified email by adding the specified amount.
    """
    sheet = client.open_by_key(sheetskey).worksheet("Quotas")
    df = get_quota_df()
    user_idx = df.index[df['email'] == email].tolist()
    if user_idx:
        row_index = user_idx[0] + 2  # Adjust for header row and 1-indexing in Sheets
        current_quota = int(df.loc[user_idx[0], column])
        new_quota = current_quota + amount_to_add
        col_idx = df.columns.get_loc(column) + 1  # convert 0-indexed to 1-indexed
        col_letter = colnum_to_letters(col_idx)
        sheet.update_acell(f"{col_letter}{row_index}", new_quota)
    else:
        st.error("User not found in quota records.")

def generate_image_hash(time_str):
    """Generates a hash of the transaction time string."""
    return hashlib.md5(time_str.encode()).hexdigest()

def update_image_hashes(email, image_hash):
    """Updates the image hashes for the specified email in Google Sheets."""
    sheet = client.open_by_key(sheetskey).worksheet("Quotas")
    df = get_quota_df()
    user_idx = df.index[df['email'] == email].tolist()
    if user_idx:
        row_index = user_idx[0] + 2
        # Get the counter cell from the column named 'counter'
        if "counter" in df.columns:
            counter_col_idx = df.columns.get_loc("counter") + 1  # convert 0-indexed to 1-indexed
            counter_col_letter = colnum_to_letters(counter_col_idx)
            counter_cell = sheet.acell(f"{counter_col_letter}{row_index}").value
            counter = int(counter_cell) if counter_cell and counter_cell.isdigit() else 1
        else:
            # If no counter column exists, default to 1
            counter = 1

        # Determine the hash column using the counter.
        hash_column_index = 5 + counter  # 5 is the base index for the first hash column (1-indexed)
        col_letter = colnum_to_letters(hash_column_index)
        sheet.update_acell(f"{col_letter}{row_index}", image_hash)

        counter += 1
        if counter > 25:
            counter = 1
        if "counter" in df.columns:
            sheet.update_acell(f"{counter_col_letter}{row_index}", counter)
    else:
        st.error("User not found in quota records.")

def check_image_hash_exists(email, image_hash):
    """Checks if the image hash already exists for the specified email."""
    sheet = client.open_by_key(sheetskey).worksheet("Quotas")
    df = get_quota_df()
    user_idx = df.index[df['email'] == email].tolist()
    if user_idx:
        row_index = user_idx[0] + 2
        for i in range(1, 26):
            hash_column_index = 5 + i  # 5 is the base index for the first hash column
            col_letter = colnum_to_letters(hash_column_index)
            cell_value = sheet.acell(f"{col_letter}{row_index}").value
            if cell_value == image_hash:
                return True
        return False
    else:
        st.error("User not found in quota records.")
        return False
    
# ========================================== STREAMLIT APP LAYOUT =========================================================================

# ========================================== REGISTRATION PAGE ================================================================================

if st.session_state.page == 1:
    col1_1, col1_2, col1_3 = st.columns([1,3,1])
    with col1_2 : 
        # st header coloring
        st.markdown("<h1 style='text-align: center; color: white;'>InsightFox</h1>", unsafe_allow_html=True)
        st.markdown("<h6 style='text-align: center; color: white;'>Automated One-step Data Reports</h6>", unsafe_allow_html=True)
        with st.expander("", expanded=True):
            reg_email = st.text_input("Email", key="reg_email")
            reg_password = st.text_input("Password", type="password", key="reg_password")
            st.warning("Please use unique passwords and do not register the same email twice.")

            col_reg1, col_reg2 = st.columns([4, 2])
            with col_reg2:
                remind_me = st.checkbox("Remind me via email")
                if st.button("Register"):
                    st.warning("Please wait...")
                    sheet_accounts = client.open_by_key(sheetskey).worksheet("Accounts")
                    sheet_quota = client.open_by_key(sheetskey).worksheet("Quotas")
                    existing_emails = sheet_accounts.col_values(1)
                    if reg_email in existing_emails:
                        st.error("Email already registered.")
                    else:
                        # Append to Accounts sheet
                        sheet_accounts.append_row([reg_email, reg_password])
                        # Append to Quotas sheet with initial quotas (adjust initial values as needed)
                        sheet_quota.append_row([reg_email, 0, 0, 0, 0])
                        st.success("Registration successful!")
                        if remind_me:
                            send_email(reg_email, f"Your account password is: {reg_password}\nPlease keep this information secure.")
            with col_reg1:
                st.write("")
                st.write("")
                st.write("")
                st.write("Already have an account?")
                if st.button("Go to Login Page"):
                    st.session_state.page = 0
                    st.rerun()

# ============================================== LOGIN FUNCTION =================================================================================

elif st.session_state.page == 0 and st.session_state.login == 0:
    col2_1, col2_2, col2_3 = st.columns([1,3,1])
    with col2_2 : 
        # st header coloring
        st.markdown("<h1 style='text-align: center; color: white;'>InsightFox</h1>", unsafe_allow_html=True)
        st.markdown("<h6 style='text-align: center; color: white;'>Automated One-step Data Reports</h6>", unsafe_allow_html=True)
        with st.expander("", expanded=True):
            login_email = st.text_input("Email", key="login_email")
            login_password = st.text_input("Password", type="password", key="login_password")

            col_log1, col_log2 = st.columns([10, 2])
            with col_log2:
                if st.button("Login"):
                    df_login = get_login_df()
                    user_row = df_login[(df_login['email'] == login_email) & (df_login['password'] == login_password)]
                    if not user_row.empty:
                        st.success("Login successful!")
                        st.session_state.login = 1
                        st.session_state.email = login_email
                        st.session_state.page = 2
                        st.rerun()
                    else:
                        st.error("Invalid email or password.")
            with col_log1:
                st.write("")
                st.write("")
                st.write("")
                st.write("Don't have an account? Register")
                if st.button("Go to Register Page"):
                    st.session_state.page = 1
                    st.rerun()

# ================================================= CENTRALIZED NAVBAR =============================================================================

if st.session_state.login == 1:
    colcen1, colcen2 = st.columns([2,7])
    with colcen2 :
        # Centralized navigation widget
        if "selected_page" not in st.session_state:
            st.session_state.selected_page = "Check and Buy Quota"

        pages = ["Data Cleaning/Preprocessing", "Exploratory Data Analysis Report", "Check and Buy Quota"]
        selected = st.segmented_control(
            "",
            pages,
            selection_mode="single",
            key="selected_page"
        )
        # Update the page state based on the selection
        if selected == "Data Cleaning/Preprocessing":
            st.session_state.page = 2
        elif selected == "Exploratory Data Analysis Report":
            st.session_state.page = 3
        elif selected == "Check and Buy Quota":
            st.session_state.page = 4

# ================================================= DATA CLEANING ==========================================================================================

if st.session_state.login == 1 and st.session_state.page == 2:

    with st.expander("", expanded=True):
        st.subheader("InsightFox's Data Cleaner")
        st.write("Clean, Fill, and generally prepare your dataset for data processing and report creation.")

        uploaded_file = st.file_uploader("Upload your Excel or CSV File", type=["csv", "xlsx"])
        if uploaded_file is not None:
            try:
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file)
                else:
                    df = pd.read_excel(uploaded_file)
                    
                if df.shape[0] > 5000 or df.shape[1] > 50:
                    st.error("The maximum is 5000 rows and 50 columns. Uploaded file was not saved.")
                    if 'df' in st.session_state:
                        del st.session_state['df'] #delete previous df
                    df = None # prevent the rest of the code from running

                else:
                    st.success("File uploaded successfully!")

            except Exception as e:
                st.error(f"Error reading file: {e}")
                df = None

            if df is not None:
                st.write("Uploaded Dataframe:")
                st.dataframe(df.head())

                st.session_state['df'] = df

    if uploaded_file is not None:
        if 'df' in st.session_state:
            df = st.session_state['df'].copy()

            # Column Selection
            columns = df.columns.tolist()

            if 'selected_columns' not in st.session_state:
                st.session_state['selected_columns'] = []

            with st.expander(" ", expanded=True):
                st.subheader("Select Columns to Analyze")
                st.write("""Click "All" for All Columns, and "Custom" for Custom Columns""")
                col_selection = st.radio("Select Columns:", ["All Columns", "Custom Columns"], key="col_selection")

                if col_selection == "All Columns":
                    selected_columns = columns
                    st.session_state['selected_columns'] = columns
                elif col_selection == "Custom Columns":
                    selected_columns = st.multiselect("Select Columns to Analyze:", columns, default=st.session_state['selected_columns'])
                    st.session_state['selected_columns'] = selected_columns
                else:
                    selected_columns = st.session_state['selected_columns']

            col1, col2, col3 = st.columns([1, 1, 1])
            #col3, col4 = st.columns([2,2])

            with col1 : 
                # Fill NaN Values
                with st.expander("", expanded=True):
                    st.subheader("Handle NaN (empty) Values")
                    fill_nan = st.checkbox("Yes", key="fill_nan")
                    if fill_nan:
                        st.write("""Click "All" for All Columns, and "Custom" for Custom Columns""")
                        st.write("""Choose to remove NaN (empty values) or fill them with AI Algorithm""")
                        nan_selection = st.radio("Select Columns for NaN Handling:", ["All", "Custom"], key="nan_selection")
                        if nan_selection == "All":
                            nan_columns = selected_columns
                        elif nan_selection == "Custom":
                            nan_columns = st.multiselect("Multiselect Variables", selected_columns, key="nan_columns")
                        else:
                            nan_columns = []
                        fill_or_remove = st.radio("Fill or Remove NaN", ["Fill", "Remove"], key="fill_or_remove")
            with col2 : 
                # Remove Duplicates
                with st.expander("", expanded=True):
                    st.subheader("Remove Duplicate Values")
                    remove_duplicates = st.checkbox("Yes", key="remove_duplicates")
                    if remove_duplicates:
                        st.write("""Click "All" for All Columns, and "Custom" for Custom Columns""")
                        duplicate_selection = st.radio("Select Columns for Duplicate Handling:", ["All", "Custom"], key="duplicate_selection")
                        if duplicate_selection == "All":
                            duplicate_columns = selected_columns
                        elif duplicate_selection == "Custom":
                            duplicate_columns = st.multiselect("Multiselect Variables", selected_columns, key="duplicate_columns")
                        else:
                            duplicate_columns = []
            with col3 : 
            # Remove Outliers
                with st.expander("", expanded=True):
                    st.subheader("Remove Outlier Values")
                    remove_outliers = st.checkbox("Yes", key="remove_outliers")
                    if remove_outliers:
                        st.write("""Click "All" for All Columns, and "Custom" for Custom Columns""")
                        outlier_selection = st.radio("Select Columns for Outlier Handling:", ["All", "Custom"], key="outlier_selection")
                        if outlier_selection == "All":
                            outlier_columns = selected_columns
                        elif outlier_selection == "Custom":
                            outlier_columns = st.multiselect("Multiselect Variables", selected_columns, key="outlier_columns")
                        else:
                            outlier_columns = []
            #with col4 : 
                # Choose Certain Category
                #with st.expander("Choose Certain Category", expanded=True):
                    #choose_category = st.checkbox("Yes", key="choose_category")
                    #if choose_category:
                        #category_column = st.selectbox("Select Column", selected_columns, key="category_column")
                        #unique_values = df[category_column].unique().tolist()
                        #selected_values = st.multiselect("Multiselect Strings Inside The Columns", unique_values, key="selected_values")

            # Check for Duplicates, NaNs, Outliers
            with st.expander("", expanded=True):
                st.subheader("Check for Duplicates/NaNs/Outliers")
                st.write("Check for Duplicates/NaNs/Outliers in your Current Data to better understand the problems inside the data")
                col5, col6, col7 = st.columns(3)
                if col5.button("Check Duplicates"):
                    duplicates = df[df.duplicated(subset=selected_columns, keep=False)]
                    st.write("Dataframe of Duplicates:")
                    st.dataframe(duplicates)

                if col6.button("Check NaNs"):
                    nan_rows = df[df.isnull().any(axis=1)]
                    nan_cols = df.loc[:, df.isnull().any(axis=0)]
                    st.write("Rows with NaN values:")
                    st.dataframe(nan_rows)
                    st.write("Columns with NaN values:")
                    st.dataframe(nan_cols)

                if col7.button("Check Outliers"):
                    outlier_rows = pd.DataFrame()
                    for col in selected_columns:
                        if pd.api.types.is_numeric_dtype(df[col]):
                            Q1 = df[col].quantile(0.25)
                            Q3 = df[col].quantile(0.75)
                            IQR = Q3 - Q1
                            lower_bound = Q1 - 1.5 * IQR
                            upper_bound = Q3 + 1.5 * IQR
                            outliers = df[(df[col] < lower_bound) | (df[col] > upper_bound)]
                            outlier_rows = pd.concat([outlier_rows, outliers])
                    st.write("Rows with Outliers:")
                    st.dataframe(outlier_rows.drop_duplicates())

        def remove_outliers_iqr(df, column):
            """Removes outliers from a DataFrame column using the IQR method."""
            Q1 = df[column].quantile(0.25)
            Q3 = df[column].quantile(0.75)
            IQR = Q3 - Q1
            lower_bound = Q1 - 1.5 * IQR
            upper_bound = Q3 + 1.5 * IQR
            df_filtered = df[(df[column] >= lower_bound) & (df[column] <= upper_bound)]
            return df_filtered

        if 'df' in st.session_state:
            df = st.session_state['df'].copy()

            st.markdown('<h3 style="font-size: 19px; text-align: left; color: #FFFFFF;">Click this to clean the data!!!</h3>', unsafe_allow_html=True)
            if st.button("Clean"):
                status_container_plots = st.empty()
                status_container_plots.info("Cleaning...")
                if fill_nan and nan_columns:
                    for col in nan_columns:
                        if fill_or_remove == "Fill":
                            if pd.api.types.is_numeric_dtype(df[col]):
                                df[col] = df[col].fillna(df[col].mean())
                            else:
                                df[col] = df[col].ffill()
                        elif fill_or_remove == "Remove":
                            df.dropna(subset=[col], inplace=True)

                if remove_duplicates and duplicate_columns:
                    df.drop_duplicates(subset=duplicate_columns, inplace=True)

                if remove_outliers and outlier_columns:
                    for col in outlier_columns:
                        if pd.api.types.is_numeric_dtype(df[col]):
                            df = remove_outliers_iqr(df, col)

                #if choose_category and selected_values:
                    #df = df[df[category_column].isin(selected_values)]

                #st.session_state['cleaned_df'] = df
                df_quota = get_quota_df()
                user_quota = df_quota[df_quota['email'] == st.session_state.email]
                if not user_quota.empty:
                        # quota_clean button
                        quota_clean_value = int(user_quota.iloc[0]["quota_clean"])
                        if quota_clean_value > 0:
                            update_quota(st.session_state.email, "quota_clean")
                            st.success(f"Success !! Quota left : {quota_clean_value - 1}")
                            st.success("Please wait for a few seconds, do not press anything until done")
                            st.session_state['cleaned_df'] = df
                            time.sleep(6)
                            st.rerun()  # Refresh to show updated quota
                        else:
                            st.error("No quota left!")

        if 'cleaned_df' in st.session_state :
            cleaned_df = st.session_state['cleaned_df']

            def to_excel(df):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                processed_data = output.getvalue()
                return processed_data

            excel_file = to_excel(cleaned_df)
            st.download_button(
                label="Download Cleaned Data as Excel",
                data=excel_file,
                file_name="cleaned_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ============================================== EDA REPORTING ==========================================================================================

if st.session_state.login == 1 and st.session_state.page == 3:

    all_plot_options = ["Histograms", "Boxplots", "Lineplots", "Areaplots",
                         "Violinplots", "Correlation Heatmap", "CDF", "Categorical Barplots",
                         "Piecharts", "Wordclouds", "Countplots", "Treemaps"]

    with st.expander("", expanded=True):
        st.subheader("InsightFox's Exploratory Data Analysis Reports and Data Plotting")
        st.write("Generate Exploratory Data Analysis Reports and Plot All of the Data Analytics of your data")

        uploaded_file = st.file_uploader("Upload your Excel or CSV File", type=["csv", "xlsx"], accept_multiple_files=False)

        if uploaded_file is not None:
            try:
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file)
                else:
                    df = pd.read_excel(uploaded_file)

                if df.shape[0] > 5000 or df.shape[1] > 50:
                    st.error("The maximum is 5000 rows and 50 columns. Uploaded file was not saved.")
                    if 'df' in st.session_state:
                        del st.session_state['df'] #delete previous df
                    df = None  # Prevent further processing
                else:
                    st.success("File uploaded successfully!")

            except Exception as e:
                st.error(f"Error reading file: {e}")
                df = None

            if df is not None:
                st.write("Uploaded Dataframe:")
                st.dataframe(df.head())

                st.session_state['df'] = df

    if 'plot_fig' not in st.session_state:
        st.session_state['plot_fig'] = None

    if uploaded_file is not None:
        if 'df' in st.session_state:
            col1, col2, col3 = st.columns([2, 2, 2])

            with col1:
                # Plot Examples (Expander)
                with st.expander("", expanded=True):
                    st.subheader("Plot Examples")
                    st.write("Confused of what plots to generate? see these examples for yourselves!")

                    example_plot_choice = st.selectbox("Choose Plot", options=all_plot_options, index=0)

                    theme_choice = st.selectbox(
                        "Select Plot Theme",
                        ["Orange", "Red", "Green", "Purple", "Blue", "Gray", "Pastel"],
                        index=0
                    )

                    if example_plot_choice == "Histograms":
                        st.session_state['plot_fig'] = plot_histograms(df_sample_int, theme=theme_choice)

                    elif example_plot_choice == "Boxplots":
                        st.session_state['plot_fig'] = plot_boxplots(df_sample_int, theme=theme_choice)

                    elif example_plot_choice == "Lineplots":
                        st.session_state['plot_fig'] = plot_lineplots(df_sample_int, theme=theme_choice)

                    elif example_plot_choice == "Areaplots":
                        st.session_state['plot_fig'] = plot_areaplots(df_sample_int, theme=theme_choice)

                    elif example_plot_choice == "Violinplots":
                        st.session_state['plot_fig'] = plot_violinplots(df_sample_int, theme=theme_choice)

                    elif example_plot_choice == "Correlation Heatmap":
                        st.session_state['plot_fig'] = plot_correlation_heatmap(df_sample_int, cmap_option=theme_choice)

                    elif example_plot_choice == "CDF":
                        st.session_state['plot_fig'] = plot_cdf(df_sample_int, theme=theme_choice)

                    elif example_plot_choice == "Categorical Barplots":
                        st.session_state['plot_fig'] = plot_categorical_barplots(df_sample_str, theme=theme_choice)

                    elif example_plot_choice == "Piecharts":
                        st.session_state['plot_fig'] = plot_piecharts(df_sample_str, theme=theme_choice)

                    elif example_plot_choice == "Stacked Barplots":
                        st.session_state['plot_fig'] = plot_stacked_barplots(df_sample_str, theme=theme_choice)

                    elif example_plot_choice == "Grouped Barplots":
                        st.session_state['plot_fig'] = plot_grouped_barplots(df_sample_str, theme=theme_choice)

                    elif example_plot_choice == "Wordclouds":
                        st.session_state['plot_fig'] = plot_wordclouds(df_sample_str, theme=theme_choice)

                    elif example_plot_choice == "Countplots":
                        st.session_state['plot_fig'] = plot_countplots(df_sample_str, theme=theme_choice)

                    elif example_plot_choice == "Treemaps":
                        st.session_state['plot_fig'] = plot_treemaps(df_sample_str, theme=theme_choice)

                if st.session_state['plot_fig'] is not None:
                    st.pyplot(st.session_state['plot_fig'], clear_figure=True)

            with col2:
                with st.expander("", expanded=True):
                    st.subheader("Key Statistics Generator")
                    st.markdown("Generate the key statistics of your Excel or CSV File and creates a :orange[Ms. Word Report]")
                    st.warning("Please do not run two buttons at once, wait until one service is done before initiating another")
                    if st.button("Generate Key Statistics"):
                        status_container = st.empty()
                        status_container.info("Generating Key Statistics reports...")
                        st.error("Please do not press any other button while the report is being generated!!!")
                        df_quota = get_quota_df()
                        user_quota = df_quota[df_quota['email'] == st.session_state.email]
                        if not user_quota.empty:
                            quota_plot_value = int(user_quota.iloc[0]["quota_plot"])
                            if quota_plot_value > 0:
                                update_quota(st.session_state.email, "quota_plot")
                                st.success(f"Success !! Quota left : {quota_plot_value - 1}")
                                st.success("Please wait for a few seconds, we're initiating the download button...")

                                # Determine language based on session state (or other logic)
                                language = st.session_state.get('language', 'en')  # Default to English if not set

                                buffer_basic = BytesIO()
                                buffer_columnwise = BytesIO()

                                doc_basic_en = eda_dataframe_to_docx_en(st.session_state['df'])
                                doc_columnwise_en = eda_all_columns_to_docx_en(st.session_state['df'])
                                doc_basic_id = eda_dataframe_to_docx_id(st.session_state['df'])
                                doc_columnwise_id = eda_all_columns_to_docx_id(st.session_state['df'])

                                doc_basic_en.save(buffer_basic)
                                buffer_basic.seek(0)
                                doc_columnwise_en.save(buffer_columnwise)
                                buffer_columnwise.seek(0)

                                buffer_basic_id = BytesIO()
                                buffer_columnwise_id = BytesIO()

                                doc_basic_id.save(buffer_basic_id)
                                buffer_basic_id.seek(0)
                                doc_columnwise_id.save(buffer_columnwise_id)
                                buffer_columnwise_id.seek(0)

                                zip_buffer = BytesIO()
                                with zipfile.ZipFile(zip_buffer, 'w') as zipf:
                                    zipf.writestr("EN - Basic Key Statistics Report.docx", buffer_basic.getvalue())
                                    zipf.writestr("EN - Columnwise Key Statistics Report.docx", buffer_columnwise.getvalue())
                                    zipf.writestr("ID - Basic Key Statistics Report.docx", buffer_basic_id.getvalue())
                                    zipf.writestr("ID - Columnwise Key Statistics Report.docx", buffer_columnwise_id.getvalue())
                                zip_buffer.seek(0)

                                status_container.success("Key Statistics reports generated!")
                                st.session_state['eda_zip_buffer'] = zip_buffer
                                time.sleep(6)
                            else:
                                st.error("No quota left!")

                    if 'eda_zip_buffer' in st.session_state:
                        st.download_button(
                            label="Download Key Statistics Reports (ZIP)",
                            data=st.session_state['eda_zip_buffer'],
                            file_name="ID_EN Key Statistics Report.zip",
                            mime="application/zip"
                        )

                with st.expander("", expanded=True):
                    st.subheader("Exploratory Data Analysis")
                    st.markdown("Generate the key statistics and analyzes (paragraphed) of your Excel or CSV File and creates a :orange[Ms. Word Report]")
                    st.warning("Please do not run two buttons at once, wait until one service is done before initiating another")
                    if st.button("Generate Report"):
                        status_container = st.empty()
                        status_container.info("Generating Exploratory Data Analysis reports...")
                        st.error("Please do not press any other button while the report is being generated!!!")
                        df_quota = get_quota_df()
                        user_quota = df_quota[df_quota['email'] == st.session_state.email]
                        if not user_quota.empty:
                            quota_plot_value = int(user_quota.iloc[0]["quota_plot"])
                            if quota_plot_value > 0:
                                update_quota(st.session_state.email, "quota_plot")
                                st.success(f"Success !! Quota left : {quota_plot_value - 1}")
                                st.success("Please wait for a few seconds, we're initiating the download button...")
                                doc_en = generate_doc_report_en(st.session_state['df'])
                                buffer_en = BytesIO()
                                doc_en.save(buffer_en)
                                buffer_en.seek(0)

                                doc_id = generate_doc_report_id(st.session_state['df'])
                                buffer_id = BytesIO()
                                doc_id.save(buffer_id)
                                buffer_id.seek(0)

                                zip_doc = BytesIO()
                                with zipfile.ZipFile(zip_doc, 'w') as zipd:
                                    zipd.writestr("ENGLISH - Key Insights Analysis.docx", buffer_en.getvalue())
                                    zipd.writestr("INDONESIAN - Key Insights Analysis.docx", buffer_id.getvalue())
                                zip_doc.seek(0)

                                #status_container.success("Key Insight reports generated!")
                                st.session_state['eda_zip_doc'] = zip_doc
                                time.sleep(6)
                            else:
                                st.error("No quota left!")

                    if 'eda_zip_doc' in st.session_state:
                        st.download_button(
                            label="Download EDA Reports (ZIP)",
                            data=st.session_state['eda_zip_doc'],
                            file_name="ID_EN Exploratory Data Analysis Reports.zip",
                            mime="application/zip"
                        )

            with col3:
                with st.expander("", expanded=True):
                    st.subheader("Finished Plot Downloader")
                    st.markdown("Generate the desired plots (graphs, bars, etc) of your Excel or CSV File. Downloads a :orange[ZIP File containing all the plots (Max : 5 Types).]")
                    st.warning("Please do not run two buttons at once, wait until one service is done before initiating another")
                    selected_plot_types = st.multiselect(
                        "Choose plots to download",
                        options=[
                            "Histograms", "Boxplots", "Scatterplots", "Lineplots", "Areaplots",
                            "Violinplots", "Correlation Heatmap", "CDF", "Categorical Barplots",
                            "Piecharts", "Stacked Barplots", "Grouped Barplots", "Wordclouds",
                            "Countplots", "Treemaps"
                        ],
                        default=[
                            "Histograms", "Areaplots",
                            "Piecharts", "Wordclouds",
                            "Treemaps"
                        ],
                        max_selections=5
                    )
                    theme = st.selectbox("Select Theme", ["Orange", "Green", "Red", "Purple", "Blue", "Gray", "Pastel"], index=0)
                    if st.button("Generate Visualizations"):
                        status_container_plots = st.empty()
                        status_container_plots.info("Generating plots...")
                        st.error("Please do not press any other button while the plots are being generated!!!")
                        df_quota = get_quota_df()
                        user_quota = df_quota[df_quota['email'] == st.session_state.email]
                        if not user_quota.empty:
                            quota_plot_value = int(user_quota.iloc[0]["quota_plot"])
                            if quota_plot_value > 0:
                                update_quota(st.session_state.email, "quota_plot")
                                zip_buffer_plots = aggregate_and_download_plots(st.session_state['df'], selected_plot_types, theme)
                                #status_container_plots.success("Plots generated!")
                                st.session_state['plot_zip_buffer'] = zip_buffer_plots
                                st.success(f"Success !! Quota left : {quota_plot_value - 1}")
                                st.success("Please wait for a few seconds, we're initiating the download button...")
                                time.sleep(6)
                            else:
                                st.error("No quota left!")

                    if 'plot_zip_buffer' in st.session_state:
                        st.download_button(
                            label="Download Selected Visualizations (ZIP)",
                            data=st.session_state['plot_zip_buffer'],
                            file_name="Data Visualizations.zip",
                            mime="application/zip"
                        )

# ========================================== BUYING QUOTA ======================================================================================

if st.session_state.login == 1 and st.session_state.page == 4:
    
    # Load apikey dataframe if not already in session state
    if 'df_apikey' not in st.session_state:
        apisheetskey = "1sIEI-_9N96ndRJgWDyl0iL65bACeGQ74MncOV4HQCXY"
        url_apikey = f'https://docs.google.com/spreadsheet/ccc?key={apisheetskey}&output=csv'
        st.session_state.df_apikey = pd.read_csv(url_apikey)
    
    df_apikey = st.session_state.df_apikey
    platform = "Gemini"
    apikeyxloc = df_apikey['Platform'].str.contains(platform).idxmax()
    apikey = df_apikey.iloc[apikeyxloc, 2]
    
    # Configure genai only once
    if 'genai_configured' not in st.session_state:
        genai.configure(api_key=apikey)
        st.session_state.genai_configured = True
    
    # Load quota dataframe if not already in session state
    if 'df_quota' not in st.session_state:
        url = f'https://docs.google.com/spreadsheet/ccc?key={sheetskey}&output=xlsx'
        st.session_state.df_quota = pd.read_excel(url, sheet_name="Quotas")
    
    df_quota = st.session_state.df_quota
    user_row = df_quota[df_quota['email'] == st.session_state.email]
    if not user_row.empty:
        quota_clean = user_row['quota_clean'].values[0]
        quota_plot = user_row['quota_plot'].values[0]
        with st.expander("", expanded=True) : 
            st.markdown('<h3 style="font-size: 35px; text-align: center; color: #28282b;">MY QUOTA</h3>', unsafe_allow_html=True)
            st.subheader(f"Account In Use : {st.session_state.email}")
            st.success(f"Cleaning Quota: {quota_clean}")
            st.success(f"EDA Report, Key Stats Report, and Data Visualization Quota : {quota_plot}")
    else:
        st.warning("User not found in quota records.")

    st.subheader("")
    st.markdown('<h3 style="font-size: 35px; text-align: center; color: #FFFFFF;">BUY QUOTAS</h3>', unsafe_allow_html=True)

    colbuy1, colbuy2, colbuy3 = st.columns([1,1,1])
    with colbuy1: 
        with st.expander("", expanded=True) : 
            st.markdown('<h3 style="font-size: 35px; text-align: center; color: #28282b;">skibidi package</h3>', unsafe_allow_html=True)
            st.write("")
            st.markdown("This is cool ig... you can automate 2 reports for 2 files, or a library of an immense amount plots for :orange[2 files] with the maximum being :orange[50 columns] and :orange[5000 rows] per file.... yk 35k is kinda rlly cheap for all of this")
            st.markdown("This means that one report/A WHOLE LOT of plots is :orange[about IDR 15k]")
            st.write("")
            st.success("Cleaning Quota : 1")
            st.success("EDA Report, Key Stats Report, and Data Visualization Quota : 2")
            st.write("")
            st.markdown('<div style="background-color: #bd5c34; padding: 3px; border-radius: 15px; width: 80%; height: 50%; margin: auto; display: flex; justify-content: center; align-items: center;"><h1 style="color: #FFFFFF; font-size: 24px; margin-left: 22px;">35k</h1></div>', unsafe_allow_html=True)
            st.write("")
            st.write("")

    with colbuy2: 
        with st.expander("", expanded=True) : 
            st.markdown('<h3 style="font-size: 35px; text-align: center; color: #28282b;">Rizz Package</h3>', unsafe_allow_html=True)
            st.write("")
            st.markdown("It gets better?! now you have double the quota for :orange[less than the price of two...] you can automate the reports  or create a library of an immense amount plots for of :orange[4 files] containing a maximum of :orange[50 columns] and :orange[5000 rows] each....")
            st.markdown("This means that one report/A WHOLE LOT of plots is :orange[about IDR 15k]")
            st.write("")
            st.success("Cleaning Quota : 2")
            st.success("EDA Report, Key Stats Report, and Data Visualization Quota : 4")
            st.write("")
            st.markdown('<div style="background-color: #bd5c34; padding: 3px; border-radius: 15px; width: 80%; height: 50%; margin: auto; display: flex; justify-content: center; align-items: center;"><h1 style="color: #FFFFFF; font-size: 24px; margin-left: 22px;">60k</h1></div>', unsafe_allow_html=True)
            st.write("")
            st.write("")

    with colbuy3: 
        with st.expander("", expanded=True) : 
            st.markdown('<h3 style="font-size: 35px; text-align: center; color: #28282b;">SIGMA PACKAGE</h3>', unsafe_allow_html=True)
            st.write("")
            st.markdown("Now this is GAS!! :fire::fire::fire: now you have TRIPLE the quota for :orange[significantly less than the price of three...], automate the reports  or a create library of an immense amount plots for of :orange[6 files] with a max :orange[50 columns] and :orange[5000 rows] each....")
            st.markdown("This means that one report/A WHOLE LOT of plots is :orange[LESS THAN IDR 15k]")
            st.write("")
            st.success("Cleaning Quota : 3")
            st.success("EDA Report, Key Stats Report, and Data Visualization Quota : 6")
            st.write("")
            st.markdown('<div style="background-color: #bd5c34; padding: 3px; border-radius: 15px; width: 80%; height: 50%; margin: auto; display: flex; justify-content: center; align-items: center;"><h1 style="color: #FFFFFF; font-size: 24px; margin-left: 22px;">85k</h1></div>', unsafe_allow_html=True)
            st.write("")
            st.write("")  

    with st.expander("", expanded=True) : 
        st.header("Self-Checkout")
        st.write("Proceed Payment through QRIS provided and upload the transaction proof")
        st.warning("Make sure to only pay these EXACT amounts : IDR 35.000, IDR 60.000, IDR 85.000. The system will NOT detect other nominal payments.")
        st.write("")
        colqris1, colqris2, colqris3 = st.columns([1, 1, 1])
        with colqris2:
            st.image("QRIS_Keno.jpg", use_container_width=True)
        uploaded_file = st.file_uploader("Upload Transaction Image (JPEG or PNG)", type=["jpg", "jpeg", "png"])

        if uploaded_file is not None:
            image_parts = input_image_setup(uploaded_file)

            input_prompt = """
            Please extract the amount of money (along with its currency) that was transferred in the transaction along with the time of transaction.
            Just return me the currency, amount, and time. Follow the example below strictly. Use "." dot and not comma for the numbers.
            Remember, JUST INPUT LIKE THE EXAMPLE, NO NEED ANYTHING ELSE, NOT EVEN AN INTRO. 

            Example : 
            IDR - 45.000 : 19:21 (HH-MM)
            IDR - 35.000 : 21:07 (HH-MM)
            """

            if st.button("Verify Payment"):
                with st.spinner("Processing..."):
                    response_text = get_gemini_repsonse(input_prompt, image_parts)
                    try:
                        time_str = response_text.split(":")[-1].strip()  # Extract time part
                        transaction_time = time_str.split(" ")[0].strip()  # Remove "(HH-MM)"
                        image_hash = generate_image_hash(transaction_time)  # Hash the time string

                        if check_image_hash_exists(st.session_state.email, image_hash):
                            st.error("You have already paid with this transaction!")
                        else:
                            if "35" and "000" in response_text:
                                add_quota(st.session_state.email, "quota_clean", 1)
                                add_quota(st.session_state.email, "quota_plot", 2)
                                st.success("Payment Successful!!!")
                                st.warning("Please re-login (refresh the page) to refresh the quota")
                            elif "60" and "000" in response_text:
                                add_quota(st.session_state.email, "quota_clean", 2)
                                add_quota(st.session_state.email, "quota_plot", 4)
                                st.success("Payment Successful!!!")
                                st.warning("Please re-login (refresh the page) to refresh the quota")
                            elif "85" and "000" in response_text:
                                add_quota(st.session_state.email, "quota_clean", 3)
                                add_quota(st.session_state.email, "quota_plot", 6)
                                st.success("Payment Successful!!!")
                                st.warning("Please re-login (refresh the page) to refresh the quota")
                            else:
                                st.error("Self checkout failed.")
                                st.error("Please refer to contacting insightfoxa@gmail.com so we can manually add your quota, Include Transaction Proof and Account wish to be granted quota.")

                            update_image_hashes(st.session_state.email, image_hash)
                    except:
                        st.error("Self checkout failed.")
                        st.error("Please refer to contacting insightfoxa@gmail.com so we can manually add your quota, Include Transaction Proof and Account wish to be granted quota.")
