import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

# Set page config
st.set_page_config(page_title="Student Statistics", layout="wide")

# Title and intro
st.title("📊 Statistical Summary - Student Grades")
st.markdown("---")

# File path
file_path = 'الثانوية التأهيلية صلاح الدين الايوبي_الثالثة إعدادي مسار دولي.xlsx'

# Load data
@st.cache_data
def load_data():
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names
    
    # Filter out the first sheet if it's just a summary
    data_sheets = [s for s in sheet_names if s not in ['ExportMoGenNoteCcParMatie']]
    
    all_data = []
    for sheet in data_sheets:
        df = pd.read_excel(file_path, sheet_name=sheet, header=7)
        df['الفصل'] = sheet  # Add class name
        all_data.append(df)
    
    return pd.concat(all_data, ignore_index=True)

# Load the data
df = load_data()

# Convert grades from string (with commas) to float
subject_columns = [
    'اللغة العربية', 'اللغة الفرنسية', 'اللغة الإنجليزية',
    'الاجتماعيات', 'الرياضيات', 'علوم الحياة والأرض',
    'الفيزياء والكيمياء', 'التربية الإسلامية', 'التربية البدنية',
    'المعلوميات', 'المعدل'
]

for col in subject_columns:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '.'), errors='coerce')

# Sidebar for filtering
st.sidebar.header("🔍 Filter Options")
if 'الفصل' in df.columns:
    classes = ['All Classes'] + list(df['الفصل'].unique())
    selected_class = st.sidebar.selectbox("Select Class:", classes)
    if selected_class == 'All Classes':
        df_filtered = df.copy()
    else:
        df_filtered = df[df['الفصل'] == selected_class].copy()
else:
    df_filtered = df.copy()

# Remove rows with NaN in اسم التلميذ
df_filtered = df_filtered.dropna(subset=['اسم التلميذ'])

# Overall Statistics
st.header("📈 Overall Statistics")
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric("Total Students", len(df_filtered))

with col2:
    avg_grade = df_filtered['المعدل'].mean()
    st.metric("Average Grade", f"{avg_grade:.2f}")

with col3:
    max_grade = df_filtered['المعدل'].max()
    st.metric("Highest Grade", f"{max_grade:.2f}")

with col4:
    min_grade = df_filtered['المعدل'].min()
    st.metric("Lowest Grade", f"{min_grade:.2f}")

st.markdown("---")

# Grade Brackets Analysis
st.header("📊 Grade Brackets Analysis")

# Create grade brackets
def get_bracket(grade):
    if pd.isna(grade):
        return None
    elif grade < 10:
        return "0 - 9.99 (Below Average)"
    elif grade < 12:
        return "10 - 11.99 (Average)"
    else:
        return "12 - 20 (Good/Excellent)"

df_filtered['Bracket'] = df_filtered['المعدل'].apply(get_bracket)

# Calculate bracket statistics
bracket_stats = df_filtered.groupby('Bracket').agg({
    'المعدل': ['count', 'mean', 'min', 'max', 'std']
}).round(2)
bracket_stats.columns = ['Count', 'Mean', 'Min', 'Max', 'Std Dev']
bracket_stats = bracket_stats.reset_index()

# Display metrics for each bracket
col1, col2, col3 = st.columns(3)

below_avg = df_filtered[df_filtered['المعدل'] < 10]
average = df_filtered[(df_filtered['المعدل'] >= 10) & (df_filtered['المعدل'] < 12)]
good = df_filtered[df_filtered['المعدل'] >= 12]

with col1:
    st.markdown("### 🔴 Below Average (0 - 9.99)")
    st.metric("Students", len(below_avg))
    if len(below_avg) > 0:
        st.metric("Percentage", f"{len(below_avg)/len(df_filtered)*100:.1f}%")
        st.metric("Avg Grade", f"{below_avg['المعدل'].mean():.2f}")

with col2:
    st.markdown("### 🟡 Average (10 - 11.99)")
    st.metric("Students", len(average))
    if len(average) > 0:
        st.metric("Percentage", f"{len(average)/len(df_filtered)*100:.1f}%")
        st.metric("Avg Grade", f"{average['المعدل'].mean():.2f}")

with col3:
    st.markdown("### 🟢 Good/Excellent (12 - 20)")
    st.metric("Students", len(good))
    if len(good) > 0:
        st.metric("Percentage", f"{len(good)/len(df_filtered)*100:.1f}%")
        st.metric("Avg Grade", f"{good['المعدل'].mean():.2f}")

# Pie chart for bracket distribution
st.subheader("Grade Distribution by Bracket")
bracket_counts = df_filtered['Bracket'].value_counts().reset_index()
bracket_counts.columns = ['Bracket', 'Count']

col1, col2 = st.columns(2)

with col1:
    fig = px.pie(
        bracket_counts,
        values='Count',
        names='Bracket',
        color='Bracket',
        color_discrete_map={
            "0 - 9.99 (Below Average)": "#EF553B",
            "10 - 11.99 (Average)": "#FECB52",
            "12 - 20 (Good/Excellent)": "#00CC96"
        }
    )
    fig.update_traces(textposition='inside', textinfo='percent+value')
    fig.update_layout(height=400)
    st.plotly_chart(fig, use_container_width=True)

with col2:
    # Insights summary
    st.markdown("### 💡 Key Insights")
    total = len(df_filtered)
    
    # Success rate (>=10)
    success_rate = (len(average) + len(good)) / total * 100 if total > 0 else 0
    st.info(f"**Success Rate (≥10):** {success_rate:.1f}% of students passed")
    
    # Excellence rate (>=12)
    excellence_rate = len(good) / total * 100 if total > 0 else 0
    st.success(f"**Excellence Rate (≥12):** {excellence_rate:.1f}% of students achieved good/excellent grades")
    
    # At-risk students
    at_risk_rate = len(below_avg) / total * 100 if total > 0 else 0
    if at_risk_rate > 0:
        st.warning(f"**At-Risk Students (<10):** {at_risk_rate:.1f}% need additional support")
    
    # Performance summary
    if success_rate >= 80:
        st.markdown("✅ **Overall Performance:** Excellent - Most students are passing")
    elif success_rate >= 60:
        st.markdown("⚠️ **Overall Performance:** Good - Majority passing but room for improvement")
    else:
        st.markdown("🚨 **Overall Performance:** Needs Attention - Many students struggling")

# Students list by bracket
st.subheader("📋 Students by Bracket")
bracket_tab1, bracket_tab2, bracket_tab3 = st.tabs(["🔴 Below Average", "🟡 Average", "🟢 Good/Excellent"])

with bracket_tab1:
    if len(below_avg) > 0:
        st.dataframe(below_avg[['اسم التلميذ', 'الفصل', 'المعدل']].sort_values('المعدل', ascending=False), use_container_width=True)
    else:
        st.success("No students in this bracket!")

with bracket_tab2:
    if len(average) > 0:
        st.dataframe(average[['اسم التلميذ', 'الفصل', 'المعدل']].sort_values('المعدل', ascending=False), use_container_width=True)
    else:
        st.info("No students in this bracket")

with bracket_tab3:
    if len(good) > 0:
        st.dataframe(good[['اسم التلميذ', 'الفصل', 'المعدل']].sort_values('المعدل', ascending=False), use_container_width=True)
    else:
        st.info("No students in this bracket")

st.markdown("---")

# Detailed Statistics by Subject
st.header("📚 Statistics by Subject")

# Calculate statistics for each subject
stats_data = []
for col in subject_columns:
    if col in df_filtered.columns:
        valid_data = df_filtered[col].dropna()
        if len(valid_data) > 0:
            stats_data.append({
                'المادة': col,
                'المتوسط': valid_data.mean(),
                'الأعلى': valid_data.max(),
                'الأقل': valid_data.min(),
                'الانحراف المعياري': valid_data.std(),
                'عدد الطلاب': len(valid_data)
            })

stats_df = pd.DataFrame(stats_data)

# Display table
st.dataframe(
    stats_df.style.format({
        'المتوسط': '{:.2f}',
        'الأعلى': '{:.2f}',
        'الأقل': '{:.2f}',
        'الانحراف المعياري': '{:.2f}'
    }),
    use_container_width=True
)

st.markdown("---")

# Visualizations
st.header("📊 Visualizations")

col1, col2 = st.columns(2)

# Average grades by subject
with col1:
    st.subheader("Average Grades by Subject")
    fig = px.bar(
        stats_df.sort_values('المتوسط', ascending=True),
        x='المتوسط',
        y='المادة',
        orientation='h',
        color='المتوسط',
        color_continuous_scale='Viridis'
    )
    fig.update_layout(height=400)
    st.plotly_chart(fig, use_container_width=True)

# Grade distribution
with col2:
    st.subheader("Average Grade Distribution")
    fig = px.histogram(
        df_filtered,
        x='المعدل',
        nbins=20,
        color_discrete_sequence=['#636EFA']
    )
    fig.add_vline(df_filtered['المعدل'].mean(), line_dash="dash", line_color="red", 
                   annotation_text=f"Mean: {df_filtered['المعدل'].mean():.2f}")
    fig.update_layout(height=400)
    st.plotly_chart(fig, use_container_width=True)

st.markdown("---")

# Student Rankings
st.header("🏆 Top 10 Students by Average Grade")
top_students = df_filtered[['اسم التلميذ', 'المعدل']].dropna().nlargest(10, 'المعدل')
st.dataframe(top_students.reset_index(drop=True), use_container_width=True)

st.markdown("---")

# Performance by Subject - Box Plot
st.header("📊 Grade Distribution by Subject")
subject_data = []
for col in subject_columns:
    if col in df_filtered.columns:
        valid_data = df_filtered[col].dropna()
        for grade in valid_data:
            subject_data.append({'المادة': col, 'التقدير': grade})

if subject_data:
    subject_box_df = pd.DataFrame(subject_data)
    fig = px.box(subject_box_df, x='المادة', y='التقدير', color='المادة')
    fig.update_layout(height=500, showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

st.markdown("---")

# Raw Data Table
st.header("📋 Full Student Data")
st.dataframe(df_filtered[['ر.ت', 'رقم التلميذ', 'اسم التلميذ'] + subject_columns], 
             use_container_width=True, height=400)

# Download option
st.markdown("---")
csv = df_filtered.to_csv(index=False)
st.download_button(
    label="📥 Download Data as CSV",
    data=csv,
    file_name=f"student_data_statistics.csv",
    mime="text/csv"
)
