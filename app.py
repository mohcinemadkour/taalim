import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import io
import tempfile
import os

# Set page config
st.set_page_config(page_title="إحصائيات التلاميذ", layout="wide")

# File uploader in sidebar
st.sidebar.header("📁 تحميل الملف")
uploaded_file = st.sidebar.file_uploader(
    "اختر ملف Excel",
    type=['xlsx', 'xls'],
    help="قم بتحميل ملف Excel يحتوي على بيانات التلاميذ"
)

if uploaded_file is None:
    st.title("📊 إحصائيات نتائج التلاميذ")
    st.markdown("---")
    st.info("👈 الرجاء تحميل ملف Excel من القائمة الجانبية للبدء")
    st.markdown("""
    ### 📋 تعليمات الاستخدام:
    1. اضغط على **Browse files** في القائمة الجانبية
    2. اختر ملف Excel يحتوي على بيانات التلاميذ
    3. انتظر حتى يتم تحميل البيانات
    4. استعرض الإحصائيات والرسوم البيانية
    """)
    st.stop()

# Extract title from filename
app_title = Path(uploaded_file.name).stem.replace('_', ' - ')

# Title and intro
st.title(f"📊 {app_title}")
st.markdown("---")

# Load data
@st.cache_data
def load_data(file_content, file_name):
    xls = pd.ExcelFile(io.BytesIO(file_content))
    sheet_names = xls.sheet_names
    
    # Filter out the first sheet if it's just a summary
    data_sheets = [s for s in sheet_names if s not in ['ExportMoGenNoteCcParMatie']]
    
    all_data = []
    for sheet in data_sheets:
        df = pd.read_excel(io.BytesIO(file_content), sheet_name=sheet, header=7)
        df['الفصل'] = sheet  # Add class name
        all_data.append(df)
    
    return pd.concat(all_data, ignore_index=True)

# Load the data
file_content = uploaded_file.read()
df = load_data(file_content, uploaded_file.name)

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
st.sidebar.markdown("---")
st.sidebar.header("🔍 خيارات التصفية")
if 'الفصل' in df.columns:
    classes = ['جميع الفصول'] + list(df['الفصل'].unique())
    selected_class = st.sidebar.selectbox("اختر الفصل:", classes)
    if selected_class == 'جميع الفصول':
        df_filtered = df.copy()
    else:
        df_filtered = df[df['الفصل'] == selected_class].copy()
else:
    df_filtered = df.copy()

# Remove rows with NaN in اسم التلميذ
df_filtered = df_filtered.dropna(subset=['اسم التلميذ'])

# Overall Statistics
st.header("📈 الإحصائيات العامة")
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric("عدد التلاميذ", len(df_filtered))

with col2:
    avg_grade = df_filtered['المعدل'].mean()
    st.metric("المعدل العام", f"{avg_grade:.2f}")

with col3:
    max_grade = df_filtered['المعدل'].max()
    st.metric("أعلى معدل", f"{max_grade:.2f}")

with col4:
    min_grade = df_filtered['المعدل'].min()
    st.metric("أدنى معدل", f"{min_grade:.2f}")

st.markdown("---")

# Grade Brackets Analysis
st.header("📊 تحليل شرائح المعدلات")

# Create grade brackets
def get_bracket(grade):
    if pd.isna(grade):
        return None
    elif grade < 10:
        return "0 - 9.99 (دون المعدل)"
    elif grade < 12:
        return "10 - 11.99 (متوسط)"
    else:
        return "12 - 20 (جيد/ممتاز)"

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
    st.markdown("### 🔴 دون المعدل (0 - 9.99)")
    st.metric("عدد التلاميذ", len(below_avg))
    if len(below_avg) > 0:
        st.metric("النسبة المئوية", f"{len(below_avg)/len(df_filtered)*100:.1f}%")
        st.metric("متوسط المعدل", f"{below_avg['المعدل'].mean():.2f}")

with col2:
    st.markdown("### 🟡 متوسط (10 - 11.99)")
    st.metric("عدد التلاميذ", len(average))
    if len(average) > 0:
        st.metric("النسبة المئوية", f"{len(average)/len(df_filtered)*100:.1f}%")
        st.metric("متوسط المعدل", f"{average['المعدل'].mean():.2f}")

with col3:
    st.markdown("### 🟢 جيد/ممتاز (12 - 20)")
    st.metric("عدد التلاميذ", len(good))
    if len(good) > 0:
        st.metric("النسبة المئوية", f"{len(good)/len(df_filtered)*100:.1f}%")
        st.metric("متوسط المعدل", f"{good['المعدل'].mean():.2f}")

# Pie chart for bracket distribution
st.subheader("توزيع المعدلات حسب الشرائح")
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
            "0 - 9.99 (دون المعدل)": "#EF553B",
            "10 - 11.99 (متوسط)": "#FECB52",
            "12 - 20 (جيد/ممتاز)": "#00CC96"
        }
    )
    fig.update_traces(textposition='inside', textinfo='percent+value')
    fig.update_layout(height=400)
    st.plotly_chart(fig, use_container_width=True)

with col2:
    # Insights summary
    st.markdown("### 💡 أهم الملاحظات")
    total = len(df_filtered)
    
    # Success rate (>=10)
    success_rate = (len(average) + len(good)) / total * 100 if total > 0 else 0
    st.info(f"**نسبة النجاح (≥10):** {success_rate:.1f}% من التلاميذ ناجحون")
    
    # Excellence rate (>=12)
    excellence_rate = len(good) / total * 100 if total > 0 else 0
    st.success(f"**نسبة التميز (≥12):** {excellence_rate:.1f}% حصلوا على معدل جيد/ممتاز")
    
    # At-risk students
    at_risk_rate = len(below_avg) / total * 100 if total > 0 else 0
    if at_risk_rate > 0:
        st.warning(f"**تلاميذ يحتاجون دعماً (<10):** {at_risk_rate:.1f}% يحتاجون متابعة إضافية")
    
    # Performance summary
    if success_rate >= 80:
        st.markdown("✅ **الأداء العام:** ممتاز - معظم التلاميذ ناجحون")
    elif success_rate >= 60:
        st.markdown("⚠️ **الأداء العام:** جيد - الأغلبية ناجحون مع إمكانية التحسن")
    else:
        st.markdown("🚨 **الأداء العام:** يحتاج اهتماماً - كثير من التلاميذ يواجهون صعوبات")

# Students list by bracket
st.subheader("📋 التلاميذ حسب الشريحة")
bracket_tab1, bracket_tab2, bracket_tab3 = st.tabs(["🔴 دون المعدل", "🟡 متوسط", "🟢 جيد/ممتاز"])

with bracket_tab1:
    if len(below_avg) > 0:
        st.dataframe(below_avg[['اسم التلميذ', 'الفصل', 'المعدل']].sort_values('المعدل', ascending=False), use_container_width=True)
    else:
        st.success("لا يوجد تلاميذ في هذه الشريحة!")

with bracket_tab2:
    if len(average) > 0:
        st.dataframe(average[['اسم التلميذ', 'الفصل', 'المعدل']].sort_values('المعدل', ascending=False), use_container_width=True)
    else:
        st.info("لا يوجد تلاميذ في هذه الشريحة")

with bracket_tab3:
    if len(good) > 0:
        st.dataframe(good[['اسم التلميذ', 'الفصل', 'المعدل']].sort_values('المعدل', ascending=False), use_container_width=True)
    else:
        st.info("لا يوجد تلاميذ في هذه الشريحة")

st.markdown("---")

# Detailed Statistics by Subject
st.header("📚 إحصائيات حسب المادة")

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
st.header("📊 الرسوم البيانية")

col1, col2 = st.columns(2)

# Average grades by subject
with col1:
    st.subheader("متوسط المعدلات حسب المادة")
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
    st.subheader("توزيع المعدلات")
    fig = px.histogram(
        df_filtered,
        x='المعدل',
        nbins=20,
        color_discrete_sequence=['#636EFA']
    )
    fig.add_vline(df_filtered['المعدل'].mean(), line_dash="dash", line_color="red", 
                   annotation_text=f"المتوسط: {df_filtered['المعدل'].mean():.2f}")
    fig.update_layout(height=400)
    st.plotly_chart(fig, use_container_width=True)

st.markdown("---")

# Student Rankings
st.header("🏆 أفضل 10 تلاميذ حسب المعدل")
top_students = df_filtered[['اسم التلميذ', 'المعدل']].dropna().nlargest(10, 'المعدل')
st.dataframe(top_students.reset_index(drop=True), use_container_width=True)

st.markdown("---")

# Performance by Subject - Box Plot
st.header("📊 توزيع المعدلات حسب المادة")

st.markdown("""
**📖 كيفية قراءة هذا الرسم البياني:**
- **الصندوق** يوضح أين تقع معظم معدلات التلاميذ (50% الوسطى)
- **الخط داخل الصندوق** هو الوسيط (المعدل الأوسط)
- **الشعيرات** (الخطوط الممتدة من الصندوق) توضح نطاق المعدلات النموذجية
- **النقاط خارج** الشعيرات هي قيم شاذة (معدلات مرتفعة أو منخفضة بشكل غير عادي)
- **صندوق أطول** يعني تباين أكبر في المعدلات لتلك المادة
- **صندوق في موضع أعلى** يعني أداء عام أفضل في تلك المادة
""")

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
    
    # Add subject-specific insights
    st.markdown("### 📈 ملاحظات حول المواد")
    col1, col2 = st.columns(2)
    
    with col1:
        # Best performing subject
        best_subject = stats_df.loc[stats_df['المتوسط'].idxmax()]
        st.success(f"**أفضل مادة أداءً:** {best_subject['المادة']} (المتوسط: {best_subject['المتوسط']:.2f})")
        
        # Most consistent subject (lowest std dev)
        most_consistent = stats_df.loc[stats_df['الانحراف المعياري'].idxmin()]
        st.info(f"**الأكثر استقراراً:** {most_consistent['المادة']} (الانحراف المعياري: {most_consistent['الانحراف المعياري']:.2f})")
    
    with col2:
        # Subject needing attention
        worst_subject = stats_df.loc[stats_df['المتوسط'].idxmin()]
        st.warning(f"**تحتاج اهتماماً:** {worst_subject['المادة']} (المتوسط: {worst_subject['المتوسط']:.2f})")
        
        # Most varied subject (highest std dev)
        most_varied = stats_df.loc[stats_df['الانحراف المعياري'].idxmax()]
        st.info(f"**الأكثر تبايناً:** {most_varied['المادة']} (الانحراف المعياري: {most_varied['الانحراف المعياري']:.2f})")

st.markdown("---")

# Raw Data Table
st.header("📋 جميع بيانات التلاميذ")
st.dataframe(df_filtered[['ر.ت', 'رقم التلميذ', 'اسم التلميذ'] + subject_columns], 
             use_container_width=True, height=400)

# Download option
st.markdown("---")

col_csv, col_ppt = st.columns(2)

with col_csv:
    csv = df_filtered.to_csv(index=False)
    st.download_button(
        label="📥 تحميل البيانات كـ CSV",
        data=csv,
        file_name=f"student_data_statistics.csv",
        mime="text/csv"
    )

with col_ppt:
    st.subheader("📊 إنشاء عرض تقديمي")
    
    # Get all available classes
    all_classes = list(df['الفصل'].unique())
    
    # Option to combine all classes
    combine_all_classes = st.checkbox(
        "دمج جميع الفصول في عرض واحد",
        value=True,
        help="عند التفعيل، سيتم دمج بيانات جميع الفصول المختارة في إحصائيات موحدة"
    )
    
    # Multi-select for classes to include in presentation
    selected_classes_ppt = st.multiselect(
        "اختر الفصول للعرض التقديمي:",
        options=all_classes,
        default=all_classes,
        help="اختر الفصول التي تريد تضمينها في العرض التقديمي"
    )
    
    if len(selected_classes_ppt) == 0:
        st.warning("⚠️ الرجاء اختيار فصل واحد على الأقل")
    
    # Filter data for presentation based on selected classes
    df_ppt = df[df['الفصل'].isin(selected_classes_ppt)].copy()
    df_ppt = df_ppt.dropna(subset=['اسم التلميذ'])
    
    # Show summary of selection
    if len(selected_classes_ppt) > 0:
        if combine_all_classes:
            st.info(f"📋 سيتم دمج **{len(df_ppt)}** تلميذ من **{len(selected_classes_ppt)}** فصل/فصول في عرض واحد")
        else:
            st.info(f"📋 سيتم إنشاء عرض منفصل لكل فصل (**{len(selected_classes_ppt)}** فصل/فصول)")
    
    if st.button("📊 إنشاء العرض التقديمي (PPTX)", disabled=len(selected_classes_ppt) == 0):
        with st.spinner("جاري إنشاء العرض التقديمي..."):
            # Check Kaleido availability early and warn user
            try:
                import kaleido
                test_fig = go.Figure()
                test_fig.to_image(format="png", width=100, height=100)
                kaleido_available = True
            except Exception:
                kaleido_available = False
                st.warning("⚠️ تصدير الرسوم البيانية غير متاح على هذا الخادم. سيتم إنشاء العرض بدون الرسوم البيانية.")
            
            # Create presentation
            prs = Presentation()
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(7.5)
            
            # Helper function to add title slide
            def add_title_slide(prs, title, subtitle=""):
                slide_layout = prs.slide_layouts[6]  # Blank layout
                slide = prs.slides.add_slide(slide_layout)
                
                # Title
                title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(12.333), Inches(1.5))
                title_frame = title_box.text_frame
                title_para = title_frame.paragraphs[0]
                title_para.text = title
                title_para.font.size = Pt(44)
                title_para.font.bold = True
                title_para.alignment = PP_ALIGN.CENTER
                
                if subtitle:
                    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(4), Inches(12.333), Inches(1))
                    sub_frame = subtitle_box.text_frame
                    sub_para = sub_frame.paragraphs[0]
                    sub_para.text = subtitle
                    sub_para.font.size = Pt(24)
                    sub_para.alignment = PP_ALIGN.CENTER
                
                return slide
            
            # Helper function to add content slide
            def add_content_slide(prs, title):
                slide_layout = prs.slide_layouts[6]  # Blank layout
                slide = prs.slides.add_slide(slide_layout)
                
                # Title
                title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.333), Inches(0.8))
                title_frame = title_box.text_frame
                title_para = title_frame.paragraphs[0]
                title_para.text = title
                title_para.font.size = Pt(32)
                title_para.font.bold = True
                
                return slide
            
            # Check if Kaleido/Chrome is available for image export
            def check_kaleido_available():
                try:
                    import kaleido
                    # Try a simple test
                    test_fig = go.Figure()
                    test_fig.to_image(format="png", width=100, height=100)
                    return True
                except Exception:
                    return False
            
            KALEIDO_AVAILABLE = check_kaleido_available()
            
            # Helper to save plotly figure as image
            def fig_to_image(fig):
                if not KALEIDO_AVAILABLE:
                    return None
                try:
                    img_bytes = fig.to_image(format="png", width=900, height=500, scale=2)
                    return io.BytesIO(img_bytes)
                except Exception:
                    return None
            
            # Helper function to add table of contents slide
            def add_toc_slide(prs):
                slide_layout = prs.slide_layouts[6]  # Blank layout
                slide = prs.slides.add_slide(slide_layout)
                
                # Title
                title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.333), Inches(0.8))
                title_frame = title_box.text_frame
                title_para = title_frame.paragraphs[0]
                title_para.text = "📋 فهرس المحتويات"
                title_para.font.size = Pt(36)
                title_para.font.bold = True
                title_para.alignment = PP_ALIGN.CENTER
                
                # Table of contents items
                toc_items = [
                    "1. الإحصائيات العامة",
                    "2. توزيع شرائح المعدلات",
                    "3. متوسط المعدلات حسب المادة",
                    "4. توزيع المعدلات",
                    "5. توزيع المعدلات حسب المادة (مخطط صندوقي)",
                    "6. أفضل 10 تلاميذ",
                    "7. أهم الملاحظات"
                ]
                
                toc_box = slide.shapes.add_textbox(Inches(2), Inches(1.5), Inches(9), Inches(5))
                toc_frame = toc_box.text_frame
                toc_frame.word_wrap = True
                
                for item in toc_items:
                    p = toc_frame.add_paragraph()
                    p.text = item
                    p.font.size = Pt(24)
                    p.space_after = Pt(16)
                
                return slide
            
            # Function to generate slides for a dataset
            def generate_slides_for_data(prs, data_df, title_suffix="", class_names=None):
                if class_names is None:
                    class_names = selected_classes_ppt
                
                # Title slide
                if len(class_names) == 1:
                    classes_text = class_names[0]
                elif len(class_names) <= 3:
                    classes_text = ', '.join(class_names)
                else:
                    classes_text = f"{len(class_names)} فصول"
                
                add_title_slide(prs, f"📊 إحصائيات نتائج التلاميذ {title_suffix}".strip(), 
                               f"الفصول: {classes_text} | عدد التلاميذ: {len(data_df)}")
                
                # Table of Contents
                add_toc_slide(prs)
                
                # Overall Statistics
                slide = add_content_slide(prs, "📈 الإحصائيات العامة")
                
                stats_text = f"""
عدد التلاميذ: {len(data_df)}
المعدل العام: {data_df['المعدل'].mean():.2f}
أعلى معدل: {data_df['المعدل'].max():.2f}
أدنى معدل: {data_df['المعدل'].min():.2f}
الانحراف المعياري: {data_df['المعدل'].std():.2f}
عدد الفصول: {len(class_names)}
                """
                
                stats_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(5), Inches(4))
                stats_frame = stats_box.text_frame
                stats_frame.word_wrap = True
                for line in stats_text.strip().split('\n'):
                    p = stats_frame.add_paragraph()
                    p.text = line.strip()
                    p.font.size = Pt(24)
                    p.space_after = Pt(12)
                
                # Grade Brackets
                slide = add_content_slide(prs, "📊 توزيع شرائح المعدلات")
                
                below_avg_count = len(data_df[data_df['المعدل'] < 10])
                avg_count = len(data_df[(data_df['المعدل'] >= 10) & (data_df['المعدل'] < 12)])
                good_count = len(data_df[data_df['المعدل'] >= 12])
                total = len(data_df)
                
                brackets_text = f"""
🔴 دون المعدل (0 - 9.99): {below_avg_count} تلميذ ({below_avg_count/total*100:.1f}%)

🟡 متوسط (10 - 11.99): {avg_count} تلميذ ({avg_count/total*100:.1f}%)

🟢 جيد/ممتاز (12 - 20): {good_count} تلميذ ({good_count/total*100:.1f}%)

✅ نسبة النجاح (≥10): {(avg_count + good_count)/total*100:.1f}%
                """
                
                bracket_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(6), Inches(5))
                bracket_frame = bracket_box.text_frame
                bracket_frame.word_wrap = True
                for line in brackets_text.strip().split('\n'):
                    p = bracket_frame.add_paragraph()
                    p.text = line.strip()
                    p.font.size = Pt(22)
                    p.space_after = Pt(8)
                
                # Pie chart
                bracket_counts_ppt = pd.DataFrame({
                    'Bracket': ['دون المعدل (0-9.99)', 'متوسط (10-11.99)', 'جيد/ممتاز (12-20)'],
                    'Count': [below_avg_count, avg_count, good_count]
                })
                fig_pie = px.pie(bracket_counts_ppt, values='Count', names='Bracket',
                                color='Bracket',
                                color_discrete_map={
                                    'دون المعدل (0-9.99)': '#EF553B',
                                    'متوسط (10-11.99)': '#FECB52',
                                    'جيد/ممتاز (12-20)': '#00CC96'
                                })
                fig_pie.update_traces(textposition='inside', textinfo='percent+value')
                fig_pie.update_layout(showlegend=True, legend=dict(orientation="h", y=-0.1))
                
                img_stream = fig_to_image(fig_pie)
                if img_stream:
                    slide.shapes.add_picture(img_stream, Inches(6.5), Inches(1.5), width=Inches(6))
                
                # Average by Subject
                slide = add_content_slide(prs, "📚 متوسط المعدلات حسب المادة")
                
                stats_data_ppt = []
                for col in subject_columns:
                    if col in data_df.columns:
                        valid_data = data_df[col].dropna()
                        if len(valid_data) > 0:
                            stats_data_ppt.append({
                                'المادة': col,
                                'المتوسط': valid_data.mean(),
                                'الأعلى': valid_data.max(),
                                'الأقل': valid_data.min(),
                                'الانحراف المعياري': valid_data.std(),
                                'عدد الطلاب': len(valid_data)
                            })
                stats_df_ppt = pd.DataFrame(stats_data_ppt)
                
                fig_bar = px.bar(
                    stats_df_ppt.sort_values('المتوسط', ascending=True),
                    x='المتوسط',
                    y='المادة',
                    orientation='h',
                    color='المتوسط',
                    color_continuous_scale='Viridis'
                )
                fig_bar.update_layout(height=500, width=1100)
                
                img_stream = fig_to_image(fig_bar)
                if img_stream:
                    slide.shapes.add_picture(img_stream, Inches(1), Inches(1.3), width=Inches(11))
                
                # Grade Distribution Histogram
                slide = add_content_slide(prs, "📊 توزيع المعدلات")
                
                fig_hist = px.histogram(
                    data_df,
                    x='المعدل',
                    nbins=20,
                    color_discrete_sequence=['#636EFA']
                )
                fig_hist.add_vline(data_df['المعدل'].mean(), line_dash="dash", line_color="red",
                                  annotation_text=f"المتوسط: {data_df['المعدل'].mean():.2f}")
                fig_hist.update_layout(height=500, width=1100)
                
                img_stream = fig_to_image(fig_hist)
                if img_stream:
                    slide.shapes.add_picture(img_stream, Inches(1), Inches(1.3), width=Inches(11))
                
                # Box Plot
                slide = add_content_slide(prs, "📊 توزيع المعدلات حسب المادة (مخطط صندوقي)")
                
                subject_data_ppt_list = []
                for col in subject_columns:
                    if col in data_df.columns:
                        valid_data = data_df[col].dropna()
                        for grade in valid_data:
                            subject_data_ppt_list.append({'المادة': col, 'التقدير': grade})
                
                if subject_data_ppt_list:
                    subject_box_df_ppt = pd.DataFrame(subject_data_ppt_list)
                    fig_box = px.box(subject_box_df_ppt, x='المادة', y='التقدير', color='المادة')
                    fig_box.update_layout(height=500, width=1100, showlegend=False)
                    
                    img_stream = fig_to_image(fig_box)
                    if img_stream:
                        slide.shapes.add_picture(img_stream, Inches(1), Inches(1.3), width=Inches(11))
                
                # Top 10 Students
                slide = add_content_slide(prs, "🏆 أفضل 10 تلاميذ")
                
                top_10 = data_df[['اسم التلميذ', 'المعدل']].dropna().nlargest(10, 'المعدل')
                
                rows = len(top_10) + 1
                cols = 3
                table = slide.shapes.add_table(rows, cols, Inches(2), Inches(1.3), Inches(9), Inches(5)).table
                
                table.cell(0, 0).text = "الترتيب"
                table.cell(0, 1).text = "اسم التلميذ"
                table.cell(0, 2).text = "المعدل"
                
                for i, (idx, row) in enumerate(top_10.iterrows()):
                    table.cell(i+1, 0).text = str(i+1)
                    table.cell(i+1, 1).text = str(row['اسم التلميذ'])
                    table.cell(i+1, 2).text = f"{row['المعدل']:.2f}"
                
                for i in range(rows):
                    for j in range(cols):
                        cell = table.cell(i, j)
                        cell.text_frame.paragraphs[0].font.size = Pt(14)
                        cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                
                # Subject Insights
                slide = add_content_slide(prs, "💡 أهم الملاحظات")
                
                best_subject = stats_df_ppt.loc[stats_df_ppt['المتوسط'].idxmax()]
                worst_subject = stats_df_ppt.loc[stats_df_ppt['المتوسط'].idxmin()]
                most_consistent = stats_df_ppt.loc[stats_df_ppt['الانحراف المعياري'].idxmin()]
                most_varied = stats_df_ppt.loc[stats_df_ppt['الانحراف المعياري'].idxmax()]
                
                insights_text = f"""
✅ أفضل مادة أداءً: {best_subject['المادة']} (المتوسط: {best_subject['المتوسط']:.2f})

⚠️ مادة تحتاج اهتماماً: {worst_subject['المادة']} (المتوسط: {worst_subject['المتوسط']:.2f})

📊 المادة الأكثر استقراراً: {most_consistent['المادة']} (الانحراف المعياري: {most_consistent['الانحراف المعياري']:.2f})

📈 المادة الأكثر تبايناً: {most_varied['المادة']} (الانحراف المعياري: {most_varied['الانحراف المعياري']:.2f})

🎯 نسبة النجاح الإجمالية: {(avg_count + good_count)/total*100:.1f}%

🌟 نسبة التميز (≥12): {good_count/total*100:.1f}%
                """
                
                insights_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(12), Inches(5))
                insights_frame = insights_box.text_frame
                insights_frame.word_wrap = True
                for line in insights_text.strip().split('\n'):
                    p = insights_frame.add_paragraph()
                    p.text = line.strip()
                    p.font.size = Pt(24)
                    p.space_after = Pt(12)
                
                # Thank you slide
                add_title_slide(prs, "شكراً لكم!", "تم الإنشاء من لوحة إحصائيات التلاميذ")
            
            # Generate presentation based on combine option
            if combine_all_classes:
                # Combined presentation for all selected classes
                generate_slides_for_data(prs, df_ppt, "", selected_classes_ppt)
            else:
                # Separate sections for each class
                for i, class_name in enumerate(selected_classes_ppt):
                    class_df = df_ppt[df_ppt['الفصل'] == class_name].copy()
                    if len(class_df) > 0:
                        if i > 0:
                            # Add separator slide between classes
                            add_title_slide(prs, f"📚 {class_name}", f"الفصل {i+1} من {len(selected_classes_ppt)}")
                        generate_slides_for_data(prs, class_df, f"- {class_name}", [class_name])
            
            # Save presentation
            pptx_buffer = io.BytesIO()
            prs.save(pptx_buffer)
            pptx_buffer.seek(0)
            
            st.success("✅ تم إنشاء العرض التقديمي بنجاح!")
            st.download_button(
                label="📥 تحميل العرض التقديمي",
                data=pptx_buffer,
                file_name=f"student_statistics_{'_'.join(selected_classes_ppt)}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
