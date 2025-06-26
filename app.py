import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
import io

# Función para crear PowerPoint más visual
def create_ppt_report(df, selected_column):
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    
    prs = Presentation()
    
    # === SLIDE 1: Portada profesional ===
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    
    title.text = "📊 PAID MEDIA REPORT"
    title.text_frame.paragraphs[0].font.size = Pt(44)
    title.text_frame.paragraphs[0].font.color.rgb = RGBColor(31, 56, 100)
    
    subtitle.text = f"Deep Analysis of {selected_column}\nGenerated automatically from your data"
    subtitle.text_frame.paragraphs[0].font.size = Pt(24)
    
    # === SLIDE 2: Dashboard de métricas clave ===
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Layout en blanco
    
    # Título del slide
    title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
    title_frame = title_shape.text_frame
    title_frame.text = "🎯 KEY PERFORMANCE METRICS"
    title_frame.paragraphs[0].font.size = Pt(32)
    title_frame.paragraphs[0].font.color.rgb = RGBColor(31, 56, 100)
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # Crear cajas de métricas estilo dashboard
    metrics = [
        ("TOTAL", f"{df[selected_column].sum():,.0f}", "💰"),
        ("AVERAGE", f"{df[selected_column].mean():,.0f}", "📊"),
        ("MAXIMUM", f"{df[selected_column].max():,.0f}", "🚀"),
        ("RECORDS", f"{len(df):,}", "📋")
    ]
    
    x_positions = [1, 3.5, 6, 8.5]
    for i, (label, value, emoji) in enumerate(metrics):
        # Caja de métrica
        metric_box = slide.shapes.add_textbox(Inches(x_positions[i]), Inches(2), Inches(2), Inches(2.5))
        metric_frame = metric_box.text_frame
        metric_frame.margin_left = Inches(0.1)
        metric_frame.margin_right = Inches(0.1)
        
        # Emoji
        p1 = metric_frame.paragraphs[0]
        p1.text = emoji
        p1.font.size = Pt(36)
        p1.alignment = PP_ALIGN.CENTER
        
        # Label
        p2 = metric_frame.add_paragraph()
        p2.text = label
        p2.font.size = Pt(14)
        p2.font.bold = True
        p2.font.color.rgb = RGBColor(100, 100, 100)
        p2.alignment = PP_ALIGN.CENTER
        
        # Value
        p3 = metric_frame.add_paragraph()
        p3.text = value
        p3.font.size = Pt(20)
        p3.font.bold = True
        p3.font.color.rgb = RGBColor(31, 56, 100)
        p3.alignment = PP_ALIGN.CENTER
    
    # === SLIDE 3: Gráfico por plataforma (si existe) ===
    if 'Platform' in df.columns:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # Título
        title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
        title_frame = title_shape.text_frame
        title_frame.text = f"📈 {selected_column.upper()} BY PLATFORM"
        title_frame.paragraphs[0].font.size = Pt(28)
        title_frame.paragraphs[0].font.color.rgb = RGBColor(31, 56, 100)
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # Datos para el gráfico
        platform_data = df.groupby('Platform')[selected_column].sum().sort_values(ascending=False)
        
        # Crear gráfico de barras
        chart_data = CategoryChartData()
        chart_data.categories = list(platform_data.index)
        chart_data.add_series('Values', list(platform_data.values))
        
        chart_shape = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(1), Inches(1.5), Inches(8), Inches(5),
            chart_data
        )
        chart = chart_shape.chart
        chart.has_legend = False
        
        # Personalizar gráfico
        chart.value_axis.has_major_gridlines = True
        chart.category_axis.tick_labels.font.size = Pt(12)
        chart.value_axis.tick_labels.font.size = Pt(12)
    
    # === SLIDE 4: Top Campaigns (si existe) ===
    if 'Campaign' in df.columns:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        title = slide.shapes.title
        title.text = f"🏆 TOP 5 CAMPAIGNS - {selected_column}"
        
        top_campaigns = df.groupby('Campaign')[selected_column].sum().sort_values(ascending=False).head(5)
        
        content = slide.placeholders[1]
        campaign_text = "🥇 " + "\n🥈 ".join([f"{camp}: {value:,.0f}" for camp, value in top_campaigns.items()])
        content.text = campaign_text
        content.text_frame.paragraphs[0].font.size = Pt(16)
    
    # === SLIDE 5: Resumen ejecutivo ===
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    title = slide.shapes.title
    title.text = "📋 EXECUTIVE SUMMARY"
    
    # Calcular insights automáticos
    total_platforms = df['Platform'].nunique() if 'Platform' in df.columns else 0
    date_range = f"{df['Date'].min()} to {df['Date'].max()}" if 'Date' in df.columns else 'N/A'
    best_platform = df.groupby('Platform')[selected_column].sum().idxmax() if 'Platform' in df.columns else 'N/A'
    
    content = slide.placeholders[1]
    content.text = f"""📊 ANALYSIS PERIOD: {date_range}
    
🎯 TOTAL {selected_column.upper()}: {df[selected_column].sum():,.0f}

🏆 TOP PERFORMING PLATFORM: {best_platform}

📈 PLATFORMS ANALYZED: {total_platforms}

💡 RECOMMENDATION: Focus budget optimization on top-performing platforms and campaigns for maximum ROI."""
    
    content.text_frame.paragraphs[0].font.size = Pt(16)
    content.text_frame.paragraphs[0].line_spacing = 1.5
    
    # Guardar en memoria
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io

# Configuración de la página
st.set_page_config(page_title="Paid Media Report", page_icon="📊")

# Título principal
st.title("📊 Paid Media Report Processor")
st.write("Upload your Excel file with paid media data for quick analysis")

# Widget para subir archivo
uploaded_file = st.file_uploader(
    "Choose your Excel file", 
    type=['xlsx', 'xls'],
    help="File should contain columns like: Campaign, Impressions, Clicks, Cost"
)

# Si hay archivo subido
if uploaded_file is not None:
    try:
        # Leer el Excel
        df = pd.read_excel(uploaded_file)
        
        # Información básica
        st.subheader("📋 Data Overview")
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Total Rows", len(df))
        with col2:
            st.metric("Columns", len(df.columns))
        
        st.write("**Columns found:**", list(df.columns))
        
        # Mostrar muestra de datos
        st.subheader("📊 Data Sample")
        st.dataframe(df.head(10))
        
        # Análisis básico si hay columnas numéricas
        numeric_columns = df.select_dtypes(include=['number']).columns
        
        if len(numeric_columns) > 0:
            st.subheader("📈 Quick Analysis")
            
            # Seleccionar columna para analizar
            selected_column = st.selectbox("Select a metric to analyze:", numeric_columns)
            
            if selected_column:
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total", f"{df[selected_column].sum():,.0f}")
                with col2:
                    st.metric("Average", f"{df[selected_column].mean():,.0f}")
                with col3:
                    st.metric("Max", f"{df[selected_column].max():,.0f}")
                
                # Gráfico simple
                st.subheader(f"📊 {selected_column} Distribution")
                fig, ax = plt.subplots(figsize=(10, 6))
                ax.hist(df[selected_column].dropna(), bins=20, alpha=0.7, color='skyblue')
                ax.set_xlabel(selected_column)
                ax.set_ylabel('Frequency')
                ax.set_title(f'Distribution of {selected_column}')
                st.pyplot(fig)
                
        # Descarga de datos procesados
        st.subheader("💾 Download Options")
        
        col1, col2 = st.columns(2)
        
        with col1:
            csv = df.to_csv(index=False)
            st.download_button(
                label="📄 Download as CSV",
                data=csv,
                file_name='processed_paid_media_data.csv',
                mime='text/csv'
            )
        
        with col2:
            # AQUÍ SE AGREGA EL BOTÓN DE POWERPOINT
            if len(numeric_columns) > 0 and 'selected_column' in locals():
                ppt_file = create_ppt_report(df, selected_column)
                st.download_button(
                    label="📊 Download PowerPoint Report",
                    data=ppt_file,
                    file_name=f'paid_media_report_{selected_column}.pptx',
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        st.write("Make sure it's a valid Excel file")

else:
    # Mostrar ejemplo de datos esperados
    st.info("👆 Upload an Excel file to get started")
    
    st.subheader("📝 Expected File Format:")
    sample_data = {
        'Campaign': ['Google Ads - Search', 'Facebook Ads', 'Instagram Ads', 'YouTube Ads'],
        'Impressions': [15000, 25000, 18000, 12000],
        'Clicks': [850, 1200, 950, 600],
        'Cost': [420.50, 680.25, 510.75, 290.30]
    }
    sample_df = pd.DataFrame(sample_data)
    st.dataframe(sample_df)
    
    st.write("**The app will automatically detect and analyze any numeric columns in your data.**")