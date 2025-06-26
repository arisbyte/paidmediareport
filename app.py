import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# Configuración de la página
st.set_page_config(page_title="Paid Media Report Processor", page_icon="📊")

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
        st.subheader("💾 Download Processed Data")
        csv = df.to_csv(index=False)
        st.download_button(
            label="Download as CSV",
            data=csv,
            file_name='processed_paid_media_data.csv',
            mime='text/csv'
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