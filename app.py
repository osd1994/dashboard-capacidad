import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.io as pio
from datetime import datetime
import base64
import os
from fpdf import FPDF
import tempfile

# Configuración general
st.set_page_config(page_title="Dashboard de Capacidad", layout="wide")
st.title("📊 Dashboard de Plan de Capacidad de Proyectos")

# Subir archivo
uploaded_file = st.file_uploader("📎 Sube el archivo Excel de Plan de Capacidad", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name="Capacidad por Division")
    month_columns = [col for col in df.columns if '-' in col]

    df["Total Horas"] = df[month_columns].sum(axis=1)
    df["Promedio Mensual"] = df[month_columns].mean(axis=1)
    df["Sobreejecutado"] = df[month_columns].gt(160).any(axis=1)
    df["Alto Uso (+80%)"] = df["Promedio Mensual"] > 128

    df_summary = df.groupby(["Responsable", "Rol", "Division"]).agg(
        Total_Horas=("Total Horas", "sum"),
        Promedio_Horas_Mes=("Promedio Mensual", "mean")
    ).reset_index()

    st.sidebar.title("🔍 Navegación")
    opcion = st.sidebar.radio("Selecciona una vista", [
        "Resumen General",
        "Visualización por Persona",
        "Proyectos con Mayor Carga",
        "Personas Sobreejecutadas",
        "Personas con Uso ≥ 80%",
        "Tabla Completa"
    ])

    col1, col2, col3 = st.columns([1, 1, 2])
    with col1:
        selected_div = st.selectbox("📂 División", ["Todas"] + sorted(df_summary["Division"].dropna().unique()))
    with col2:
        selected_rol = st.selectbox("🧑‍💼 Rol", ["Todos"] + sorted(df_summary["Rol"].dropna().unique()))
    with col3:
        search_name = st.text_input("🔎 Buscar por nombre de persona")

    if "Proyecto" in df.columns:
        proyecto_list = df["Proyecto"].dropna().unique().tolist()
        selected_proj = st.selectbox("📁 Proyecto", ["Todos"] + sorted(proyecto_list))
    else:
        selected_proj = "Todos"

    filtered_df = df_summary.copy()
    if selected_div != "Todas":
        filtered_df = filtered_df[filtered_df["Division"] == selected_div]
    if selected_rol != "Todos":
        filtered_df = filtered_df[filtered_df["Rol"] == selected_rol]
    if search_name.strip() != "":
        filtered_df = filtered_df[filtered_df["Responsable"].str.contains(search_name, case=False, na=False)]
    if selected_proj != "Todos" and "Proyecto" in df.columns:
        responsables_del_proyecto = df[df["Proyecto"] == selected_proj]["Responsable"].unique()
        filtered_df = filtered_df[filtered_df["Responsable"].isin(responsables_del_proyecto)]

    if opcion == "Resumen General":
        st.subheader("📊 Panel Ejecutivo de Carga de Capacidad")
        kpi1, kpi2, kpi3 = st.columns(3)
        kpi1.metric("👥 Total Personas", df['Responsable'].nunique())
        kpi2.metric("🕐 Horas Totales", int(df["Total Horas"].sum()))
        kpi3.metric("📁 Total Proyectos", df["Proyecto"].nunique() if "Proyecto" in df.columns else "N/A")
        kpi4, kpi5 = st.columns(2)
        kpi4.metric("⚠ Personas Sobreejecutadas", df[df["Sobreejecutado"]]["Responsable"].nunique())
        kpi5.metric("🔶 Uso ≥80% en Promedio", df[df["Alto Uso (+80%)"]]["Responsable"].nunique())

        st.markdown("---")
        top_carga = df_summary.sort_values("Total_Horas", ascending=False).head(10)
        fig_top = px.bar(top_carga, x="Responsable", y="Total_Horas", color="Division",
                         title="🔝 Top 10 Personas con más carga")
        st.plotly_chart(fig_top, use_container_width=True)

        df_div = df.groupby("Division").agg(Horas_Totales=("Total Horas", "sum")).reset_index()
        fig_div = px.pie(df_div, names="Division", values="Horas_Totales", title="🏢 Carga total por División")
        st.plotly_chart(fig_div, use_container_width=True)

        df_mes = df[month_columns].sum().reset_index()
        df_mes.columns = ["Mes", "Horas"]
        fig_line = px.line(df_mes, x="Mes", y="Horas", title="📅 Carga Total Mensual")
        st.plotly_chart(fig_line, use_container_width=True)

        df_sobree = df[df["Sobreejecutado"]]
        if not df_sobree.empty:
            fig_bubble = px.scatter(df_sobree, x="Responsable", y="Total Horas", size="Promedio Mensual",
                                    color="Division", title="⚠ Personas Sobreejecutadas")
            st.plotly_chart(fig_bubble, use_container_width=True)

        if "Proyecto" in df.columns:
            st.markdown("---")
            st.markdown("### 🏗️ Proyectos con Mayor Carga")
            proyectos_df = df.groupby("Proyecto").agg(
                Horas_Totales=("Total Horas", "sum"),
                Personas_Involucradas=("Responsable", "nunique")
            ).sort_values(by="Horas_Totales", ascending=False).reset_index()

            fig_proj = px.bar(
                proyectos_df,
                x="Proyecto",
                y="Horas_Totales",
                color="Personas_Involucradas",
                title="Proyectos con mayor demanda de esfuerzo",
                height=400
            )
            st.plotly_chart(fig_proj, use_container_width=True)
            st.dataframe(proyectos_df, use_container_width=True)

        with tempfile.TemporaryDirectory() as tmpdirname:
            top_path = os.path.join(tmpdirname, "top_carga.png")
            div_path = os.path.join(tmpdirname, "carga_division.png")
            mes_path = os.path.join(tmpdirname, "carga_mensual.png")
            bubble_path = os.path.join(tmpdirname, "sobreejecutados.png")
            proj_path = os.path.join(tmpdirname, "proyectos.png")

            fig_top.write_image(top_path)
            fig_div.write_image(div_path)
            fig_line.write_image(mes_path)
            if not df_sobree.empty:
                fig_bubble.write_image(bubble_path)
            if "Proyecto" in df.columns:
                fig_proj.write_image(proj_path)

            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            pdf.set_font("Arial", 'B', 16)
            pdf.cell(200, 10, "Resumen de Carga de Capacidad", ln=True, align='C')
            pdf.set_font("Arial", size=12)
            pdf.cell(200, 10, datetime.today().strftime('%Y-%m-%d'), ln=True, align='C')

            for img_path in [top_path, div_path, mes_path, bubble_path, proj_path]:
                if os.path.exists(img_path):
                    pdf.add_page()
                    pdf.image(img_path, x=10, y=20, w=180)

            pdf_output = os.path.join(tmpdirname, "Resumen_Capacidad.pdf")
            pdf.output(pdf_output)

            with open(pdf_output, "rb") as f:
                pdf_data = f.read()
                b64 = base64.b64encode(pdf_data).decode()
                href = f'<a href="data:application/octet-stream;base64,{b64}" download="Resumen_Capacidad.pdf">📥 Descargar PDF del Resumen</a>'
                st.markdown(href, unsafe_allow_html=True)

    elif opcion == "Visualización por Persona":
        st.subheader("📈 Visualización de Carga por Persona")
        fig = px.bar(
            filtered_df,
            x="Responsable",
            y="Total_Horas",
            color="Promedio_Horas_Mes",
            color_continuous_scale="reds",
            title="Horas Totales Asignadas por Persona",
            labels={"Total_Horas": "Horas Totales", "Promedio_Horas_Mes": "Promedio Mensual"},
            height=500
        )
        st.plotly_chart(fig, use_container_width=True)

    elif opcion == "Proyectos con Mayor Carga" and "Proyecto" in df.columns:
        st.subheader("🏗️ Proyectos que más esfuerzo requieren")
        proyectos_df = df.groupby("Proyecto").agg(
            Horas_Totales=("Total Horas", "sum"),
            Personas_Involucradas=("Responsable", "nunique")
        ).sort_values(by="Horas_Totales", ascending=False).reset_index()

        fig_proj = px.bar(
            proyectos_df,
            x="Proyecto",
            y="Horas_Totales",
            color="Personas_Involucradas",
            title="Proyectos con mayor demanda de esfuerzo",
            height=400
        )
        st.plotly_chart(fig_proj, use_container_width=True)
        st.dataframe(proyectos_df, use_container_width=True)

    elif opcion == "Personas Sobreejecutadas":
        st.subheader("⚠️ Personas Sobreejecutadas (>160h en algún mes)")
        detalles = []
        for _, row in df.iterrows():
            for mes in month_columns:
                if row[mes] > 160:
                    detalles.append({
                        "Responsable": row["Responsable"],
                        "Rol": row["Rol"],
                        "División": row["Division"],
                        "Proyecto": row["Proyecto"] if "Proyecto" in row else "No especificado",
                        "Mes": mes,
                        "Horas": row[mes]
                    })
        if detalles:
            df_detalle = pd.DataFrame(detalles)
            st.dataframe(df_detalle, use_container_width=True)
            resumen = df_detalle.groupby(["Responsable", "Rol", "División"]).agg(
                Total_Meses_Sobreejecucion=("Mes", "nunique"),
                Total_Horas_Sobreejecutadas=("Horas", "sum")
            ).reset_index()
            st.markdown("### 📌 Resumen de personas sobreejecutadas")
            st.dataframe(resumen, use_container_width=True)
        else:
            st.success("✅ Nadie está sobreejecutado en ningún mes.")

    elif opcion == "Personas con Uso ≥ 80%":
        st.subheader("🔶 Personas con uso mensual ≥ 80% (más de 128h y hasta 160h)")
        detalles_uso_alto = []
        for _, row in df.iterrows():
            for mes in month_columns:
                if 128 < row[mes] <= 160:
                    detalles_uso_alto.append({
                        "Responsable": row["Responsable"],
                        "Rol": row["Rol"],
                        "División": row["Division"],
                        "Proyecto": row["Proyecto"] if "Proyecto" in row else "No especificado",
                        "Mes": mes,
                        "Horas": row[mes]
                    })
        if detalles_uso_alto:
            df_alto_uso = pd.DataFrame(detalles_uso_alto)
            st.dataframe(df_alto_uso, use_container_width=True)
            resumen_alto = df_alto_uso.groupby(["Responsable", "Rol", "División"]).agg(
                Meses_Alta_Carga=("Mes", "nunique"),
                Horas_Acumuladas=("Horas", "sum")
            ).reset_index()
            st.markdown("### 📌 Resumen de personas con alta carga mensual")
            st.dataframe(resumen_alto, use_container_width=True)
        else:
            st.success("✅ Nadie está utilizando más del 80% de su tiempo en ningún mes.")

    elif opcion == "Tabla Completa":
        st.subheader("📋 Tabla Completa de Datos")
        st.dataframe(df, use_container_width=True)
else:
    st.info("📥 Por favor sube un archivo Excel para comenzar.")





