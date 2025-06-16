import streamlit as st
import pandas as pd
import locale
import io
from datetime import datetime

st.set_page_config(page_title="Deducciones IVA", page_icon="📊")

st.title("📊 Deducciones IVA")

# Campo para el nombre del contribuyente
contribuyente = st.text_input("Nombre del Contribuyente", key="contribuyente")

# Widget para seleccionar archivo
uploaded_file = st.file_uploader(
    "Selecciona el archivo Excel descargado desde mis retenciones en ARCA",
    type=["xlsx", "xls"],
)

if uploaded_file is not None and contribuyente:
    try:
        # Leer el archivo Excel
        df = pd.read_excel(uploaded_file)

        # Columnas a eliminar
        columnas_a_eliminar = [
            "Impuesto",
            "Descripción Impuesto",
            "Régimen",
            "Número Certificado",
            "Descripción Régimen",
            "Descripción Operación",
            "Fecha Registración DJ Ag.Ret.",
            "Fecha Comprobante",
        ]

        # Eliminar las columnas especificadas
        df = df.drop(columns=columnas_a_eliminar, errors="ignore")

        # Renombrar las columnas
        df = df.rename(
            columns={
                "Número Comprobante": "Nro Comprobante",
                "Importe Ret./Perc.": "Importe",
                "CUIT Agente Ret./Perc.": "CUIT",
                "Fecha Ret./Perc.": "Fecha",
                "Denominación o Razón Social": "Razón Social",
                "Descripción Comprobante": "Comprobante",
            }
        )

        # Convertir columnas a los formatos deseados
        df["Nro Comprobante"] = df["Nro Comprobante"].astype(str)
        df["CUIT"] = df["CUIT"].astype(str)
        df["Importe"] = pd.to_numeric(df["Importe"], errors="coerce").round(2)

        # Convertir fecha a datetime y luego a string con formato específico
        df["Fecha"] = pd.to_datetime(df["Fecha"], dayfirst=True).dt.strftime("%d/%m/%Y")

        # Ordenar por fecha de menor a mayor (convertir a datetime para ordenar correctamente)
        df = df.sort_values(
            by="Fecha",
            key=lambda x: pd.to_datetime(x, format="%d/%m/%Y", dayfirst=True),
        )

        # Obtener el mes y año de la primera fecha
        primera_fecha = pd.to_datetime(
            df["Fecha"].iloc[0], format="%d/%m/%Y", dayfirst=True
        )
        mes_anio = primera_fecha.strftime("%m-%Y")

        # Reordenar las columnas
        columnas_ordenadas = [
            "CUIT",
            "Razón Social",
            "Fecha",
            "Nro Comprobante",
            "Comprobante",
            "Importe",
        ]
        df = df[columnas_ordenadas]

        # Configurar el formato de números
        pd.options.display.float_format = "{:,.2f}".format

        # Crear DataFrame con la fila de total
        df_total = df.copy()
        fila_total = pd.DataFrame(
            {
                "CUIT": [""],
                "Razón Social": ["TOTAL"],
                "Fecha": [""],
                "Nro Comprobante": [""],
                "Comprobante": [""],
                "Importe": [df["Importe"].sum()],
            }
        )
        df_total = pd.concat([df_total, fila_total], ignore_index=True)

        # Mostrar información básica del DataFrame
        st.write("### Vista de los datos:")
        st.write(f"Total de registros: {len(df)}")
        st.dataframe(df_total, use_container_width=True, hide_index=True)

        # Botón para descargar el archivo Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df.to_excel(
                writer, sheet_name="Percepciones ARCA", index=False, startrow=4
            )  # Empezar en la fila 5
            worksheet = writer.sheets["Percepciones ARCA"]

            # Crear formato para el título
            formato_titulo = writer.book.add_format(
                {
                    "bold": True,
                    "align": "center",
                    "valign": "vcenter",
                    "border": 2,  # Borde grueso
                    "text_wrap": True,
                }
            )

            # Crear formato para el subtítulo
            formato_subtitulo = writer.book.add_format(
                {
                    "bold": True,
                    "align": "center",
                    "valign": "vcenter",
                    "font_size": 12,
                }
            )

            # Crear formato para la columna de importe
            formato_moneda = writer.book.add_format(
                {
                    "num_format": '"$ "#,##0.00;[Red]"$ "#,##0.00',
                    "align": "right",
                }
            )

            # Crear formato para la columna de fecha
            formato_fecha = writer.book.add_format(
                {
                    "num_format": "dd/mm/yyyy",
                    "align": "center",
                }
            )

            # Crear formato para la fila de total
            formato_total = writer.book.add_format(
                {
                    "bold": True,
                    "align": "center",
                    "valign": "vcenter",
                }
            )

            # Agregar el nombre del contribuyente en las primeras dos filas
            worksheet.merge_range(0, 0, 1, 5, contribuyente.upper(), formato_titulo)

            # Agregar el subtítulo con el mes y año
            worksheet.merge_range(
                3, 0, 3, 5, f"PERCEPCIONES IVA - {mes_anio}", formato_subtitulo
            )

            # Ajustar el ancho de las columnas
            for i, col in enumerate(df.columns):
                max_length = max(df[col].astype(str).apply(len).max(), len(col)) + 2
                worksheet.set_column(i, i, max_length)

            # Aplicar formato de moneda a la columna Importe
            worksheet.set_column(5, 5, None, formato_moneda)  # Columna F (índice 5)

            # Aplicar formato de fecha a la columna Fecha
            worksheet.set_column(2, 2, None, formato_fecha)  # Columna C (índice 2)

            # Agregar la fila de total con fórmula
            ultima_fila = len(df) + 5  # +5 porque empezamos en la fila 5

            # Combinar y centrar las celdas desde CUIT hasta Comprobante
            worksheet.merge_range(
                ultima_fila, 0, ultima_fila, 4, "TOTAL", formato_total
            )

            # Agregar la fórmula SUM con formato de moneda
            formula = f"=SUM(F5:F{ultima_fila})"
            worksheet.write_formula(ultima_fila, 5, formula, formato_moneda)

            # Ajustar la altura de las filas del título
            worksheet.set_row(0, 30)  # Altura de la primera fila
            worksheet.set_row(1, 30)  # Altura de la segunda fila
            worksheet.set_row(3, 25)  # Altura de la fila del subtítulo

        st.download_button(
            label="📥 Descargar Excel",
            data=buffer.getvalue(),
            file_name="datos_procesados.xlsx",
            mime="application/vnd.ms-excel",
        )

        # Mostrar información adicional
        st.write("### Información del archivo:")
        st.write(f"Número de filas: {len(df)}")
        st.write(f"Número de columnas: {len(df.columns)}")

        # Mostrar nombres de las columnas restantes
        st.write("### Columnas disponibles:")
        st.write(df.columns.tolist())

    except Exception as e:
        st.error(f"Error al leer el archivo: {str(e)}")
elif uploaded_file is not None and not contribuyente:
    st.error("Por favor, ingresa el nombre del contribuyente")
else:
    st.info(
        "Por favor, ingresa el nombre del contribuyente y selecciona un archivo Excel para comenzar."
    )
