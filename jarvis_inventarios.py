from pdf2image import convert_from_bytes
import cloudinary
import cloudinary.uploader
from google.oauth2 import service_account

from googleapiclient.discovery import build

from googleapiclient.http import MediaFileUpload
import streamlit as st
import pandas as pd
import os
import io
cloudinary.config(
    cloud_name=os.environ[
        "CLOUDINARY_CLOUD_NAME"
    ],
    api_key=os.environ[
        "CLOUDINARY_API_KEY"
    ],
    api_secret=os.environ[
        "CLOUDINARY_API_SECRET"
    ]
)
SCOPES = [
    "https://www.googleapis.com/auth/drive"
]
FOLDER_ID = (
    "1vMT6gXgMU4TymjXiodWgoJC5murwCFgI"
)
import json

google_credentials = json.loads(
    os.environ[
        "GOOGLE_CREDENTIALS"
    ]
)

credentials = (
    service_account.Credentials
    .from_service_account_info(
        google_credentials,
        scopes=SCOPES
    )
)
drive_service = build(
    "drive",
    "v3",
    credentials=credentials
)

from datetime import datetime
REGISTRO_DOCS = (
    "registro_documentos.xlsx"
)
# =========================
# CONFIGURACIÓN
# =========================

st.set_page_config(
    page_title="JARVIS INVENTARIOS 2025",
    layout="wide"
)

# =========================
# ESTILO
# =========================

st.markdown("""
<style>

/* APP */
.stApp{
    background-color:#F4F6F9;
    font-family:'Segoe UI',sans-serif;
}

/* TITULOS */
.main-title{
    font-size:42px;
    font-weight:900;
    color:#9F2241;
}

.sub-title{
    font-size:18px;
    color:#235B4E;
    margin-top:-10px;
}

/* SIDEBAR */
section[data-testid="stSidebar"]{
    background:linear-gradient(
        180deg,
        #235B4E 0%,
        #163832 100%
    );
}

section[data-testid="stSidebar"] *{
    color:white !important;
}

/* KPI */
.kpi-card{
    background:white;
    padding:20px;
    border-radius:18px;
    box-shadow:0 4px 12px rgba(0,0,0,0.08);
    border-left:8px solid #235B4E;
}

.kpi-title{
    font-size:16px;
    font-weight:700;
    color:#666;
}

.kpi-value{
    font-size:34px;
    font-weight:900;
    color:#1E1E1E;
}

/* TABLAS */
[data-testid="stDataFrame"]{
    background:white;
    border-radius:16px;
    padding:10px;
}

/* HEADER TABLAS */
thead tr th{
    background-color:#235B4E !important;
    color:white !important;
    font-weight:bold !important;
}

</style>
""", unsafe_allow_html=True)
# =========================
# LOGIN
# =========================

USUARIOS = {
    "admin": "jarvis2025",
    "consulta": "inventarios2025"
}

if "autenticado" not in st.session_state:

    st.session_state.autenticado = False

# =========================
# PANTALLA LOGIN
# =========================

if not st.session_state.autenticado:

    st.markdown("""
    <h1 style='text-align:center;color:#9F2241;'>
    🔒 JARVIS INVENTARIOS 2025
    </h1>
    """, unsafe_allow_html=True)

    st.markdown("##")

    c1, c2, c3 = st.columns([1,2,1])

    with c2:

        usuario = st.text_input(
            "Usuario"
        )

        password = st.text_input(
            "Contraseña",
            type="password"
        )

        entrar = st.button(
            "Ingresar"
        )

        if entrar:

            if (
                usuario in USUARIOS
                and USUARIOS[usuario] == password
            ):

                st.session_state.autenticado = True

                st.rerun()

            else:

                st.error(
                    "Usuario o contraseña incorrectos."
                )

    st.stop()

# =========================
# HEADER
# =========================

col1, col2 = st.columns([6,1])

with col1:

    st.markdown(
        '<div class="main-title">🧠 JARVIS INVENTARIOS 2025</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="sub-title">Centro de Monitoreo y Seguimiento Documental</div>',
        unsafe_allow_html=True
    )

with col2:

    st.success("🟢 ACTIVO")

# =========================
# SIDEBAR
# =========================

st.sidebar.markdown("# 🧠 JARVIS")

menu = st.sidebar.radio(
    "Navegación",
    [
        "🏠 Dashboard",
        "📍 Entidades",
        "🚨 Riesgos",
        "📎 Documentos",
        "📊 Analítica"
    ]
)

# =========================
# CARGA EXCEL
# =========================

st.sidebar.markdown("---")

EXCEL_FILE_ID = (
    "1cKS4CA5m7m_aQ3ncWQ9Gl3WOZHAKXVp3"
)

excel_request = (
    drive_service.files()
    .get_media(
        fileId=EXCEL_FILE_ID
    )
)

excel_data = (
    excel_request.execute()
)

archivo_excel = io.BytesIO(
    excel_data
)
# =========================
# VALIDACIÓN
# =========================


# =========================
# LEER EXCEL
# =========================

df = pd.read_excel(
    archivo_excel
)

df.columns = (
    df.columns
    .astype(str)
    .str.strip()
)

# =========================
# LIMPIEZA
# =========================

for col in df.columns:

    df[col] = (
        df[col]
        .astype(str)
        .str.strip()
    )

# =========================
# COLUMNAS
# =========================

col_entidad = "ENTIDAD"
col_clues = "CLUES"
col_almacen = "ALMACÉN"

# =========================
# OBSERVACIONES
# =========================

if "OBSERVACIONES DASHBOARD" not in df.columns:

    df["OBSERVACIONES DASHBOARD"] = ""

# =========================
# ESTATUS AUTOMÁTICO
# =========================

def calcular_estatus(row):

    carpeta = str(
        row.get(
            "CARPETA FÍSCA (Si/no)",
            ""
        )
    ).strip().upper()

    correcto = str(
        row.get(
            "CORRECTO/INCORRECTO",
            ""
        )
    ).strip().upper()

    # 🔵 EN CORRECCIÓN
    if (
        (
            "SI" in carpeta
            or "SÍ" in carpeta
        )
        and "INCORRECTO" in correcto
    ):

        return "🔵 En corrección"

    # 🟢 COMPLETO
    elif (
        (
            "SI" in carpeta
            or "SÍ" in carpeta
        )
        and correcto == "CORRECTO"
    ):

        return "🟢 Completo"

    # ⚫ SIN CARPETA
    elif "NO" in carpeta:

        return "⚫ Sin carpeta física"

    # ⚪ REVISAR
    else:

        return "⚪ Revisar"
# =========================
# GENERAR ESTATUS
# =========================

df["ESTATUS"] = (
    df.apply(
        calcular_estatus,
        axis=1
    )
)
# =========================
# DASHBOARD
# =========================

if menu == "🏠 Dashboard":

    st.markdown("## 🇲🇽 Dashboard Nacional")

    total = len(df)

    completos = (
        df["ESTATUS"]
        .str.contains("Completo")
        .sum()
    )

    sin_carpeta = (
        df["ESTATUS"]
        .str.contains("Sin carpeta")
        .sum()
    )
    en_correccion = (
        df["ESTATUS"]
        .str.contains("En corrección")
        .sum()
    )

    k1, k2, k3, k4 = st.columns(4)

    with k1:

        st.markdown(f"""
        <div class="kpi-card">
        <div class="kpi-title">
        🏥 Total registros
        </div>

        <div class="kpi-value">
        {total:,}
        </div>
        </div>
        """, unsafe_allow_html=True)

    with k2:

        st.markdown(f"""
        <div class="kpi-card">
        <div class="kpi-title">
        🟢 Completos
        </div>

        <div class="kpi-value">
        {completos:,}
        </div>
        </div>
        """, unsafe_allow_html=True)

    with k3:

        st.markdown(f"""
        <div class="kpi-card">
        <div class="kpi-title">
        🔵 En corrección
        </div>

        <div class="kpi-value">
        {en_correccion:,}
        </div>
        </div>
        """, unsafe_allow_html=True)


    with k4:

        st.markdown(f"""
        <div class="kpi-card">
        <div class="kpi-title">
        ⚫ Sin carpeta física
        </div>

        <div class="kpi-value">
        {sin_carpeta:,}
        </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("## 📋 Seguimiento nacional")

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True
    )# =========================
# ENTIDADES
# =========================

if menu == "📍 Entidades":

    estados = sorted(
        df[col_entidad]
        .dropna()
        .unique()
    )

    estado_sel = st.selectbox(
        "Entidad",
        estados
    )

    df_estado = (
        df[
            df[col_entidad]
            == estado_sel
        ]
        .copy()
    )

    st.markdown(
        f"## 📍 {estado_sel}"
    )

    st.dataframe(
        df_estado,
        use_container_width=True,
        hide_index=True
    )
    st.markdown(
        "## 📚 Expediente documental"
    )

    if os.path.exists(
        REGISTRO_DOCS
    ):

        historial_entidad = pd.read_excel(
            REGISTRO_DOCS
        )

        historial_entidad = historial_entidad[
            historial_entidad["Entidad"]
            == estado_sel
        ]

        if len(historial_entidad) > 0:

            st.dataframe(
                historial_entidad,
                use_container_width=True,
                hide_index=True
            )

        else:

            st.info(
                "No hay documentos cargados."
            )

    st.markdown(
        "## 📝 Observaciones"
    )

    fila = st.selectbox(
        "Seleccionar CLUES",
        df_estado[col_clues]
    )

    texto = st.text_area(
        "Agregar observación"
    )

    if st.button(
        "💾 Guardar observación"
    ):

        idx = df[
            (df[col_entidad] == estado_sel)
            &
            (df[col_clues] == fila)
        ].index

        fecha = datetime.now().strftime(
            "%d/%m/%Y %H:%M"
        )

        observacion = (
            f"[{fecha}] "
            + texto
        )

        df.loc[
            idx,
            "OBSERVACIONES DASHBOARD"
        ] = observacion

        st.success(
            "Observación guardada correctamente."
        )
# =========================
# DOCUMENTOS
# =========================

if menu == "📎 Documentos":

    st.markdown(
        "## 📎 Gestor documental"
    )

    # =========================
    # FILTROS
    # =========================

    entidad_doc = st.selectbox(
        "Entidad",
        sorted(
            df[col_entidad]
            .dropna()
            .unique()
        )
    )

    df_doc = (
        df[
            df[col_entidad]
            == entidad_doc
        ]
        .copy()
    )

    clues_doc = st.selectbox(
        "CLUES",
        sorted(
            df_doc[col_clues]
            .dropna()
            .unique()
        )
    )

    tipo_doc = st.radio(
        "Tipo documental",
        [
            "🟢 Entrega inventario",
            "🔵 Corrección",
            "🟡 Primer reiterativo",
            "🟠 Segundo reiterativo",
            "🔴 Tercer reiterativo",
            "🏛 Entrega UAS/OIC",
            "📧 Correo",
            "📎 Otro"
        ]
    )

    # =========================
    # SUBIR DOCUMENTO
    # =========================

    fecha_oficio = st.date_input(
        "📅 Fecha del oficio",
        value=None
    )

    archivo_doc = st.file_uploader(
        "Subir documento",
        type=[
            "pdf",
            "docx",
            "xlsx",
            "msg",
            "png",
            "jpg"
        ]
    )

    if archivo_doc and fecha_oficio:

        preview_url = ""

        if archivo_doc.name.lower().endswith(
            ".pdf"
        ):

            paginas = convert_from_bytes(
                archivo_doc.getvalue(),
                first_page=1,
                last_page=1
            )

            imagen = paginas[0]

            temp_img = (
                "preview.png"
            )

            imagen.save(
                temp_img,
                "PNG"
            )

            resultado = (
                cloudinary.uploader.upload(
                    temp_img,
                    folder="jarvis_previews"
                )
            )

            preview_url = (
                resultado["secure_url"]
            )

            os.remove(
                temp_img
            )

        nombre_limpio = (
            tipo_doc
            .replace("🟢", "")
            .replace("🔵", "")
            .replace("🟡", "")
            .replace("🟠", "")
            .replace("🔴", "")
            .replace("🏛", "")
            .replace("📧", "")
            .replace("📎", "")
            .strip()
            .replace(" ", "_")
        )

        fecha_doc = fecha_oficio.strftime(
            "%Y-%m-%d"
        )

        nombre_original = (
            archivo_doc.name
            .replace(" ", "_")
        )

        nombre = (
            fecha_doc
            + "_"
            + nombre_limpio
            + "_"
            + nombre_original
        )

        carpeta_destino = os.path.join(
            "documentos",
            entidad_doc,
            clues_doc
        )

        os.makedirs(
            carpeta_destino,
            exist_ok=True
        )

        ruta_archivo = os.path.join(
            carpeta_destino,
            nombre
        )

        with open(
            ruta_archivo,
            "wb"
        ) as f:

            f.write(
                archivo_doc.getbuffer()
            )

        st.success(

            "✅ Documento guardado correctamente"
        )

        st.write("📂 Archivo guardado:")

        st.code(
            ruta_archivo
        )
        nuevo_registro = pd.DataFrame([
            {
                "Entidad": entidad_doc,
                "CLUES": clues_doc,
                "Tipo": tipo_doc,
                "Fecha oficio": fecha_doc,
                "Archivo": nombre,
                "Ruta": ruta_archivo,
                "Preview": preview_url
            }
        ])
        if preview_url:

            if st.button(
                "👁 Vista previa"
            ):

                st.image(
                    preview_url,
                    caption="Vista previa documento"
                )
        if os.path.exists(
            REGISTRO_DOCS
        ):

            historial = pd.read_excel(
                REGISTRO_DOCS
            )

            historial = pd.concat(
                [
                    historial,
                    nuevo_registro
                ],
                ignore_index=True
            )

        else:

            historial = nuevo_registro

        historial.to_excel(
            REGISTRO_DOCS,
            index=False
        )
    st.markdown(
        "## 📚 Expediente documental"
    )

    if os.path.exists(
        REGISTRO_DOCS
    ):

        historial_docs = pd.read_excel(
            REGISTRO_DOCS
        )

        historial_docs = historial_docs[
            historial_docs["CLUES"]
            == clues_doc
        ]

        if len(historial_docs) > 0:

            st.dataframe(
                historial_docs,
                use_container_width=True,
                hide_index=True
            )

            for i, row in historial_docs.iterrows():

                c1, c2 = st.columns([8,1])

                with c1:

                    with open(
                        row["Ruta"],
                        "rb"
                    ) as file:

                        st.download_button(
                            label=f"📂 {row['Archivo']}",
                            data=file,
                            file_name=row["Archivo"],
                            mime="application/pdf",
                            key=f"down_{i}"
                        )

                with c2:

                    if st.button(
                        "🗑",
                        key=f"del_{i}"
                    ):

                        if os.path.exists(
                            row["Ruta"]
                        ):

                            os.remove(
                                row["Ruta"]
                            )

                        historial_docs = historial_docs.drop(i)

                        historial_docs.to_excel(
                            REGISTRO_DOCS,
                            index=False
                        )

                        st.success(
                            "Documento eliminado."
                        )

                        st.rerun()

    elif archivo_doc and not fecha_oficio:

        st.warning(
            "⚠ Selecciona la fecha del oficio."
        )