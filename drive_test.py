import streamlit as st
import os
import io
import json

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# =========================
# CONFIG
# =========================

st.set_page_config(
    page_title="Drive Test",
    layout="wide"
)

st.title("🧪 Prueba Google Drive")

SCOPES = [
    "https://www.googleapis.com/auth/drive"
]

FOLDER_ID = "1vMT6gXgMU4TymjXiodWgoJC5murwCFgI"

# =========================
# CREDENCIALES
# =========================

credentials = (
    service_account.Credentials
    .from_service_account_file(
        "credenciales.json",
        scopes=SCOPES
    )
)

drive_service = build(
    "drive",
    "v3",
    credentials=credentials
)

st.success("✅ Conexión con Google Drive correcta")

# =========================
# SUBIR ARCHIVO
# =========================

archivo = st.file_uploader(
    "Subir archivo"
)

if archivo:

    temp_path = os.path.join(
        "/tmp",
        archivo.name
    )

    with open(
        temp_path,
        "wb"
    ) as f:

        f.write(
            archivo.getbuffer()
        )

    file_metadata = {
        "name": archivo.name,
        "parents": [FOLDER_ID]
    }

    media = MediaFileUpload(
        temp_path,
        resumable=True
    )

    uploaded_file = (
        drive_service.files()
        .create(
            body=file_metadata,
            media_body=media,
            fields="id, webViewLink"
        )
        .execute()
    )

    drive_link = uploaded_file[
        "webViewLink"
    ]

    st.success(
        "✅ Archivo subido correctamente"
    )

    st.write(drive_link)

    st.link_button(
        "📂 Abrir archivo",
        drive_link
    )

    os.remove(temp_path)

# =========================
# LISTAR ARCHIVOS
# =========================

st.markdown("## 📁 Archivos en carpeta")

results = (
    drive_service.files()
    .list(
        q=f"'{FOLDER_ID}' in parents and trashed=false",
        pageSize=10,
        fields="files(id, name, webViewLink)"
    )
    .execute()
)

files = results.get("files", [])

if files:

    for file in files:

        st.write(f"📄 {file['name']}")

        st.link_button(
            "Abrir",
            file["webViewLink"]
        )

else:

    st.info(
        "No hay archivos."
    )