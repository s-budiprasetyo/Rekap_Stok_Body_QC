"""
Auto-simpan file Excel hasil generate ke Google Drive, dan ambil kembali untuk
fitur 'Ambil Data'. Pakai service account yang sama dengan gsheets.py
(st.secrets['gcp_service_account']). Kalau belum dikonfigurasi, semua fungsi
di sini gagal dengan aman (return False/None) tanpa bikin app crash.

Opsional: isi `gdrive_folder_id` di Streamlit Secrets kalau mau file-file ini
disimpan di folder Google Drive tertentu (bukan My Drive root service account).
"""
import io
import datetime


def is_configured(st) -> bool:
    try:
        return "gcp_service_account" in st.secrets
    except Exception:
        return False


def _get_service(st):
    from googleapiclient.discovery import build
    from google.oauth2.service_account import Credentials

    scopes = ["https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]), scopes=scopes)
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def _folder_id(st):
    try:
        return st.secrets.get("gdrive_folder_id")
    except Exception:
        return None


def upload_excel(st, xlsx_bytes: bytes, tanggal_str: str, sesi: str) -> tuple[bool, str]:
    """Upload file Excel hasil generate ke Google Drive. Nama file menyertakan
    jam:menit:detik supaya beberapa file di tanggal & sesi yang sama tetap bisa
    dibedakan waktunya."""
    if not is_configured(st):
        return False, "Google Drive belum dikonfigurasi (simpan otomatis dilewati)."
    try:
        from googleapiclient.http import MediaIoBaseUpload

        service = _get_service(st)
        now = datetime.datetime.now()
        ts = now.strftime("%Y-%m-%d_%H-%M-%S")
        fname = f"Stok_Body_{tanggal_str.replace(' ', '_')}_{sesi}_{ts}.xlsx"
        metadata = {"name": fname}
        folder_id = _folder_id(st)
        if folder_id:
            metadata["parents"] = [folder_id]
        media = MediaIoBaseUpload(
            io.BytesIO(xlsx_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            resumable=False,
        )
        file = service.files().create(body=metadata, media_body=media, fields="id,name,createdTime").execute()
        return True, f"✅ Excel tersimpan otomatis ke Google Drive: {file.get('name')}"
    except Exception as e:
        return False, f"⚠️ Gagal simpan otomatis ke Google Drive: {e}"


def find_latest_file(st, tanggal_str: str, sesi: str):
    """Cari file Excel paling baru untuk tanggal+sesi tertentu. Return dict
    {'id','name','createdTime'} atau None kalau tidak ketemu."""
    if not is_configured(st):
        return None
    try:
        service = _get_service(st)
        prefix = f"Stok_Body_{tanggal_str.replace(' ', '_')}_{sesi}_"
        q = f"name contains '{prefix}' and trashed = false"
        folder_id = _folder_id(st)
        if folder_id:
            q += f" and '{folder_id}' in parents"
        resp = service.files().list(
            q=q, orderBy="createdTime desc", pageSize=5,
            fields="files(id,name,createdTime)",
        ).execute()
        files = resp.get("files", [])
        return files[0] if files else None
    except Exception:
        return None


def download_file(st, file_id: str) -> bytes:
    from googleapiclient.http import MediaIoBaseDownload

    service = _get_service(st)
    request = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buf.getvalue()
