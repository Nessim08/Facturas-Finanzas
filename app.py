import os
import io
import zipfile
import extract_msg
from email import policy
from email.parser import BytesParser
from flask import Flask, request, send_file, render_template

app = Flask(__name__)

# Carpeta temporal para guardar archivos .msg antes de procesar
UPLOAD_FOLDER = "/tmp/uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def procesar_eml(fp, zipf):
    """
    Lee un archivo .eml desde el file‑pointer `fp`
    y añade sus PDFs adjuntos al ZipFile `zipf`.
    """
    msg = BytesParser(policy=policy.default).parse(fp)
    for part in msg.iter_attachments():
        nombre = part.get_filename() or ""
        if nombre.lower().endswith(".pdf"):
            zipf.writestr(nombre, part.get_content())

def procesar_msg(path, zipf):
    """
    Lee un archivo .msg desde `path` en disco
    y añade sus PDFs adjuntos al ZipFile `zipf`.
    """
    msg = extract_msg.Message(path)
    for att in msg.attachments:
        nombre = att.longFilename or att.shortFilename or ""
        if nombre.lower().endswith(".pdf"):
            zipf.writestr(nombre, att.data)

@app.route("/", methods=("GET", "POST"))
def index():
    if request.method == "POST":
        # Construir un ZIP en memoria
        mem_zip = io.BytesIO()
        with zipfile.ZipFile(mem_zip, mode="w", compression=zipfile.ZIP_DEFLATED) as z:
            for f in request.files.getlist("files"):
                fname = f.filename.lower()
                if fname.endswith(".eml"):
                    procesar_eml(f.stream, z)
                elif fname.endswith(".msg"):
                    tmp_path = os.path.join(UPLOAD_FOLDER, f.filename)
                    f.save(tmp_path)
                    procesar_msg(tmp_path, z)
                    os.remove(tmp_path)
        mem_zip.seek(0)
        return send_file(
            mem_zip,
            download_name="pdfs_extraidos.zip",
            as_attachment=True,
            mimetype="application/zip"
        )

    # GET → renderizar el formulario bonito
    return render_template("index.html")

if __name__ == "__main__":
    # En local puede servir con debug, en producción Render usará gunicorn
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)), debug=True)

