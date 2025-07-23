from flask import Flask, request, send_file, render_template_string
import os, io, zipfile, extract_msg
from email import policy
from email.parser import BytesParser

app = Flask(__name__)

UPLOAD_FOLDER = "/tmp/uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

TEMPLATE = """
<!doctype html><title>Extraer PDFs</title>
<h1>Sube tus .eml/.msg</h1>
<form method=post enctype=multipart/form-data>
  <input type=file name=files multiple>
  <input type=submit value=Procesar>
</form>
"""

def procesar_eml(fp, zipf):
    msg = BytesParser(policy=policy.default).parse(fp)
    for part in msg.iter_attachments():
        fn = part.get_filename() or ""
        if fn.lower().endswith(".pdf"):
            zipf.writestr(fn, part.get_content())

def procesar_msg(path, zipf):
    msg = extract_msg.Message(path)
    for att in msg.attachments:
        fn = att.longFilename or att.shortFilename or ""
        if fn.lower().endswith(".pdf"):
            zipf.writestr(fn, att.data)

@app.route("/", methods=("GET","POST"))
def index():
    if request.method=="POST":
        # crea un ZIP en memoria
        mem_zip = io.BytesIO()
        with zipfile.ZipFile(mem_zip, "w") as z:
            for f in request.files.getlist("files"):
                name = f.filename.lower()
                if name.endswith(".eml"):
                    procesar_eml(f.stream, z)
                elif name.endswith(".msg"):
                    tmp = os.path.join(UPLOAD_FOLDER, f.filename)
                    f.save(tmp)
                    procesar_msg(tmp, z)
                    os.remove(tmp)
        mem_zip.seek(0)
        return send_file(mem_zip,
                         download_name="pdfs_extraidos.zip",
                         as_attachment=True)
    return render_template_string(TEMPLATE)

if __name__=="__main__":
    app.run(host="0.0.0.0", port=10000)
