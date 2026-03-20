import os
import uuid
import time
from flask import Flask, request, render_template, send_from_directory, flash, redirect, url_for
from pdf2docx import Converter
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "pdf2word_yhy_2025"
app.config["UPLOAD_FOLDER"] = "uploads"
app.config["OUTPUT_FOLDER"] = "outputs"
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
os.makedirs(app.config["OUTPUT_FOLDER"], exist_ok=True)

ALLOWED_EXTENSIONS = {"pdf"}

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def convert_pdf(pdf_path, docx_path):
    try:
        cv = Converter(pdf_path)
        cv.convert(
            docx_path,
            parse_images=True,
            parse_table=True,
            parse_layout=True,
            parse_header=True,
            parse_footer=True
        )
        cv.close()
        return True, "转换成功"
    except Exception as e:
        return False, str(e)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "file" not in request.files:
            flash("⚠️ 请选择PDF文件")
            return redirect(request.url)
        file = request.files["file"]
        if file.filename.strip() == "":
            flash("⚠️ 未选择任何文件")
            return redirect(request.url)
        if not allowed_file(file.filename):
            flash("⚠️ 仅支持上传 .pdf 格式文件")
            return redirect(request.url)

        try:
            filename = secure_filename(file.filename)
            unique_name = str(uuid.uuid4())
            pdf_path = os.path.join(app.config["UPLOAD_FOLDER"], unique_name + ".pdf")
            file.save(pdf_path)

            docx_path = os.path.join(app.config["OUTPUT_FOLDER"], unique_name + ".docx")
            ok, msg = convert_pdf(pdf_path, docx_path)

            if ok:
                return redirect(url_for("download_file", filename=unique_name + ".docx"))
            else:
                flash(f"❌ 转换失败：{msg[:100]}")
                return redirect(request.url)
        except Exception as e:
            flash(f"❌ 服务异常：{str(e)[:100]}")
            return redirect(request.url)

    return render_template("index.html")

@app.route("/download/<filename>")
def download_file(filename):
    return send_from_directory(app.config["OUTPUT_FOLDER"], filename, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)