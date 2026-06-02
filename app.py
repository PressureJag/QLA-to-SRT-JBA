import json
import os
import uuid
from pathlib import Path

from flask import (Flask, render_template, request, redirect,
                   url_for, session, send_from_directory, send_file, flash)
from werkzeug.utils import secure_filename

# Load .env if present
try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent / ".env")
except ImportError:
    pass

from analyser import deep_analyse_qla

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-change-in-prod")

UPLOAD_FOLDER  = Path(__file__).parent / "uploads"
OUTPUT_FOLDER  = Path(__file__).parent / "output"
UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {"xlsx", "xls"}


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


# ── Upload ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    if "qla_file" not in request.files:
        return redirect(url_for("index"))

    f = request.files["qla_file"]
    if not f.filename or not allowed_file(f.filename):
        return render_template("index.html", error="Please upload a valid .xlsx file.")

    filename    = secure_filename(f.filename)
    unique_name = f"{uuid.uuid4().hex}_{filename}"
    save_path   = UPLOAD_FOLDER / unique_name
    f.save(save_path)

    try:
        analysis = deep_analyse_qla(save_path)
    except Exception as e:
        save_path.unlink(missing_ok=True)
        return render_template("index.html", error=f"Could not read file: {e}")

    # Store only the file path in session; re-derive analysis on demand
    session["qla_path"] = str(save_path)

    return render_template("analysis.html", data=analysis)


# ── SRT generation ────────────────────────────────────────────────────────────

@app.route("/generate/srt", methods=["POST"])
def generate_srt():
    qla_path = session.get("qla_path")
    if not qla_path or not Path(qla_path).exists():
        return redirect(url_for("index"))

    # Re-run analysis from the stored file (avoids session size limits)
    try:
        analysis = deep_analyse_qla(qla_path)
    except Exception as e:
        return render_template("index.html", error=f"Could not re-read file: {e}")

    if not analysis.get("has_any_scores"):
        return render_template(
            "analysis.html",
            data=analysis,
            generate_error="No student score data found in this file. Upload a QLA file with scores to generate sheets.",
        )

    if not os.environ.get("ANTHROPIC_API_KEY"):
        return render_template(
            "analysis.html",
            data=analysis,
            generate_error="ANTHROPIC_API_KEY is not set. Add it to the .env file in the project folder and restart the app.",
        )

    selected_groups = request.form.getlist("groups") or list(analysis["groups"].keys())

    try:
        from generators.srt import generate_srt_sheets
        files = generate_srt_sheets(analysis, selected_groups)
    except Exception as e:
        return render_template("analysis.html", data=analysis, generate_error=f"Generation failed: {e}")

    if not files:
        return render_template(
            "analysis.html",
            data=analysis,
            generate_error="No classes with score data were found in the selected groups.",
        )

    return render_template("srt_results.html", files=files, analysis=analysis)


# ── Downloads ─────────────────────────────────────────────────────────────────

@app.route("/download/<path:filename>")
def download_file(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)


@app.route("/download-zip", methods=["POST"])
def download_zip():
    files_json = request.form.get("files_json", "[]")
    files      = json.loads(files_json)
    from generators.srt import create_zip
    zip_buf    = create_zip(files)
    return send_file(
        zip_buf,
        mimetype="application/zip",
        as_attachment=True,
        download_name="SRT_Intervention_Sheets.zip",
    )


if __name__ == "__main__":
    app.run(debug=True, port=5050)
