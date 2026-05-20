#!/usr/bin/env python3
"""Web app wrapper for the ITR PDF to Computation converter."""

from __future__ import annotations

import tempfile
import os
from io import BytesIO
from pathlib import Path

from flask import Flask, Response, flash, redirect, render_template_string, request, send_file, url_for

from itr_pdf_to_computation import output_name, parse_itr_pdf, render_doc


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024
app.secret_key = os.environ.get("SECRET_KEY", "local-development-secret")


PAGE = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>ITR PDF to Computation</title>
  <style>
    :root {
      color-scheme: light;
      font-family: Arial, sans-serif;
      background: #f5f7fb;
      color: #172033;
    }
    body {
      margin: 0;
      min-height: 100vh;
      display: grid;
      place-items: center;
      padding: 24px;
    }
    main {
      width: min(720px, 100%);
      background: #fff;
      border: 1px solid #d8dfeb;
      border-radius: 8px;
      padding: 28px;
      box-shadow: 0 16px 40px rgba(23, 32, 51, 0.08);
    }
    h1 {
      margin: 0 0 8px;
      font-size: 28px;
    }
    p {
      margin: 0 0 22px;
      color: #516070;
      line-height: 1.5;
    }
    form {
      display: grid;
      gap: 14px;
    }
    input[type="file"] {
      border: 1px dashed #9aa8bc;
      border-radius: 8px;
      padding: 18px;
      background: #fafbfe;
    }
    button {
      justify-self: start;
      border: 0;
      border-radius: 6px;
      background: #1769e0;
      color: white;
      font-weight: 700;
      padding: 11px 18px;
      cursor: pointer;
    }
    button:hover {
      background: #1259c4;
    }
    .message {
      margin-bottom: 16px;
      padding: 10px 12px;
      border-radius: 6px;
      background: #fff4d7;
      color: #6b4b00;
    }
  </style>
</head>
<body>
  <main>
    <h1>ITR PDF to Computation</h1>
    <p>Upload an Income Tax portal PDF and download a Word-openable computation document.</p>
    {% with messages = get_flashed_messages() %}
      {% if messages %}
        {% for message in messages %}
          <div class="message">{{ message }}</div>
        {% endfor %}
      {% endif %}
    {% endwith %}
    <form method="post" action="{{ url_for('convert') }}" enctype="multipart/form-data">
      <input type="file" name="pdf" accept="application/pdf,.pdf" required>
      <button type="submit">Convert PDF</button>
    </form>
  </main>
</body>
</html>
"""


@app.get("/")
def index() -> str:
    return render_template_string(PAGE)


@app.post("/convert")
def convert() -> Response:
    upload = request.files.get("pdf")
    if not upload or not upload.filename:
        flash("Please choose a PDF file.")
        return redirect(url_for("index"))

    if not upload.filename.lower().endswith(".pdf"):
        flash("Please upload a PDF file.")
        return redirect(url_for("index"))

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = Path(tmpdir) / "input.pdf"
            upload.save(pdf_path)
            data = parse_itr_pdf(pdf_path)
            filename = output_name(data, pdf_path)
            document = render_doc(data).encode("utf-8")
    except Exception as exc:
        flash(f"Conversion failed: {exc}")
        return redirect(url_for("index"))

    return send_file(
        BytesIO(document),
        mimetype="application/msword",
        as_attachment=True,
        download_name=filename,
    )


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
