from __future__ import annotations

import os
import uuid
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from calculator import calculate_assets, create_result_workbook, create_template, read_assets


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}


def create_app() -> Flask:
    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024
    UPLOAD_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)

    @app.get("/")
    def index():
        return render_template("index.html")

    @app.get("/sablon-indir")
    def download_template():
        path = OUTPUT_DIR / "SABLON_SABIT_KIYMET_LISTESI.xlsx"
        create_template(path)
        return send_file(path, as_attachment=True, download_name="SABLON_SABIT_KIYMET_LISTESI.xlsx")

    @app.post("/hesapla")
    def calculate():
        uploaded = request.files.get("excel_file")
        if not uploaded or uploaded.filename == "":
            return jsonify(success=False, error="Excel dosyası seçilmedi."), 400

        extension = Path(uploaded.filename).suffix.lower()
        if extension not in ALLOWED_EXTENSIONS:
            return jsonify(success=False, error="Lütfen .xlsx formatında Excel dosyası yükleyin."), 400

        try:
            islem_yili = int(request.form.get("islem_yili", "2025"))
            donem = int(request.form.get("donem", "4"))
            yd_orani = _parse_rate(request.form.get("yd_orani", "0"))
        except ValueError:
            return jsonify(success=False, error="Yıl, dönem veya oran formatı hatalı."), 400

        token = uuid.uuid4().hex
        filename = secure_filename(uploaded.filename) or f"upload{extension}"
        upload_path = UPLOAD_DIR / f"{token}_{filename}"
        period_name = {1: "1Donem", 2: "2Donem", 3: "3Donem", 4: "Yillik"}.get(donem, "Yillik")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = OUTPUT_DIR / f"YD_Amortisman_Sonuc_{islem_yili}_{period_name}_{timestamp}_{token}.xlsx"
        uploaded.save(upload_path)

        try:
            assets = read_assets(upload_path)
            results = calculate_assets(assets, islem_yili, donem, yd_orani)
            summary = create_result_workbook(results, output_path, islem_yili, donem, yd_orani)
        except Exception as exc:
            return jsonify(success=False, error=str(exc)), 400

        return jsonify(
            success=True,
            download_url=url_for("download_result", file_id=output_path.name),
            **summary,
        )

    @app.get("/download/<file_id>")
    def download_result(file_id: str):
        path = OUTPUT_DIR / secure_filename(file_id)
        if not path.exists():
            return jsonify(success=False, error="Sonuç dosyası bulunamadı."), 404
        return send_file(path, as_attachment=True, download_name=file_id)

    return app


def _parse_rate(value: str) -> float:
    return float(str(value).strip().replace(",", "."))


app = create_app()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=True)
