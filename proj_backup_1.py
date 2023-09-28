import logging
import os
import traceback
import json

from flask import Flask, render_template, jsonify, Response, request, send_file

from AutoPdf import redact
from app_log import init_logging
from database import Database
from excel_processor import ExcelProcessor

app = Flask(__name__)
app.config.from_object("config")

init_logging(logging.DEBUG)

database = Database(app.config["DATABASE_PATH"])
excel_processor = ExcelProcessor(app)

PDF_TEMPLATE = os.path.join(app.config["XL_TEMPLATE_PATH"], "disclaimer", "Disclaimer IP.pdf")

@app.get("/api/data")
def data_get():
    records, columns = database.get_projects()

    return Response(
        json.dumps(
            {
                "sheet_name": Database.PROJECTS_SHEET,
                "data": records,
                "columns": columns,
            },
        ),
        mimetype="application/json",
    )


@app.post("/api/data")
def data_post():
    request_data = request.get_json()
    resp = {"status": "ok"}

    try:
        print(request_data["data"])
        database.write_projects(request_data["data"])
    except Exception as e:
        traceback.print_exc()
        resp = {"status": "error", "error": str(e)}
    return Response(
        json.dumps(resp),
        mimetype="application/json",
    )


@app.post("/api/create_new_project_from_raw_report")
def create_new_project_from_raw_report():
    return excel_processor.create_new_project_from_raw_report(request)


@app.post("/api/reprocess_audit_report")
def reprocess_audit_report():
    return excel_processor.reprocess_audit_report(request)


@app.post("/api/reprocess_audit_report_with_defaults")
def reprocess_audit_report_with_defaults():
    return excel_processor.reprocess_audit_report(request, reset=True)


@app.post("/api/create_proposal_pdf")
def create_proposal_pdf():
    return excel_processor.create_proposal_pdf(request)


@app.post("/api/create_proposal_custom")
def create_proposal_custom():
    return excel_processor.create_proposal_custom(request)


@app.post("/api/process_audit_report_standalone")
def process_audit_report_standalone():
    return excel_processor.process_audit_report(request)


@app.post("/api/create_photos_check_list")
def create_photos_check_list():
   return excel_processor.create_photos_check_list(request)


@app.post("/api/create_project_scope_pdf")
def create_project_scope_pdf():
    return excel_processor.create_project_scope_pdf(request)


@app.post("/api/create_scope_pdf")
def create_scope_pdf():
    return excel_processor.create_scope_pdf(request)


@app.post("/api/create_clip")
def create_clip():
    return excel_processor.create_clip(request)


@app.post("/api/create_materials")
def create_materials():
    return excel_processor.create_materials(request)


@app.post("/api/create_materials_combined")
def create_materials_combined():
    return excel_processor.create_materials(request, combined=True)


@app.route("/")
def index():
    db_config = database.get_config()
    PROJECTS_FOLDER = db_config[Database.DEFAULT_PROJECTS_PATH].replace("\\", "\\\\")
    return render_template("main.html", PATH_COLUMN=Database.PROJECT_PATH_COL, PROJECTS_FOLDER=PROJECTS_FOLDER)


@app.route("/create")
def create():
    identifier = request.args.get('identifier')
    pdf_path = redact(
        db_path=app.config["DATABASE_PATH"], sheet_name=Database.PROJECTS_SHEET, header_row=Database.PROJECT_HEADER_ROW,
        template_path=PDF_TEMPLATE, identifier=identifier
    )
    return send_file(pdf_path)


