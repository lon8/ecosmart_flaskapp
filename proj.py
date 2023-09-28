import json
import logging
import os
import traceback
from decouple import config
from flask import Flask, render_template, Response, request, send_file
from flask_bcrypt import Bcrypt
from flask_login import LoginManager
from flask_mail import Mail
from flask_migrate import Migrate
from flask_sqlalchemy import SQLAlchemy

from AutoPdf import redact
from accounts.utils.auth import login_required
from accounts.utils.decorators import check_is_confirmed
from app_log import init_logging
from database import Database
from excel_processor import ExcelProcessor

app = Flask(__name__)
app.config.from_object(config("APP_SETTINGS"))
#print(app)

login_manager = LoginManager()
login_manager.init_app(app)
bcrypt = Bcrypt(app)
mail = Mail(app)
db = SQLAlchemy(app)
migrate = Migrate(app, db)

init_logging(logging.DEBUG)

database = Database(config("DATABASE_PATH"))
excel_processor = ExcelProcessor(app)

PDF_TEMPLATE = os.path.join(app.config["XL_TEMPLATE_PATH"], "disclaimer", "Disclaimer IP.pdf")


# Registering blueprints
from accounts.views import accounts_bp

app.register_blueprint(accounts_bp)
# app.register_blueprint(app)

from accounts.models import User

login_manager.login_view = "accounts.login"
login_manager.login_message_category = "danger"


@login_manager.user_loader
def load_user(user_id):
    return User.query.filter(User.id == int(user_id)).first()


"""ERRORS"""


@app.errorhandler(401)
def unauthorized_page(error):
    return render_template("errors/401.html"), 401


@app.errorhandler(403)
def access_denied(error):
    return render_template("errors/403.html"), 403


@app.errorhandler(404)
def page_not_found(error):
    return render_template("errors/404.html"), 404


@app.errorhandler(500)
def server_error_page(error):
    return render_template("errors/500.html"), 500


"""API"""


@app.get("/api/data")
@login_required
@check_is_confirmed
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
@login_required
@check_is_confirmed
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
@login_required
@check_is_confirmed
def create_new_project_from_raw_report():
    return excel_processor.create_new_project_from_raw_report(request)


@app.post("/api/reprocess_audit_report")
@login_required
@check_is_confirmed
def reprocess_audit_report():
    return excel_processor.reprocess_audit_report(request)


@app.post("/api/reprocess_audit_report_with_defaults")
@login_required
@check_is_confirmed
def reprocess_audit_report_with_defaults():
    return excel_processor.reprocess_audit_report(request, reset=True)


@app.post("/api/create_proposal_pdf")
@login_required
@check_is_confirmed
def create_proposal_pdf():
    return excel_processor.create_proposal_pdf(request)


@app.post("/api/create_proposal_custom")
@login_required
@check_is_confirmed
def create_proposal_custom():
    return excel_processor.create_proposal_custom(request)


@app.post("/api/process_audit_report_standalone")
@login_required
@check_is_confirmed
def process_audit_report_standalone():
    return excel_processor.process_audit_report(request)


@app.post("/api/create_photos_check_list")
@login_required
@check_is_confirmed
def create_photos_check_list():
   return excel_processor.create_photos_check_list(request)


#@app.post("/api/create_project_scope_pdf")
#def create_project_scope_pdf():
#    return excel_processor.create_project_scope_pdf(request)


#@app.post("/api/create_scope_pdf")
#def create_scope_pdf():
#    return excel_processor.create_scope_pdf(request)


@app.post("/api/create_scope_pdf_v2")
@login_required
@check_is_confirmed
def create_scope_pdf_v2():
    return excel_processor.create_scope_pdf_v2(request)


@app.post("/api/create_purchase_order")
@login_required
@check_is_confirmed
def create_purchase_order():
    return excel_processor.create_purchase_order(request)


@app.post("/api/create_clip")
@login_required
@check_is_confirmed
def create_clip():
    return excel_processor.create_clip(request)


@app.post("/api/create_materials")
@login_required
@check_is_confirmed
def create_materials():
    return excel_processor.create_materials(request)


@app.post("/api/create_materials_combined")
@login_required
@check_is_confirmed
def create_materials_combined():
    return excel_processor.create_materials(request, combined=True)


@app.post("/api/create_customer_invoice")
@login_required
@check_is_confirmed
def create_customer_invoice():
    return excel_processor.create_customer_invoice(request)


@app.post("/api/image_compression")
@login_required
@check_is_confirmed
def image_compression():
    return excel_processor.image_compression(request)


@app.route("/")
@login_required
@check_is_confirmed
def index():
    db_config = database.get_config()
    PROJECTS_FOLDER = db_config[Database.DEFAULT_PROJECTS_PATH].replace("\\", "\\\\")
    return render_template("main.html", PATH_COLUMN=Database.PROJECT_PATH_COL, PROJECTS_FOLDER=PROJECTS_FOLDER)


@app.route("/create")
@login_required
@check_is_confirmed
def create():
    identifier = request.args.get('identifier')
    pdf_path = redact(
        db_path=app.config["DATABASE_PATH"], sheet_name=Database.PROJECTS_SHEET, header_row=Database.PROJECT_HEADER_ROW,
        template_path=PDF_TEMPLATE, identifier=identifier
    )
    return send_file(pdf_path)


