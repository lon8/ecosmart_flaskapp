#PROJECTS_BASE_PATH = r"C:\Users\sandi\EcoSmart Solutions Dropbox\EcoSmart Solutions Team Folder"
#DATABASE_PATH = r"C:\Users\database.xlsx"
#XL_TEMPLATE_PATH = r"C:\Users\sandi\EcoSmart Solutions Dropbox\EcoSmart Solutions Team Folder\Anton\auto_templates"


from decouple import config


DATABASE_URI = config("SQLITE_DATABASE_URL")
if DATABASE_URI.startswith("postgres://"):
    DATABASE_URI = DATABASE_URI.replace("postgres://", "postgresql://", 1)


class Config(object):
    DEBUG = False
    TESTING = False
    CSRF_ENABLED = True
    SECRET_KEY = config("SECRET_KEY", default="guess-me")
    SQLALCHEMY_DATABASE_URI = DATABASE_URI
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    BCRYPT_LOG_ROUNDS = 13
    WTF_CSRF_ENABLED = True
    DEBUG_TB_ENABLED = False
    DEBUG_TB_INTERCEPT_REDIRECTS = False
    SECURITY_PASSWORD_SALT = config("SECURITY_PASSWORD_SALT", default="very-important")

    # Mail Settings
    MAIL_DEFAULT_SENDER = "noreply@flask.com"
    MAIL_SERVER = "smtp.gmail.com"
    MAIL_PORT = 465
    MAIL_USE_TLS = False
    MAIL_USE_SSL = True
    MAIL_DEBUG = False
    MAIL_USERNAME = config("EMAIL_USER")
    MAIL_PASSWORD = config("EMAIL_PASSWORD")

    PROJECTS_BASE_PATH = (
        r"C:\Users\sandi\EcoSmart Solutions Dropbox\EcoSmart Solutions Team Folder"
    )
    DATABASE_PATH = r"C:\Users\sandi\EcoSmart Solutions Dropbox\EcoSmart Solutions Team Folder\Anton\database.xlsx"
    XL_TEMPLATE_PATH = r"C:\Users\sandi\EcoSmart Solutions Dropbox\EcoSmart Solutions Team Folder\Anton\auto_templates"

    LOG_LEVEL = "INFO"
    LOG_BACKTRACE = False

    AUTH_USERS_FILE_PATH = "auth_users.txt"


class DevelopmentConfig(Config):
    DEVELOPMENT = True
    DEBUG = True
    WTF_CSRF_ENABLED = False
    DEBUG_TB_ENABLED = True


class TestingConfig(Config):
    TESTING = True
    DEBUG = True
    SQLALCHEMY_DATABASE_URI = "sqlite:///testdb.sqlite"
    BCRYPT_LOG_ROUNDS = 1
    WTF_CSRF_ENABLED = False


class ProductionConfig(Config):
    DEBUG = False
    DEBUG_TB_ENABLED = False
