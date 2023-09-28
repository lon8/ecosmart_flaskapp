import pathlib
from functools import wraps

from decouple import config
from flask import request, current_app, render_template
from flask_login import current_user
from flask_login.config import EXEMPT_METHODS


def authorize_user(email):
    if not email:
        return None
    items = pathlib.Path(config("AUTH_USERS_FILE_PATH")).open("r").readlines()
    items = list(map(lambda x: x.replace("\n", ""), items))
    if email not in items:
        return False
    return True


def login_required(func):
    """
    If you decorate a view with this, it will ensure that the current user is
    logged in and authenticated before calling the actual view. (If they are
    not, it calls the :attr:`LoginManager.unauthorized` callback.) For
    example::

        @app.route('/post')
        @login_required
        def post():
            pass

    If there are only certain times you need to require that your user is
    logged in, you can do so with::

        if not current_user.is_authenticated:
            return current_app.login_manager.unauthorized()

    ...which is essentially the code that this function adds to your views.

    It can be convenient to globally turn off authentication when unit testing.
    To enable this, if the application configuration variable `LOGIN_DISABLED`
    is set to `True`, this decorator will be ignored.

    .. Note ::

        Per `W3 guidelines for CORS preflight requests
        <http://www.w3.org/TR/cors/#cross-origin-request-with-preflight-0>`_,
        HTTP ``OPTIONS`` requests are exempt from login checks.

    :param func: The view function to decorate.
    :type func: function
    """

    @wraps(func)
    def decorated_view(*args, **kwargs):
        if request.method in EXEMPT_METHODS or current_app.config.get("LOGIN_DISABLED"):
            pass
        elif not current_user.is_authenticated:
            return current_app.login_manager.unauthorized()  # noqa
        elif current_user.is_authenticated:
            if not authorize_user(current_user.email):
                return render_template("errors/403.html")

        # flask 1.x compatibility
        # current_app.ensure_sync is only available in Flask >= 2.0
        if callable(getattr(current_app, "ensure_sync", None)):
            return current_app.ensure_sync(func)(*args, **kwargs)
        return func(*args, **kwargs)

    return decorated_view
