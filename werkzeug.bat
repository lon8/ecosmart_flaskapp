set path=%~dp0\venv\Scripts\;%path%
call %~dp0\venv\Scripts\activate
set FLASK_APP=proj.py
set FLASK_DEBUG=1
flask run --host=0.0.0.0