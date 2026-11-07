@echo off
python "%~dp0run_acceptance.py" %*
exit /b %errorlevel%
