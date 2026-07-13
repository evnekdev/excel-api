@echo.
@echo -----------------------------
@echo      %2 %1
@echo -----------------------------
@pushd .\%1
@if not exist makefile goto dne
nmake -f makefile %2 -nologo
@if %errorlevel% NEQ 0 goto fail
@popd
@goto :EOF

:fail
@popd
@echo.
@echo !!!---------------------------------!!!
@echo !!!           %2 FAILED
@echo !!!---------------------------------!!!
@echo.
@goto :EOF

:dne
@popd
@echo No makefile found; aborting
@echo.