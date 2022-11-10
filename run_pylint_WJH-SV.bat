:: "run_pylint_WJH-SV.bat"
:: run pylint over WJH-SV code

@echo off

set TARGET_DIR=Source
set IGNORE_LIST=
set CONFIG_FILE=%TARGET_DIR%\.pylintrc
set LOG_FILE=pylint_WJH-SV.log

::pylint --generate-rcfile > %CONFIG_FILE%
pylint --rcfile=%CONFIG_FILE% %TARGET_DIR% --reports=y --output=%TARGET_DIR%\%LOG_FILE% --ignore=%IGNORE_LIST%
echo 'pylint exit code: %ERRORLEVEL%'>>%TARGET_DIR%\%LOG_FILE%
cd %TARGET_DIR%
notepad %LOG_FILE%