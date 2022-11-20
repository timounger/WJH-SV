:: "generate_WJH-SV_executable.bat"
:: generate WJH-SV.exe

@echo off

IF "%~1"=="" (
	echo No explicit python version given. Using folowing:
	python --version
	set path_python_exe=python
) ELSE (
	echo Using python %1
	set path_python_exe=%1
)

set workpath=build

set PYTHONHASHSEED=1

%path_python_exe% .\generate_version_file.py %workpath%
IF %ERRORLEVEL% NEQ 0 (
	exit /b %ERRORLEVEL%
)

%path_python_exe% -m PyInstaller --clean ^
								 --paths ..\..\ ^
								 --icon ..\Resources\wjh_sv.png ^
								 --version-file %workpath%\wjh_sv_version_info.txt ^
								 --hidden-import html.parser ^
								 --exclude-module lxml ^
								 --exclude-module numpy ^
								 --exclude-module typing_extensions ^
								 --exclude-module pygments ^
								 --name WJH-SV ^
								 --onefile ^
								 --noupx ^
								 --distpath bin ^
								 --workpath %workpath% ^
								 ..\Source\wjh_sv.py
IF %ERRORLEVEL% NEQ 0 (
	exit /b %ERRORLEVEL%
)

%path_python_exe% .\check_included_packages.py

:: Cleanup
del WJH-SV.spec
::rmdir /s/q "build"
set PYTHONHASHSEED=
pause