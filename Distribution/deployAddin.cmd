rem copy Addin and settings...
@echo off
if exist "C:\Program Files\Microsoft Office\root\" (
	echo 64bit office
	copy /Y ScriptAddin64.xll "%appdata%\Microsoft\AddIns\ScriptAddin.xll"
	copy /Y ScriptAddin.xll.config "%appdata%\Microsoft\AddIns\ScriptAddin.xll.config"
) else (
	echo 32bit office
	copy /Y ScriptAddin32.xll "%appdata%\Microsoft\AddIns\ScriptAddin.xll"
	copy /Y ScriptAddin.xll.config "%appdata%\Microsoft\AddIns\ScriptAddin.xll.config"
)
rem start Excel and install Addin there..
cscript //nologo installAddinInExcel.vbs
pause
