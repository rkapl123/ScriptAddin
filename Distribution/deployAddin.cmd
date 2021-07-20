rem copy Addin and settings...
@echo off
if exist "C:\Program Files\Microsoft Office\root\" (
	echo 64bit office
	copy /Y Raddin64.xll "%appdata%\Microsoft\AddIns\Raddin.xll"
	copy /Y Raddin.xll.config "%appdata%\Microsoft\AddIns\Raddin.xll.config"
) else (
	echo 32bit office
	copy /Y Raddin32.xll "%appdata%\Microsoft\AddIns\Raddin.xll"
	copy /Y Raddin.xll.config "%appdata%\Microsoft\AddIns\Raddin.xll.config"
)
rem start Excel and install Addin there..
cscript //nologo installRAddinInExcel.vbs
pause
