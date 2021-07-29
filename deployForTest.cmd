@echo off
Set /P answr=deploy (r)elease (empty for debug)? 
set source=bin\Debug
If "%answr%"=="r" (
	set source=bin\Release
	copy /Y %source%\ScriptAddin-AddIn64-packed.xll Distribution\ScriptAddin64.xll"
	copy /Y %source%\ScriptAddin-AddIn-packed.xll Distribution\ScriptAddin32.xll"
	copy /Y %source%\ScriptAddin.dll.config Distribution\ScriptAddin.xll.config
	copy /Y ScriptAddinCentral.config Distribution
	copy /Y ScriptAddinUser.config Distribution
)

if exist "C:\Program Files\Microsoft Office\root\" (
	echo 64bit office
	copy /Y %source%\ScriptAddin-AddIn64-packed.xll "%appdata%\Microsoft\AddIns\ScriptAddin.xll"
	copy /Y %source%\ScriptAddin.pdb "%appdata%\Microsoft\AddIns"
	copy /Y %source%\ScriptAddin.dll.config "%appdata%\Microsoft\AddIns\ScriptAddin.xll.config"
	copy /Y ScriptAddinCentral.config "%appdata%\Microsoft\AddIns\ScriptAddinCentral.config"
) else (
	echo 32bit office
	copy /Y %source%\ScriptAddin-AddIn-packed.xll "%appdata%\Microsoft\AddIns\ScriptAddin.xll"
	copy /Y %source%\ScriptAddin.pdb "%appdata%\Microsoft\AddIns"
	copy /Y %source%\ScriptAddin.dll.config "%appdata%\Microsoft\AddIns\ScriptAddin.xll.config"
	copy /Y ScriptAddinCentral.config "%appdata%\Microsoft\AddIns\ScriptAddinCentral.config"
)
pause