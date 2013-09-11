@echo off

REM Set Variables
SET PATCHDIR=d:\software\patches\{version #}
SET PATCHDIRSPECIAL=x:\software\patches\special

REM Make list of all patches to run
dir %PATCHDIR% /b > d:\sofware\patches\patches.txt

REM Install patches to client machine based on patch list
dir %PATCHDIR% /b > d:\software\patches\patches.txt
For /f "tokens=1" %%A in (d:\software\patches\patches.txt) do (
%PATCHDIR%\%%A /passive /norestart
)



REM cd %PATCHDIRSPECIAL%
REM Add special patching here and remove REM above if need be
