@echo off
rem
rem Double click on this file to start the Strategic Pricing Database
rem
xcopy "\\BRICKLIN\COMMON\Mtce\weldproc\Weld Proc Database\Weld Proc DB-Archive.accde"  H:\\temp_db\ /Q /Y
start "" "H:\\temp_db\Weld Proc DB-Archive.accde"
exit