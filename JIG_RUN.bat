@echo off

cd /d D:\App_Jig_v0
call D:\App_Jig_v0\myenv\Scripts\Activate
python App_Jig.py
pause

@REM powershell -NoExit -Command "& {cd 'D:\App_Jig_v0'; ./myenv/Scripts/Activate; python App_Jig.py}"