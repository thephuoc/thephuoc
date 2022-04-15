@set @_cmd=1 /*
@Echo Off
setlocal EnableExtensions
Color 1b
title Edit Code : Hamano Kaito
whoami /groups | find "S-1-16-12288" >nul && goto :RUNME
if "%~1"=="RunAsAdmin" goto :Loi
echo Quyen Admin se duoc khoi tao de chay fix nay........................
cscript /nologo /e:javascript "%~f0" || goto :Loi
exit /b
:Loi
echo.
echo    Loi: Yeu cau quyen Admin that bai,
echo Vui long chay nhu quyen Admin de tiep tuc.
echo.
goto :KT
:RUNME
pushd "%~dp0"
CD /D "%~dp0"
Set IDMAN64b="C:\Program Files (x86)\Internet Download Manager\IDMan.exe"
Set IDMAN32b="C:\Program Files\Internet Download Manager\IDMan.exe"
Set K64b=HKEY_CURRENT_USER\Software\Classes\WOW6432Node\CLSID
Set K32b=HKEY_CURRENT_USER\Software\Classes\CLSID
if exist "%SystemRoot%\SysWOW64\" (set TM=%IDMAN64b%) else (set TM=%IDMAN32b%)
if exist "%SystemRoot%\SysWOW64\" (set ROOT=%K64b%) else (set ROOT=%K32b%)
Cls
Taskkill /f /im IDMan.exe >nul
Taskkill /f /im IDMIntegrator64.exe >nul
Taskkill /f /im IDMIntegrator.exe >nul
Taskkill /f /im IEMonitor.exe >nul
Cls
Echo Truoc tien se kiem tra backup va tao backup
Echo.
Echo Lenh se tu tao backup cho cac ban, neu co loi
Echo thi cac ban chay file BACKUPKEY.REG trong
Echo phan vung C.
Echo.
Set /p= Nhan phim bat ky de tiep tuc....
IF EXIST "c:\BACKUPKEY.REG" (Goto 1) ELSE (Goto 2)
:2
Cls
Reg Export "%ROOT%" c:\BACKUPKEY.REG >nul
Reg Delete "%ROOT%" /f >nul
Echo Hoan tat Backup va Xoa Key
Timeout 3 >nul
Goto 1
:1
Cls
echo Su dung "DK" neu ban muon day du tinh nang.
Echo.
Echo Su dung "DKL" neu ban muon chi de khoa reg.
Echo.
echo Nhan "DK" de dang ky Full, "DKL" de chi khoa reg.
echo.
set /p ask=DK hoac DKL? (DK/DKL)
if %ask%==dk goto 1reg
If %ask%==dkl goto 1trial
if not defined ask goto error
Cls
:error
Echo Vui long ghi dung ky tu.
Timeout 5 >nul
Goto 1
