@echo off
@setlocal enableextensions enabledelayedexpansion


del "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\*.*" /q


echo.
echo.
REM Accroding to NECPC-JS coordinator, do not connect to internet to prevent update from 2020/06/01.
goto StopWU

timeout /t 3 > nul
echo.
echo  Please connect to Internet and press any key to start.
pause > nul

echo.
echo.
echo ***********************************************************
echo  [1] Windows Update                            
echo ***********************************************************
echo.
echo  Windows Update will start after 10 seconds. Please wait for a while.
	timeout /t 10 > nul

echo.
echo  Updating...
	powershell explorer ms-settings:windowsupdate-action
	timeout /t 3 > nul

echo.
echo ***********************************************************
echo  [1] Store Update                            
echo ***********************************************************
echo.
echo.
echo  Updating...
	powershell explorer ms-windows-store:updates
	timeout /t 5 > nul
	c:\Jumpstart\batch\nircmd.exe sendkeypress Enter

echo.
echo  Idle 20 minutes for Windows Update and Store update...
echo.
echo  Or you can press enter to skip it.
	timeout /t 1200

echo.
echo.
echo ***********************************************************
echo  [1] Restart Windows Update and Store Update                            
echo ***********************************************************
echo.
echo  Updating...
echo.
echo.
	powershell explorer ms-settings:windowsupdate-action
	timeout /t 2 > nul
	
	powershell explorer ms-windows-store:updates
	timeout /t 5 > nul
	c:\Jumpstart\batch\nircmd.exe sendkeypress Enter
	
	echo  1st time UAC setting Reboot needed
        goto ContinueWU
	
:ContinueWU
cd C:\Jumpstart\batch > nul
type C:\Jumpstart\batch\RBC.txt > C:\Jumpstart\batch\JSP.bat
	xcopy C:\Jumpstart\batch\JSP.bat "C:\Users\%username%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\"

echo  System will reboot after 5 seconds.
	timeout /t 5 > nul
	shutdown /r /t 0
pause > nul


:StopWU
echo temp > C:\Jumpstart\batch\UPDone.log
::echo   Please disconnect to Internet, then press any key to continue.
::pause > nul

echo.
echo ***********************************************************
echo  [1] Disable Windows Update                            
echo ***********************************************************
echo. 
net stop wuauserv
sc config wuauserv start= disabled
echo.
echo.

echo.
echo ***********************************************************
echo  [2] Disable Windows Store Update                            
echo ***********************************************************
echo.
sc config wuauserv start= disabled
echo.
echo.

echo.
echo ***********************************************************
echo  [3] Change Power Option Setting                            
echo ***********************************************************	
echo.
powercfg -import c:\Jumpstart\batch\Jumpstart_test.pow 6e7b5744-f24d-4c44-a6c1-89de963d65f9
Powercfg -setactive 6e7b5744-f24d-4c44-a6c1-89de963d65f9
echo.
echo.

echo.
echo.
echo ***********************************************************
echo  [4] Disable InfoBoard application settings
echo ***********************************************************	
echo.

PowerShell.exe -ExecutionPolicy Bypass -File "C:\Jumpstart\batch\infoboardoff.ps1"

echo.  Idle 10 seconds.
timeout /t 10 > nul
echo.
echo.

echo.
echo ***********************************************************
echo  [5] Start Testing and Run Auto
echo ***********************************************************	
echo.



if exist "%USERPROFILE%\Desktop\*.html"  (
      
       del "C:\Users\%username%\desktop\get_report*.*" /q
       del "C:\Users\%username%\desktop\JS_OEM*.*" /q
   
     PowerShell.exe -ExecutionPolicy Bypass -File "C:\Jumpstart\batch\JS_auto2_baseline.ps1"

) else (
   
     PowerShell.exe -ExecutionPolicy Bypass -File "C:\Jumpstart\batch\JS_auto2.ps1"

)

net user %username% 1111
echo.
powershell Set-ExecutionPolicy -ExecutionPolicy bypass
echo temp > C:\Jumpstart\batch\PreDone.txt
echo S | powershell -NoProfile -command C:\Jumpstart\performance\performance_assessments.cmd
echo.  Idle 10 seconds.
timeout /t 10 > nul

exit



:Fail
    echo Failed  
    echo return value = %ERRORLEVEL%  
	echo "Wait 5 sec go back meun"
	timeout /t 5
	pause > nul
	echo  Press any key to return to menu.
	goto Home


:FailUWP
	explorer "%userprofile%\AppData\Local\Temp\DCVT.Working\Collateral
	echo.
	echo  UWP test failed!
	echo.
	echo  Please check fail log!
	timeout /t 2 > nul
	exit