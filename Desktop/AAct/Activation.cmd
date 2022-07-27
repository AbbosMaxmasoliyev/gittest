rem Порядок работы активаторов (ActOrder): 1 - в Windows 7 сначала AAct, потом KMSAutoNet. В Windows 8 и 10 сначала KMSAutoNet, а потом AAct.
rem Порядок работы активаторов (ActOrder): 2 - в Windows 7, 8 и 10 сначала AAct, потом KMSAutoNet.
set ActOrder=2

cd /d "%~dp0"

set step=1
set ActivatorExist=-1
set OfficeExist=-1
set osppexist=-1
set osppPFexist=-1
set osppPF86exist=-1
set WindowsActivated=-1
set OfficeActivated=-1
set CancelWinAct=-1
set win=-1
set win2=-1

if exist "%ProgramFiles%\Microsoft Office" set OfficeExist=1
if exist "%ProgramFiles(x86)%\Microsoft Office" set OfficeExist=1
if exist "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS" (
	set osppexist=1
	set "ospppath=%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS"
)
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.VBS" (
	set osppexist=1
	set "ospppath=%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.VBS"
)

if exist "WindowsActivationFailed" del /f /a /q "WindowsActivationFailed"
if exist "OfficeActivationFailed" del /f /a /q "OfficeActivationFailed"
if exist "AAct\AAct_x64.exe" if /i not "%Processor_Architecture%" equ "x86" (move /y "AAct\AAct_x64.exe" "AAct\AAct.exe")

reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName | findstr /i 7
if not errorlevel 1 (set win=7) else (set win=810)
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName | findstr /i 10
if not errorlevel 1 (set win2=10)

:ActivationReStart

if %step% neq 1 if %step% neq 4 if not exist "Activation.Done" (
	schtasks.exe /delete /tn "AAct" /f
	schtasks.exe /delete /tn "KMSAutoNet" /f
)

if %step% equ 4 (
	if exist "%SytemDrive%\ProgramData\KMSAutoS" set ActivatorExist=1
	if exist "%SytemDrive%\ProgramData\KMSAuto" set ActivatorExist=1
	if exist "%WinDir%\AAct_Tools" set ActivatorExist=1
)

if %step% equ 4 (
	if %WindowsActivated% equ 0 @echo WindowsActivationFailed>"WindowsActivationFailed"
	if %OfficeActivated% equ 0 @echo OfficeActivationFailed>"OfficeActivationFailed"
	if %ActivatorExist% equ 1 if not exist "..\IsRetail" if not exist "MSDM Key\KEYOK" copy /y "%Activator%\*.lnk" "%AllUsersProfile%\Microsoft\Windows\Start Menu\Programs\"
	if %ActivatorExist% equ 1 if %OfficeExist% equ 1 copy /y "%Activator%\*.lnk" "%AllUsersProfile%\Microsoft\Windows\Start Menu\Programs\"
	if not exist "Activation.Done" @echo Activation.Done>"Activation.Done"
	exit
)

if exist "AAct\AAct.exe" (
	set Activator=AAct
	set WinAct_param=/win=act /taskwin
	set OffAct_param=/ofs=act /taskofs
	set WinGVLK_param=/wingvlk
	set OfficeGVLK_param=/ofsgvlk
	set Additional_params=/auto
	if ActOrder equ 1 (
		if %win% equ 7 if %step% equ 2 set Additional_params=
		if %win% equ 810 if %step% equ 3 set Additional_params=
		) else (
			if %step% equ 2 set Additional_params=
		)
) else (
	if ActOrder equ 1 (
		if %win% equ 7 if %step% equ 1 set "step=3" & goto :ActivationReStart
		if %win% equ 810 if %step% equ 2 set "step=4" & goto :ActivationReStart
	) else (
		if %step% equ 1 set "step=3" & goto :ActivationReStart
	))

if ActOrder equ 1 (	
	if %win% equ 7 if not %step% equ 3 goto :SkipKMSAutoNet
	if %win% equ 810 if %step% equ 2 goto :SkipKMSAutoNet
) else (	
	if not %step% equ 3 goto :SkipKMSAutoNet
)

if exist "KMSAuto Net\KMSAuto Net.exe" (
	set Activator=KMSAuto Net
	set WinAct_param=/win=act /key=yes
	set OffAct_param=/off=act
	set WinGVLK_param=
	set OfficeGVLK_param=
	set Additional_params=/tap /task=yes /sound=no
) else (
	if ActOrder equ 1 (	
		if %win% equ 7 if %step% equ 3 set "step=4" & goto :ActivationReStart
		if %win% equ 810 if %step% equ 1 set "step=2" & goto :ActivationReStart
	) else (
		set "step=4" & goto :ActivationReStart
	))

:SkipKMSAutoNet

if %osppexist% equ 1 (
	cscript "%ospppath%" /dstatus | findstr "LICENSED"
	if not errorlevel 1 ( set OfficeActivated=1 ) else ( set OfficeActivated=0 )
)

if exist "MSDM Key\KEYOK" if %OfficeActivated% neq 1 goto :KMSActivationOfficeIfExist

cscript %windir%\system32\slmgr.vbs -dli | findstr "Licensed лицензию"
if not errorlevel 1 (
	if %OfficeActivated% neq 1 goto :KMSActivationOfficeIfExist
	if %OfficeActivated% equ 1 (
		@echo Activation.Done>"Activation.Done"
		set step=4
		goto :ActivationReStart
	)
)

cscript %windir%\system32\slmgr.vbs -dli | findstr "VOLUME_KMSCLIENT"
if not errorlevel 1 (
	if %OfficeExist% equ 1 if %OfficeActivated% neq 1 goto :KMSActivationAll
	goto :KMSActivationWindows
)

if exist "..\IsRetail" if %OfficeActivated% neq 1 goto :KMSActivationOfficeIfExist

if %OfficeExist% equ 1 if %OfficeActivated% neq 1 goto :KMSActivationAll
if %OfficeActivated% equ 1 goto :CheckActivation
goto :KMSActivationWindows

:KMSActivationOfficeIfExist
if %OfficeExist% equ 1 goto :KMSActivationOffice
goto :CheckActivation

:KMSActivationOffice
start "" /wait /d "%Activator%" "%Activator%\%Activator%.exe" %Additional_params% %OffAct_param% %OfficeGVLK_param%
goto :CheckActivation

:KMSActivationAll
if exist W10DigitalActivation.exe if %win2% equ 10 if %step% equ 1 if %WindowsActivated% neq 1 (
		start "" /wait W10DigitalActivation.exe /activate /kms38
		start "" /wait /d "%Activator%" "%Activator%\%Activator%.exe" %Additional_params% %OffAct_param% %OfficeGVLK_param%
		goto :CheckActivation
)

REG QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName | findstr /C:"Windows 10 Enterprise"
if not errorlevel 1 (
	start "" /wait /d "%Activator%" "%Activator%\%Activator%.exe" %Additional_params% %WinAct_param% %OffAct_param% %OfficeGVLK_param%
	goto :CheckActivation
)

start "" /wait /d "%Activator%" "%Activator%\%Activator%.exe" %Additional_params% %WinAct_param% %WinGVLK_param% %OffAct_param% %OfficeGVLK_param%
goto :CheckActivation

:KMSActivationWindows
if exist W10DigitalActivation.exe if %win2% equ 10 if %step% equ 1 if %WindowsActivated% neq 1 (
		start "" /wait W10DigitalActivation.exe /activate /kms38
		goto :CheckActivation
)

REG QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName | findstr /C:"Windows 10 Enterprise"
if not errorlevel 1 (
	start "" /wait /d "%Activator%" "%Activator%\%Activator%.exe" %Additional_params% %WinAct_param%
	goto :CheckActivation
)

start "" /wait /d "%Activator%" "%Activator%\%Activator%.exe" %Additional_params% %WinAct_param% %WinGVLK_param%
goto :CheckActivation

:CheckActivation
if %OfficeActivated% neq 1 (
	if %osppexist% equ 1 (
		cscript "%ospppath%" /dstatus | findstr "LICENSED"
		if errorlevel 1 (
			set /a step=%step%+1
			set OfficeActivated=0
			goto :ActivationReStart	
		) else ( set OfficeActivated=1 )
	)
)
cscript %windir%\system32\slmgr.vbs -dli | findstr "Licensed лицензию"
if errorlevel 1 (
	if not exist "MSDM Key\KEYOK" (
		set /a step=%step%+1
		if exist "..\IsRetail" (
			@echo Activation.Done>"Activation.Done"
			set step=4
		)
	set WindowsActivated=0
	goto :ActivationReStart
)	)
@echo Activation.Done>"Activation.Done"
set step=4
goto :ActivationReStart