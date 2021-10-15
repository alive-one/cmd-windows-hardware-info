rem | If file with the same name already exists, skip writing CSS style and jump straight to config writing
IF EXIST %~dp0%computername%.html GOTO L1

rem | Write CSS style
@ECHO ^<head^>^<style^> >> %~dp0%computername%.html
@ECHO .div-main {width: 100%; height: 800px; padding: 10px; margin: 10px;} >> %~dp0%computername%.html
@ECHO .div-table {display: inline-block; width: auto; height: 780px; border-left: 1px solid black; padding: 10px; float: left; margin-top: 10px; margin-left: 10px; margin-bottom: 10px;} >> %~dp0%computername%.html
@ECHO .div-table-sec {display: inline-block; width: auto; height: 780px; border-left: 1px solid black; padding: 10px; margin-top: 10px; margin-bottom: 10px;} >> %~dp0%computername%.html
@ECHO .div-table-row {display: table-row;} >> %~dp0%computername%.html
@ECHO .div-table-cell-date {width: auto; padding: 10px 40px 10px; border-top: 1px solid black; border-left: 1px solid black; float: left; color: red;} >> %~dp0%computername%.html
@ECHO .div-table-cell-pcname {width: auto; padding: 10px 40px 10px; border-top: 1px solid black; border-left: 1px solid black; float: left;} >> %~dp0%computername%.html
@ECHO .div-table-cell-mb {width: 180px; padding: 10px 20px 10px; border-top: 1px solid black; border-left: 1px solid black; float: left;} >> %~dp0%computername%.html
@ECHO .div-table-cell-cpu {width: auto; padding: 10px 20px 10px; border-top: 1px solid black; border-left: 1px solid black; float: left;} >> %~dp0%computername%.html
@ECHO .div-table-cell-ram {width: 180px; padding: 10px 20px 10px; border-top: 1px solid black; border-left: 1px solid black; float: left;} >> %~dp0%computername%.html
@ECHO .div-table-cell-stor {width: 180px; padding: 10px 20px 10px; border-top: 1px solid black; border-left: 1px solid black; float: left;} >> %~dp0%computername%.html
@ECHO .div-table-cell-gpu {width: 180px; padding: 10px 20px 10px; border-top: 1px solid black; border-left: 1px solid black; float: left;} >> %~dp0%computername%.html
@ECHO .div-table-cell-net {width: 180px; padding: 10px 20px 10px; border-top: 1px solid black; border-left: 1px solid black; float: left;} >> %~dp0%computername%.html
@ECHO .div-table-cell-os {width: 180px; padding: 10px 20px 10px; border-top: 1px solid black; border-left: 1px solid black; float: left;} >> %~dp0%computername%.html
@ECHO .div-table-cell-lan {width: 180px; padding: 10px 20px 10px; border-top: 1px solid black; border-left: 1px solid black; float: left;} >> %~dp0%computername%.html
@ECHO ^</head^>^</style^> >> %~dp0%computername%.html

rem | Jump here if file with the same name already exists
:L1

rem | Main positioning div
@ECHO ^<div class=^"div-main^"^> >> %~dp0%computername%.html

rem | Open div-table
@ECHO ^<div class=^"div-table^"^> >> %~dp0%computername%.html

rem | Write date and time
@ECHO ^<div class=^"div-table-row^"^>Local Date^</div^> >> %~dp0%computername%.html
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-date^"^>%DATE% %TIME:~0,-3%^</div^>^</div^> >> %~dp0%computername%.html

rem | Write system name
@ECHO ^<div class=^"div-table-row^"^>Computer Name^</div^> >> %~dp0%computername%.html
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-pcname^"^>%COMPUTERNAME%^</div^>^</div^> >> %~dp0%computername%.html

rem | Write header for baseboard info
@ECHO ^<div class=^"div-table-row^"^>Motherboard^</div^> >> %~dp0%computername%.html

rem | Write baseboard manufacturer
@FOR /F "skip=2 delims=, tokens=2" %%i IN ('wmic baseboard get Manufacturer /format:csv') DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-mb^"^>%%i^</div^> >> %~dp0%computername%.html
) 

rem | Write baseboard model
@FOR /F "skip=2 delims=, tokens=2" %%i IN ('wmic baseboard get Product /format:csv') DO (
@ECHO ^<div class=^"div-table-cell-mb^"^>%%i^</div^>^</div^> >> %~dp0%computername%.html
)

rem | Write header for CPU(s) info
@ECHO ^<div class=^"div-table-row^"^>CPU^</div^> >> %~dp0%computername%.html

rem | Write info on CPU(s)
FOR /F "skip=2 delims=, tokens=1,*" %%i IN ('wmic cpu get name /format:csv') DO (
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-cpu^"^>%%j^</div^>^</div^> >> %~dp0%computername%.html
)

rem | Write info on operating memory
@ECHO ^<div class=^"div-table-row^"^>RAM^</div^> >> %~dp0%computername%.html

rem | Allow changing variables in cycle
Setlocal EnableDelayedExpansion

rem | Write info on RAM slot, size and speed 
@FOR /F "tokens=1,2,3" %%a IN ('wmic memorychip get capacity^,devicelocator^,speed ^| findstr [0-9]') DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-ram^"^>%%b^</div^> >> %~dp0%computername%.html

rem | Use powershell in inner cycle to divide each token of cycle by 1048576 to get size in Mb
@FOR /F %%a IN ('powershell %%a/1048576') DO (
SET /A mem_fnl=%%a
@ECHO ^<div class=^"div-table-cell-ram^"^>!mem_fnl! Mb^</div^> >> %~dp0%computername%.html
@ECHO ^<div class=^"div-table-cell-ram^"^>%%c Mhz^</div^>^</div^> >> %~dp0%computername%.html
)
)

rem | Write info on storage devices
@ECHO ^<div class=^"div-table-row^"^>Storage Devices^</div^> >> %~dp0%computername%.html

rem | Get size model and status for every fixed hard disk media storage device
rem | Use powershell to divide size by 1000000000 to get Gb
@FOR /F "skip=2 delims=, tokens=2-4" %%i IN ('wmic diskdrive where ^(MediaType^="Fixed hard disk media"^) get model^,size^,status /format:csv') DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-stor^"^>%%i^</div^> >> %~dp0%computername%.html
@FOR /F %%j IN ('powershell %%j/1000000000') DO (
SET /A stor_fnl=%%j
@ECHO ^<div class=^"div-table-cell-stor^"^>!stor_fnl! Gb^</div^> >> %~dp0%computername%.html
@ECHO ^<div class=^"div-table-cell-stor^"^>%%k^</div^>^</div^> >> %~dp0%computername%.html
))

rem | Write info on Video Controllers
@ECHO ^<div class=^"div-table-row^"^>Video Adapters^</div^> >> %~dp0%computername%.html

rem | Get VideoControllers name and DeviceID
FOR /F "skip=2 tokens=2,3 delims=," %%a IN ('wmic path win32_VideoController get name^, PNPDeviceID /format:csv') DO (
SET vc_name=%%a

rem | Filter out Device ID to get rid of & simbols 
FOR /F "tokens=2 delims=^&amp;" %%c IN ("%%b") DO (

rem | Search registry branch where stored info on Video Adapters for certain registry key where Device ID of each VideoAdapter stored
FOR /F %%d IN ('REG QUERY HKLM\SYSTEM\CurrentControlSet\Control\Class\{4D36E968-E325-11CE-BFC1-08002BE10318} /f %%c /s /t REG_SZ ^| find "{"') DO (

rem | Search memory size entry in the same registry key where was found Device ID
rem | Use powershell to convert hex value to decimal and divide by 1048576 to get size in Mb
FOR /F "tokens=3" %%e IN ('REG QUERY ^"%%d^" /f HardwareInformation.qwMemorySize /t REG_QWORD ^| find "0x"') DO (
SET vc_mem=powershell [uint64]^('%%e'^)/1048576
)
)
)
rem | Write VideoAdapter name in file
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-gpu^"^>!vc_name!^</div^> >> %~dp0%computername%.html

rem | Write amount of memory in file
rem | If variable vc_mem less than 1 echo DVMT
ECHO ^<div class=^"div-table-cell-gpu^"^> >> %~dp0%computername%.html
IF !vc_mem! LSS 1 (SET vc_mem=N/A or DVMT&ECHO !vc_mem! >> %~dp0%computername%.html) ELSE (!vc_mem! >> %~dp0%computername%.html&ECHO Mb >> %~dp0%computername%.html)
ECHO ^</div^>^</div^> >> %~dp0%computername%.html
)

rem | Disable variable changing in cycle
Setlocal DisableDelayedExpansion

rem | Write info on network adapters
@ECHO ^<div class=^"div-table-row^"^>Network Adapters^</div^> >> %~dp0%computername%.html

rem | Write info on network adapters MAC-address and model
@FOR /F "tokens=1*" %%p IN ('wmic NIC where PhysicalAdapter^=true get macaddress^,name ^| findstr [0-9]') DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-net^"^>%%q^</div^> >> %~dp0%computername%.html
@ECHO ^<div class=^"div-table-cell-net^"^>%%p^</div^>^</div^> >> %~dp0%computername%.html
)

rem | Close table with major info on hradware
@ECHO ^</div^> >> %~dp0%computername%.html

rem | Open second table with info on Windows and network connections
@ECHO ^<div class=^"div-table-sec^"^> >> %~dp0%computername%.html

rem | Write info on Operating System
@ECHO ^<div class=^"div-table-row^"^>Operating System^</div^> >> %~dp0%computername%.html
FOR /F "skip=2 tokens=2-4 delims=," %%i IN ('wmic os get caption^, osarchitecture^, version /format:csv') DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-os^"^>Name^</div^>^<div class=^"div-table-cell-os^"^>%%i^</div^>^</div^> >> %~dp0%computername%.html
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-os^"^>Architecture^</div^>^<div class=^"div-table-cell-os^"^>%%j^</div^>^</div^> >> %~dp0%computername%.html
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-os^"^>Version^</div^>^<div class=^"div-table-cell-os^"^>%%k^</div^>^</div^> >> %~dp0%computername%.html
)

rem | Write Activation Status
FOR /F "tokens=* skip=1" %%i IN ('cscript.exe /nologo c:\windows\system32\slmgr.vbs /xpr') DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-os^"^>Activation^</div^>^<div class=^"div-table-cell-os^"^>%%i^</div^>^</div^> >> %~dp0%computername%.html
)

rem | Get mac address and connection name 
@FOR /F "skip=2 delims=, tokens=2,3" %%i IN ('wmic nic where PhysicalAdapter^=true get MACAddress^, NetConnectionID /format:csv') DO (

rem | Use connection name as header for detailed info on each connection
@ECHO  ^<div class=^"div-table-row^"^>%%j^</div^> >> %~dp0%computername%.html

rem | Write MAC-address
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-lan^"^>MAC Address^</div^>^<div class=^"div-table-cell-lan^"^>%%i^</div^>^</div^> >> %~dp0%computername%.html

rem | Get Default Gateway, IP address and Subnet Mask
@FOR /F "skip=2 delims=,{} tokens=2,3,4" %%a IN ('wmic nicconfig where ^(ipenabled^="true" AND macaddress^="%%i"^) get DefaultIPGateway^, IPAddress^, IPSubnet /format:csv') DO (

rem | Get rid of IPv6 address
@FOR /F "delims=; tokens=1" %%z IN (^"%%b^") DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-lan^"^>IP Address^</div^>^<div class=^"div-table-cell-lan^"^>%%z^</div^>^</div^> >> %~dp0%computername%.html
)

rem | Get rid of unnecessary info in Subnet Mask
@FOR /F "delims=; tokens=1" %%y IN (^"%%c^") DO (
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-lan^"^>Subnet^</div^> ^<div class=^"div-table-cell-lan^"^>%%y^</div^>^</div^> >> %~dp0%computername%.html
)
@ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-lan^"^>Gateway^</div^> ^<div class=^"div-table-cell-lan^"^>%%a^</div^>^</div^> >> %~dp0%computername%.html
)
)

rem | Close second table
@ECHO ^</div^> >> %~dp0%computername%.html

rem | Close main positioning div
@ECHO ^</div^> >> %~dp0%computername%.html

pause
