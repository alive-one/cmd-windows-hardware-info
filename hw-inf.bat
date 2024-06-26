@rem | Enable to change variables in cycle
Setlocal EnableDelayedExpansion

@rem | Write System Name to variable
SET pc_name=%COMPUTERNAME%

@rem | If file with the same name already exists, skip writing CSS style and jump to config writing
IF EXIST %~dp0%pc_name%.html GOTO L1
  
@rem | Write CSS style
@ECHO ^<head^>^<style^> >> %~dp0%pc_name%.html
@ECHO .div-main {width: 100^%%; height: 800px; padding: 10px; margin: 10px;} >> %~dp0%pc_name%.html
@ECHO .div-table {display: inline-block; width: auto; height: 780px; border-left: 1px solid black; padding: 10px; float: left; margin-top: 10px; margin-left: 10px; margin-bottom: 10px;} >> %~dp0%pc_name%.html
@ECHO .div-table-sec {display: inline-block; width: auto; height: 780px; border-left: 1px solid black; padding: 10px; margin-top: 10px; margin-bottom: 10px;} >> %~dp0%pc_name%.html
@ECHO .div-table-row {display: table-row;} >> %~dp0%pc_name%.html
@ECHO .div-table-cell-date {width: auto; padding: 10px 40px 10px; border-top: 1px solid black; border-right: 1px solid black; float: right; color: red;} >> %~dp0%pc_name%.html
@ECHO .div-table-cell-pcname {width: auto; padding: 10px 40px 10px; border-top: 1px solid black; border-left: 1px solid black; float: left;} >> %~dp0%pc_name%.html
@ECHO .div-table-cell-long {width: auto; min-width: 180px; max-width: 400px; padding: 10px 20px 10px; border-top: 1px solid black; border-left: 1px solid black; float: left;} >> %~dp0%pc_name%.html
@ECHO .div-table-cell {width: 180px; padding: 10px 20px 10px; border-top: 1px solid black; border-left: 1px solid black; float: left;} >> %~dp0%pc_name%.html
@ECHO ^</head^>^</style^> >> %~dp0%pc_name%.html
  
@rem | Jump here if file with the same name already exists
:L1
  
@rem | Open main div-table
ECHO ^<div class=^"div-main^"^> >> %~dp0%pc_name%.html
  
@rem | Open first div-table for hardware info
ECHO ^<div class=^"div-table^"^> >> %~dp0%pc_name%.html
  
@rem | Write system name and date
ECHO ^<div class=^"div-table-row^"^>System Name^</div^> >> %~dp0%pc_name%.html
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-pcname^"^>%pc_name%^</div^>^<div class=^"div-table-cell-date^"^>%DATE% %TIME:~0,-3%^</div^>^</div^> >> %~dp0%pc_name%.html
  
@rem | Write info on motherboard
ECHO ^<div class=^"div-table-row^"^>Motherboard^</div^> >> %~dp0%pc_name%.html
  
@rem | Skip two first lines for it is headers when use CSV output
@rem | Use comma as delimiter and get second token from command output which is manufacturer of motherboard
@rem | Do NOT use one command to obtain both manufacturer and model because although it is possible 
@rem | some baseboards has comma in name which lead to wrong tokens enumeration and wrong output
@rem | Write manufacturer
@FOR /F "skip=2 delims=, tokens=2" %%i IN ('wmic path Win32_BaseBoard get Manufacturer /format:csv') DO (
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>%%i^</div^> >> %~dp0%pc_name%.html
)
  
@rem | Write motherboard model in the same table with manufacturer
FOR /F "skip=2 delims=, tokens=2" %%i IN ('wmic path Win32_BaseBoard get Product /format:csv') DO (
ECHO ^<div class=^"div-table-cell^"^>%%i^</div^> >> %~dp0%pc_name%.html
)

@rem | Write BIOS/UEFI version in file
FOR /F "skip=2 tokens=2 delims=," %%i IN ('wmic path Win32_Bios get SMBIOSBIOSVersion /format:csv') DO (
ECHO ^<div class=^"div-table-cell^"^>BIOS ver. %%i^</div^>^</div^> >> %~dp0%pc_name%.html 
) 

@rem | Write info on CPUs
@rem | Write header
ECHO ^<div class=^"div-table-row^"^>CPU^</div^> >> %~dp0%pc_name%.html
  
rem | Get info on CPUs in CSV format and use commas as delimiters
rem | write in file all string except first token which is "Node"
FOR /F "skip=2 delims=, tokens=2-6" %%i IN ('wmic path Win32_Processor get DeviceID^, Name^, NumberOfCores^, NumberOfLogicalProcessors^, SocketDesignation /format:csv') DO (
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>%%i^</div^>^<div class=^"div-table-cell-long^"^>%%j^</div^>^</div^> >> %~dp0%pc_name%.html
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>%%m^</div^>^<div class=^"div-table-cell-long^"^>Cores and Threads: %%k / %%l^</div^>^</div^> >> %~dp0%pc_name%.html
)
  
@rem | Write info on RAM
@rem | Write header
ECHO ^<div class=^"div-table-row^"^>RAM^</div^> >> %~dp0%pc_name%.html
  
@rem | Write info on SLOT, SIZE and SPEED of each memory module
@FOR /F "tokens=1,2,3" %%a IN ('wmic path Win32_PhysicalMemory get capacity^,devicelocator^,speed ^| findstr [0-9]') DO (
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>%%b^</div^> >> %~dp0%pc_name%.html
  
@rem | Use inner FOR loop to convert memory size represented in bytes to megabites
@rem | For that divide token %%a which contains memory size in bytes by 1048576
@rem | Use powershell for CMD can not operate with such numbers
FOR /F %%a IN ('powershell %%a/1048576') DO (
SET /A mem_fnl=%%a
ECHO ^<div class=^"div-table-cell^"^>!mem_fnl! Mb^</div^> >> %~dp0%pc_name%.html
ECHO ^<div class=^"div-table-cell^"^>%%c Mhz^</div^>^</div^> >> %~dp0%pc_name%.html
)
)
  
@rem | Get Info on FIXED Storage Devices (It will not display removable media like USB Stick)
@rem | Write header for info on fixed storage devices
ECHO ^<div class=^"div-table-row^"^>Storage Devices^</div^> >> %~dp0%pc_name%.html
    
@rem | Use CSV output so that neccessary values are divided by commas
@rem | So that even storage deivices names with spaces will be prosessed corectly
@rem | Skip two strings for CSV output use two blank strings as headers
@rem | For all fixed storage devices get model, size, status, serial number and write to div-table
@rem | Use powershell to divide size token by 1000000000 so that we have readable output in Gb
FOR /F "skip=2 delims=, tokens=2-5" %%i IN ('wmic path Win32_DiskDrive where ^(MediaType^="Fixed hard disk media"^) get model^,serialnumber^,size^,status /format:csv') DO (
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>%%i^</div^> >> %~dp0%pc_name%.html
FOR /F %%k IN ('powershell %%k/1000000000') DO (
SET /A stor_fnl=%%k
ECHO ^<div class=^"div-table-cell-long^"^>s/n: %%j^</div^>^</div^> >> %~dp0%pc_name%.html
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>!stor_fnl! Gb^</div^> >> %~dp0%pc_name%.html
ECHO ^<div class=^"div-table-cell^"^>%%l^</div^>^</div^> >> %~dp0%pc_name%.html
)
)
  
@rem | Write info on GPUs
ECHO ^<div class=^"div-table-row^"^>Video Adapters^</div^> >> %~dp0%pc_name%.html

@rem | Get GPU name and DeviceID
FOR /F "skip=2 tokens=2,3 delims=," %%a IN ('wmic path win32_VideoController get name^, PNPDeviceID /format:csv') DO (
SET vc_name=%%a
  
@rem | Get Device ID and filter out wmic output to get rid of special simbols
FOR /F "tokens=2 delims=^&amp;" %%c IN ("%%b") DO (
  
@rem | In Windows Registry search for branch where stored certain DeviceID we got in prevous step
FOR /F %%d IN ('REG QUERY HKLM\SYSTEM\CurrentControlSet\Control\Class\{4D36E968-E325-11CE-BFC1-08002BE10318} /f %%c /s /t REG_SZ ^| find "{"') DO (
  
@rem | Now that we've found registry branch with DeviceID, serch in this branch proper amount of GPU memory
@rem | Use powershell to convert memory value from HEX to decimal 
@rem | Use powershell to divide decimal memory value by 1048576 to get memory siZe in Mb
FOR /F "tokens=3" %%e IN ('REG QUERY ^"%%d^" /f HardwareInformation.qwMemorySize /t REG_QWORD ^| find "0x"') DO (
SET vc_mem=powershell [uint64]^('%%e'^)/1048576
)
)
)
  
@rem | Write GPU name 
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>!vc_name!^</div^> >> %~dp0%pc_name%.html
  
@rem | Write GPU's memory size
@rem | If memory size not defined write "N/A or DVMT" (For in most cases it is iGPU)
ECHO ^<div class=^"div-table-cell^"^> >> %~dp0%pc_name%.html
IF !vc_mem! LSS 1 (SET vc_mem=N/A or DVMT&ECHO !vc_mem! >> %~dp0%pc_name%.html) ELSE (!vc_mem! >> %~dp0%pc_name%.html&ECHO Mb >> %~dp0%pc_name%.html)
ECHO ^</div^>^</div^> >> %~dp0%pc_name%.html
@rem | В конце цикла пишем 0 в переменную vc_mem
@rem | В противном случае следующая итерация цикла будет использовать значение переменной от предыдущей итерации
@rem | Что приведет к ошибкам
SET vc_mem=0
)
  
@rem | Write info on network adapters
ECHO ^<div class=^"div-table-row^"^>Network Adapters^</div^> >> %~dp0%pc_name%.html
  
@rem | Get string with MAC as first token and NetworkAdapter name as all other tokens in string
@rem | Write all tokens except first - it is NetworkAdapter name
@rem | Write first token - it is MAC Address
FOR /F "tokens=1*" %%p IN ('wmic path Win32_NetworkAdapter where PhysicalAdapter^=true get macaddress^,name ^| findstr [0-9]') DO (
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell-long^"^>%%q^</div^> >> %~dp0%pc_name%.html
ECHO ^<div class=^"div-table-cell^"^>%%p^</div^>^</div^> >> %~dp0%pc_name%.html
)
  
@rem | Close first div-table with info on hardware
ECHO ^</div^> >> %~dp0%pc_name%.html
  
@rem | Open secon div-table for info on software and network settings
ECHO ^<div class=^"div-table-sec^"^> >> %~dp0%pc_name%.html
  
@rem | Write info on Operating System
ECHO ^<div class=^"div-table-row^"^>Operating System^</div^> >> %~dp0%pc_name%.html
FOR /F "skip=2 tokens=2,3 delims=," %%i IN ('wmic path win32_OperatingSystem get Caption^, Version /format:csv') DO (
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>Name^</div^>^<div class=^"div-table-cell^"^>%%i^</div^>^</div^> >> %~dp0%pc_name%.html
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>Version^</div^>^<div class=^"div-table-cell^"^>%%j^</div^>^</div^> >> %~dp0%pc_name%.html
)

@rem | Write OS Architectire. Use separated command for comaptibility with WinXP (Yeah, XP in 2022) for it has no OSArchitecture method
FOR /F "skip=2 tokens=2* delims=," %%i IN ('wmic path Win32_Processor where DeviceID^=^"CPU0^" get DataWidth /format:csv') DO (
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>Architecture^</div^>^<div class=^"div-table-cell^"^>%%i^</div^>^</div^> >> %~dp0%pc_name%.html
)  

@rem | Write Windows activation status
FOR /F "tokens=* skip=1" %%i IN ('cscript.exe /nologo c:\windows\system32\slmgr.vbs /xpr') DO (
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>Activation^</div^>^<div class=^"div-table-cell^"^>%%i^</div^>^</div^> >> %~dp0%pc_name%.html
)

@rem | Write Info on MS Office (if any) version and activation status

FOR /F "tokens=2,3 delims=," %%i IN ('wmic path Win32_Product where "Vendor='Microsoft Corporation' AND Name LIKE '%%Office%%'" get InstallLocation^, Name /format:csv ^| findstr "Standard Professional 365"') DO (
SET off_path=%%i
SET off_name=%%j
)

@rem | CD to Office Dir
pushd %off_path%

@rem | Get Path to OSPP.VBS and rewrite off_path variable
@rem | SO thet it contains path to ospp.vbs instead of path to MSO dir
FOR /F "tokens=2*" %%i IN ('dir /s ospp.vbs ^| find "\"') DO (
SET off_path=%%j
)

@rem | CD to OSPP.VBS dir
pushd %off_path%

@rem | Catch exception if no OSPP.VBS exists
@rem | Then goto next info block
IF %ERRORLEVEL% EQU 1 (GOTO L4)

@rem | Через файл OSPP.VBS получаем информацию о статусе лицензии
@rem | И отфильтровываем от всего лишнего
FOR /F "skip=3 tokens=2 delims=:" %%i IN ('cscript ospp.vbs /dstatus ^| find "-"') DO (
SET act_stat=%%i
)

@rem | Write info on MS Office
ECHO ^<div class=^"div-table-row^"^>MS Office^</div^> >> %~dp0%pc_name%.html
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>%off_name%^</div^>^<div class=^"div-table-cell^"^>%act_stat%^</div^>^</div^> >> %~dp0%pc_name%.html

@rem | Jump here if no MS Office installed
:L4

@rem | Write info on network connectors
@rem | Get mac address, Connection name and Connection status in CSV format
FOR /F "skip=2 tokens=2-4 delims=," %%i IN ('wmic path Win32_NetworkAdapter where "PhysicalAdapter='true'" get MACAddress^, NetConnectionID^, NetConnectionStatus /format:csv') DO (

@rem | For all variables in CMD after FOR loop has string type 
@rem | Change connection status type to digit by adding zero and assign to con_stat variable
@rem | Otherwise IF statement WILL NOT work properly because it will have to compare digit with string
SET /A con_stat=%%k+0

@rem | IF connection status is 2 (active, connected) 
@rem | Write connection name
IF !con_stat! EQU 2 (
ECHO ^<div class=^"div-table-row^"^>%%j^(Active^)^</div^> >> %~dp0%pc_name%.html

@rem | Write MAC address of active connection
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>MAC Address^</div^>^<div class=^"div-table-cell^"^>%%i^</div^>^</div^> >> %~dp0%pc_name%.html
  
@rem | Write IP address, Network mask and Gateway of active connection
@rem | Get network info on active connection in "VALUE" format
FOR /F "tokens=1,2 delims=^={}" %%a IN ('wmic path Win32_NetworkAdapterConfiguration where "ipenabled='true' AND MACADDRESS='%%i'" get DefaultIPGateway^, IPAddress^, IPSUbnet /format:value ^| findstr [0-9]') DO (

rem | Use inner FOR cycle to get rid of IPv6 address and netmask additional info, if any
rem | feed to FOR cycle variable where stored both IPv4 and IPv6 addresses separated by comma
rem | And take only first token which is IPv4 address
FOR /F "tokens=1 delims=," %%q IN (%%b) DO (

@rem | Get rid of double quotes in output
SET q_temp=%%q
SET q_final=!q_temp:"=!
)

@rem | Write results
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>%%a^</div^>^<div class=^"div-table-cell^"^>!q_final!^</div^>^</div^> >> %~dp0%pc_name%.html
)

@rem | ELSE if connection status is NOT 2
) ELSE (
@rem | Write inactive connection name 
ECHO ^<div class=^"div-table-row^"^>%%j^(Inactive^)^</div^> >> %~dp0%pc_name%.html

@rem | Write inactive connection MAC Address 
ECHO ^<div class=^"div-table-row^"^>^<div class=^"div-table-cell^"^>MAC Address^</div^>^<div class=^"div-table-cell^"^>%%i^</div^>^</div^> >> %~dp0%pc_name%.html
) 
)

@rem | Close second div-table
ECHO ^</div^> >> %~dp0%pc_name%.html
  
@rem | Close main positioning div
ECHO ^</div^> >> %~dp0%pc_name%.html

@rem | Wait for user interaction
pause
