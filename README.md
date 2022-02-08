Script display main hardware info of local PC, info on MS Windows version and activations status, info on MS Office (if any) version and activation status, info on network adapters and active connections.  
For example, if you save script as D:\hw-inf.bat and your system name is I010203 script output will be D:\I010203.html  
P.S Script correctly display memory size for video adapters with more than 4Gb memory, since script does not use AdapterRAM, but search in Windows Registry for HardwareInformation.hwMemorySize property which contains correct memory size.  
P.P.S. It is supposed that all drivers are unstalled.  
P.P.P.S. Will not work on WindowsPE unless you integrate wmic in your PE environment.
