Script collects main hardware info and network settings and store it in same folder from which it was launched as simple html-file with the same name as system name. 
For example, if you save script as D:\hw-inf.bat and your system name is I010203, script output will be D:\I010203.html 
P.S Script correctly display memory size for video adapters with more then 4Gb memory, since script does not use AdapterRAM, but search in Windows Registry for HardwareInformation.hwMemorySize property which contains correct memory size.  
P.P.S. It is supposed that all drivers are unstalled. 
P.P.P.S. Will not work on WindowsPE unless you integrate wmic in your PE environment.
