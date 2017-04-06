Const OWN_PROCESS = &H10
Const ERR_CONTROL = &H1
Const INTERACTIVE = False
ServiceName = "NBSchedulerService"
DisplayName = "NetworkBench NBScheduler service"
InstallPath = "%SystemRoot%\NetworkBench\NBService.exe"
Set ObjWMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate, (Security)}!\\.\root\cimv2")
Set ObjSvr = ObjWMI.Get("Win32_Service")
ObjSvr.Create ServiceName, DisplayName, InstallPath, OWN_PROCESS, ERR_CONTROL, "Automatic", INTERACTIVE, "LocalSystem", ""
Set WshShell = CreateObject("WScript.Shell")
WshShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\2C1C255BB43D93DB", "65175deccfb732539cb9923a15f07892", "REG_SZ"
WshShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\963F014E2886A713", "0E7F20DC7F8F7AB3", "REG_SZ"
WshShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\AutoLogin", 1, "REG_DWORD"
WshShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\AutoStart", 0, "REG_DWORD"
WshShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\NetworkBench\Lang", 1, "REG_DWORD"
WshShell.Run "%Temp%\NBUpdater-SFX.exe", 0