Set objWmiService = GetObject("winmgmts:")
Set objLogicalDisk = objWmiService.Get ("Win32_LogicalDisk.DeviceID='C:'")
WScript.Echo objLogicalDisk.FreeSpace
