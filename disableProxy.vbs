Set oReg   = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
sKeyPath   = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"
sValueName = "DefaultConnectionSettings"

' Get registry value where each byte is a different setting.
oReg.GetBinaryValue &H80000001, sKeyPath, sValueName, bValue

' Check byte to see if detect is currently on.
If (bValue(8) And 8) = 8 Then
    ' To change the value to Off.
    bValue(8) = bValue(8) And Not 8
End If

'Write back settings value
oReg.SetBinaryValue &H80000001, sKeyPath, sValueName, bValue

Set oReg = Nothing