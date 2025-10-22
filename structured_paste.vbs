Dim objShell
Set objShell = CreateObject("WScript.Shell")
' Pass the argument (%V\.) to the script
objShell.Run "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File ""C:\Windows\structured_paste.ps1"" """ & WScript.Arguments(0) & """", 0, True