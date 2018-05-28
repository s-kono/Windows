Set objShell = CreateObject("WScript.Shell")
Dim strCommand
Dim strOutput

'strCommand = "powershell -NoLogo -Command Get-FileHash " & WScript.Arguments.Item(0) & " -algorithm MD5 | Format-list | Out-String -Stream | select-string -NotMatch -Pattern Path | select-string -Pattern '\W'; powershell -NoLogo -Command Get-FileHash " & WScript.Arguments.Item(0) & " -algorithm SHA1 | Format-list | Out-String -Stream | select-string -NotMatch -Pattern Path | select-string -Pattern '\W'; powershell -NoLogo -Command Get-FileHash " & WScript.Arguments.Item(0) & " -algorithm SHA256 | Format-list | Out-String -Stream | select-string -NotMatch -Pattern Path | select-string -Pattern '\W'"

strCommand = "powershell -NoLogo -Command "" Get-FileHash " & WScript.Arguments.Item(0) & " -algorithm MD5 | Format-list | Out-String -Stream | select-string -NotMatch -Pattern Path | select-string -Pattern '\W' ; echo '' ; Get-FileHash " & WScript.Arguments.Item(0) & " -algorithm SHA1 | Format-list | Out-String -Stream | select-string -NotMatch -Pattern Path | select-string -Pattern '\W' ; echo '' ; Get-FileHash " & WScript.Arguments.Item(0) & " -algorithm SHA256 | Format-list | Out-String -Stream | select-string -NotMatch -Pattern Path | select-string -Pattern '\W' "" "
'WScript.Echo strCommand

Set objExec = objShell.Exec(strCommand)
objExec.StdIn.Close()
strOutput = vbNewLine & WScript.Arguments.Item(0) & vbNewLine & objExec.StdOut.Readall()

WScript.Echo strOutput
set objExec = Nothing
set objShell = Nothing

WScript.StdIn.ReadLine
