Dim WshShell
Set WshShell = CreateObject ("Wscript.Shell")
WshShell.Run( "cscript //nologo " & Replace(WScript.ScriptFullName,WScript.ScriptName,"") & "get-filehash_ps4.vbs " & WScript.Arguments.Item(0) )
