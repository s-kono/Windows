Dim fs
set fs = CreateObject("Scripting.FileSystemObject")
Dim f
f = fs.GetParentFolderName(WScript.Arguments.Item(0))
Dim WshShell
Set WshShell = CreateObject ("Wscript.Shell")
WshShell.CurrentDirectory = f

WshShell.Run("cmd")
