prompt = "This will keep you screen from locking by continuing clicking `F13`"
title = "keep-the-screen-from-locking.vbs"
InputBox prompt, title

Set objWSH = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

strWSHPath = objWSH.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\wscript.exe"
If objFSO.FileExists(strWSHPath) Then
    Set objFile = objFSO.GetFile(strWSHPath)
    MsgBox "WSH Version: " & objFile.Version
Else
    MsgBox "WSH not found."
End If
