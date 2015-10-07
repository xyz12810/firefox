Set oShell = CreateObject ("Wscript.Shell") 

strPath = Wscript.ScriptFullName
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(strPath)
strFolder = objFSO.GetParentFolderName(objFile) 

Dim NRun
NRun = True
For Each ps In Getobject("winmgmts:\\.\root\cimv2:win32_process").instances_
If UCase(ps.name) = UCase("lantern.2.0.10.exe") Then ps.terminate : NRun = False
Next

Dim strArgs
quo = """"
strArgs = quo & strFolder & "\Firefox.exe" & quo
oShell.Run strArgs, 0, false
strArgs = quo & strFolder & "\lantern.2.0.10.exe" & quo
oShell.Run strArgs, 0, false
