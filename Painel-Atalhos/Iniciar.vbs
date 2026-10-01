Option Explicit
Dim fso, shell, scriptDir, psExe, command
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
psExe = shell.ExpandEnvironmentStrings("%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe")
command = Chr(34) & psExe & Chr(34) & " -NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File " & Chr(34) & fso.BuildPath(scriptDir, "Atalhos.ps1") & Chr(34)
On Error Resume Next
shell.Run command, 0, False
If Err.Number <> 0 Then
    MsgBox "Nao foi possivel abrir o painel: " & Err.Description & vbCrLf & "Tente o arquivo Iniciar.cmd.", vbExclamation, "Meus atalhos"
End If
