Option Explicit

Dim WshShell, fso
Dim scriptFolder, startBatPath, cmd, q

On Error Resume Next
Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

If Err.Number <> 0 Then
    MsgBox "Blad inicjalizacji skryptu: " & Err.Description, vbCritical, "TrackMyDay - Blad"
    WScript.Quit Err.Number
End If
Err.Clear

' Zawsze pracuj na folderze, w ktorym lezy ten skrypt VBS.
scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
startBatPath = fso.BuildPath(scriptFolder, "start.bat")

If Not fso.FileExists(startBatPath) Then
    MsgBox "Blad: Nie znaleziono pliku start.bat w: " & scriptFolder, vbCritical, "TrackMyDay - Blad"
    WScript.Quit 1
End If

' Uruchom start.bat w ukrytym oknie bez czekania na zakonczenie.
q = Chr(34)
cmd = "cmd /c " & q & "cd /d " & q & scriptFolder & q & " && call " & q & startBatPath & q & q
WshShell.Run cmd, 0, False

If Err.Number <> 0 Then
    MsgBox "Blad podczas uruchamiania start.bat: " & Err.Description, vbCritical, "TrackMyDay - Blad"
    WScript.Quit Err.Number
End If

Set fso = Nothing
Set WshShell = Nothing
' Brak MsgBox sukcesu - tryb cichy.
