' AI BeiKe Assistant Launcher
' Fix: kill stale process on port 5000 before launch to avoid port conflict
Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
strPath = FSO.GetParentFolderName(WScript.ScriptFullName)
scriptPath = FSO.BuildPath(strPath, "web_app.py")
pythonPath = FSO.BuildPath(strPath, ".venv\Scripts\python.exe")

' Set working directory
WshShell.CurrentDirectory = strPath

' 1. Check Python availability in virtual environment
If Not FSO.FileExists(pythonPath) Then
    MsgBox "Virtual environment Python not found. Please run install.bat first.", vbCritical, "Error"
    WScript.Quit(1)
End If

' 2. Kill any stale process listening on port 5000 to avoid port conflict
On Error Resume Next
Dim netOut, lines, i, parts, pid
Set exec = WshShell.Exec("netstat -ano -p tcp")
netOut = exec.StdOut.ReadAll()
If Err.Number = 0 Then
    lines = Split(netOut, vbCrLf)
    For i = 0 To UBound(lines)
        If InStr(lines(i), ":5000") > 0 And InStr(lines(i), "LISTENING") > 0 Then
            parts = Split(Trim(lines(i)))
            If UBound(parts) >= 4 Then
                pid = Trim(parts(UBound(parts)))
                If IsNumeric(pid) And pid <> "0" Then
                    WshShell.Run "taskkill /F /PID " & pid, 0, True
                End If
            End If
        End If
    Next
End If
On Error GoTo 0

' 3. Launch web_app.py in background (hidden window)
'    Use python.exe (not pythonw) because web_app.py has print() statements.
'    WindowStyle=0 hides the console window.
Dim q : q = Chr(34)
WshShell.Run q & pythonPath & q & " " & q & scriptPath & q, 0, False

Set WshShell = Nothing
Set FSO = Nothing
