Option Explicit

Dim fso, shell, xampp, app, cmd, http, ready, i
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

xampp = ""
If fso.FileExists("C:\xampp\xampp-control.exe") Then
    xampp = "C:\xampp"
ElseIf fso.FileExists("D:\xampp\xampp-control.exe") Then
    xampp = "D:\xampp"
ElseIf fso.FileExists("E:\xampp\xampp-control.exe") Then
    xampp = "E:\xampp"
End If

If xampp = "" Then
    MsgBox "XAMPP was not found." & vbCrLf & vbCrLf & _
           "Please install XAMPP first.", vbExclamation, "ReclaimHub"
    WScript.Quit
End If

app = xampp & "\htdocs\reclaimhub"

If Not fso.FolderExists(app) Then
    fso.CreateFolder(app)
End If

If Not fso.FolderExists(app & "\uploads") Then
    fso.CreateFolder(app & "\uploads")
End If

' Copy the final ReclaimHub files into XAMPP.
fso.CopyFile fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "index.php"), _
             app & "\index.php", True
fso.CopyFile fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "database.sql"), _
             app & "\database.sql", True

' Start Apache and MySQL without showing a CMD window.
shell.Run """" & xampp & "\apache_start.bat""", 0, False
shell.Run """" & xampp & "\mysql_start.bat""", 0, False

' Give MySQL a moment to start, then import the database silently.
WScript.Sleep 4000
cmd = """" & xampp & "\mysql\bin\mysql.exe"" -u root < """ & app & "\database.sql"""
shell.Run "cmd.exe /c " & cmd, 0, True

' Wait until the web server responds.
ready = False
For i = 1 To 30
    WScript.Sleep 1000
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "http://localhost/reclaimhub/", False
    http.Send
    If Err.Number = 0 Then
        If http.Status >= 200 And http.Status < 500 Then
            ready = True
            Exit For
        End If
    End If
    Err.Clear
    On Error GoTo 0
Next

shell.Run "http://localhost/reclaimhub/", 1, False
