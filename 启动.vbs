Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

projectPath = fso.GetParentFolderName(WScript.ScriptFullName)

' 1. Start Vite in foreground window (visible)
cmd = "cmd /c cd /d """ & projectPath & """ && set NODE_OPTIONS=--openssl-legacy-provider && npm run start"
shell.Run cmd, 1, False

' 2. Check if service is ready
Function IsServiceReady(url)
  On Error Resume Next
  Dim http
  Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
  http.setTimeouts 2000, 2000, 2000, 2000
  http.open "GET", url, False
  http.send
  If Err.Number = 0 Then
    If http.status = 200 Then
      IsServiceReady = True
    Else
      IsServiceReady = False
    End If
  Else
    IsServiceReady = False
  End If
  On Error GoTo 0
End Function

launcherUrl = "http://127.0.0.1:5173/launcher.html"
serviceReady = False

For i = 1 To 60
  If IsServiceReady(launcherUrl) Then
    serviceReady = True
    Exit For
  End If
  WScript.Sleep 1000
Next

If Not serviceReady Then
  MsgBox "Service failed to start!" & vbCrLf & vbCrLf & _
         "1. Make sure Node.js is installed" & vbCrLf & _
         "2. Run 'npm install' in project folder (first time)" & vbCrLf & _
         "3. Check if port 5173 is in use", _
         16, "Startup Failed"
  WScript.Quit
End If

' 3. Find Edge
edgePaths = Array( _
  "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", _
  "C:\Program Files\Microsoft\Edge\Application\msedge.exe" _
)

edge = ""
For Each p In edgePaths
  If fso.FileExists(p) Then
    edge = p
    Exit For
  End If
Next

' 4. Open launcher with Edge app mode
If edge <> "" Then
  shell.Run """" & edge & """ --app=" & launcherUrl & " --window-size=1200,850", 1, False
Else
  shell.Run launcherUrl, 1, False
End If