' Define the path to your custom wallpaper
customWallpaper = "..\hellome.jpg"

' Create a loop to reset the wallpaper every 15 seconds
Do
    ' Set the wallpaper in the registry for the current user
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", customWallpaper, "REG_SZ"
    
    ' Apply the wallpaper change by refreshing the desktop
    WshShell.Run "RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters ,1 ,True", 0, True

    ' Wait for 15 seconds
    WScript.Sleep 15000  ' Time in milliseconds (15000 ms = 15 seconds)
Loop
