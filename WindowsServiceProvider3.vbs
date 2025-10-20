Set WshShell = CreateObject("WScript.Shell")

WScript.Sleep 500  ' focus the target window first

i = 0
While i < 9999
    WshShell.SendKeys "%{F4}"  ' Alt+F4
    WScript.Sleep 15000          ' wait 30000 ms between attempts
    i = i + 1
Wend
