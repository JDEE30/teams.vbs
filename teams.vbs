Option Explicit
Dim URL,WshShell,i
URL = "https://teams.microsoft.com/l/meetup-join/add-your-meeting-url-here"
Set WshShell = CreateObject("WScript.shell")
For i = 0 to 50
    WshShell.SendKeys(chr(175))
Next
WshShell.run "CMD /C start chrome.exe " & URL & "",0,False
