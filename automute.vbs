' Script that mutes Windows for a specified time.
'
' Run at cmd.exe with
' > cscript [filename] [time]
' 
' filename    this script
' time        in ms

Set args = WScript.Arguments
Set WshShell = CreateObject("WScript.Shell")
WshShell.SendKeys(chr(173))
WScript.Sleep args.Item(0)
WshShell.SendKeys(chr(173))
