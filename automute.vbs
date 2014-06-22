' Script that toggles Windows mute for a specified time.

Sub Usage()
WScript.echo("USAGE" & vbCrLf &_
"Run at cmd.exe shell with" & vbCrLf &_
"> cscript FILENAME TIME" & vbCrLf &_
"" & vbCrLf &_
"FILENAME    This script." & vbCrLf &_
"TIME        In seconds. Must be positive. Can be integer or decimal.")
End Sub

' Extracts time from command line arguments
If WScript.Arguments.Count < 1 Then
	Usage()
	WScript.Quit
End If
time_sec = WScript.Arguments.Item(0)
If NOT IsNumeric(time_sec) Then
	Usage()
	WScript.Quit
ElseIf time_sec <= 0 Then
	Usage()
	WScript.Quit
End If
time_ms = 1000 * time_sec 

' Toggle mute
Set WshShell = CreateObject("WScript.Shell")
WshShell.SendKeys(chr(173))

' Sleep
WScript.Sleep time_ms

' Toggle mute again
WshShell.SendKeys(chr(173))
