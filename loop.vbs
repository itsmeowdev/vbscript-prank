do
Set WshShell = Wscript.CreateObject("WScript.Shell")
WshShell.run "Notepad"
WshShell.AppActivate "Notepad"

WScript.Sleep 100
WshShell.SendKeys "Your"
WScript.Sleep 100
WshShell.SendKeys " "
WScript.Sleep 100
WshShell.SendKeys "Computer"
WScript.Sleep 100
WshShell.SendKeys " "
WScript.Sleep 100
WshShell.SendKeys "Is"
WScript.Sleep 100
WshShell.SendKeys " "
WScript.Sleep 100
WshShell.SendKeys "HACKED!"
WScript.Sleep 100
WshShell.SendKeys " "

Dim spV
Set spV = CreateObject("SAPI.spVoice")

spV.Speak "Your Computer is HACKED by Lucifer"
spV.Speak "Your Computer is HACKED by Lucifer"
spV.Speak "Your Computer is HACKED by Lucifer"
spV.Speak "Your Computer is HACKED by Lucifer"
spV.Speak "Your Computer is HACKED by Lucifer"

loop