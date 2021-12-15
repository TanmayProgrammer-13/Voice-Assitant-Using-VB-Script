Set Sapi = Wscript.CreateObject("SAPI.SpVoice")
set wshshell = wscript.CreateObject("wscript.shell")

dim Input

Input = InputBox("Enter Command")

if Input = "hello" OR Input = "hi" then
Sapi.Speak "Hello Sir"
Sapi.Speak "How are you"
wshshell.run "Assitant.vbs"

else
if Input = "open calculator" then
Sapi.Speak "Opening Calculator"
wshshell.run "calc"
wshshell.run "Assitant.vbs"

else
if Input = "open youtube" then
Sapi.Speak "Opening youtube"
wshshell.run "https://www.youtube.com/"
wshshell.run "Assitant.vbs"

else

if Input = "open fileX" then
Sapi.Speak " opening file explorer"
wshshell.run "explorer.exe"
wshshell.run "Assitant.vbs"

else
if Input = "open google" then
Sapi.Speak " opening google"
wshshell.run "https://www.google.com/"
wshshell.run "Assitant.vbs"

else
if Input = "open spotify" then
Sapi.Speak " opening spotify"
wshshell.run "Spotify.exe"
wshshell.run "Assitant.vbs"

else
if Input = "open chrome" then
Sapi.Speak "opening Chrome"
wshshell.run "chrome.exe"
wshshell.run "Assitant.vbs"

else
if Input = "open notepad" then
Sapi.Speak "opening notepad" 
wshshell.run "notepad.exe"
wshshell.run "Assitant.vbs"

else
if Input = "open vs code" then
Sapi.Speak "opening visual studio code" 
wshshell.run "Code.exe"
wshshell.run "Assitant.vbs"

else
if Input = "thanks" then
Sapi.Speak "welcome sir!" 
wshshell.run "Assitant.vbs"

else
if Input = "open whatsapp" then
Sapi.Speak "opening whatsapp" 
wshshell.run "https://web.whatsapp.com/"
wshshell.run "Assitant.vbs"

else
if Input = "i am fine" then
Sapi.Speak "great sir! Have a Nice Day" 
wshshell.run "Assitant.vbs"

else
if Input = "Are You Better Than Alexa"  OR Input = "are you better than alexa" then
Sapi.Speak "I Known I am Not Better But I'will Try My Best" 
wshshell.run "Assitant.vbs"

else
if Input = "open photoshop" then
Sapi.Speak "opening photoshop" 
wshshell.run "photoshop.exe"
wshshell.run "Assitant.vbs"


end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
