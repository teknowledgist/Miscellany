Dim Words, Speaker, Hours
Words = "It is " & ((hour(time) + 11) Mod 12 + 1) & " O'clock"
if hour(time) = 17 then
   Words = Words & ", and time to go home."
End If
Set Speaker = CreateObject("SAPI.spVoice")
Set Speaker.Voice = Speaker.GetVoices.Item(1)
Speaker.Volume = 50
Speaker.Speak Words
 
