Dim speaks, speech
Speaks = "Hi your name"
Set speech = CreateObject("SAPI.SpVoice")
Set speech.Voice = speech.GetVoices.Item(1)
speech.Speak speaks