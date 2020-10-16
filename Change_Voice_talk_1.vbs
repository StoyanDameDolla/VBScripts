Option Explicit

Dim Zira, David

Set Zira = CreateObject("SAPI.spVoice")
Set Zira.Voice = Zira.GetVoices.Item(1)

Set David = CreateObject("SAPI.spVoice")
Set David.Voice = David.GetVoices.Item(0)

Zira.Rate = 2
Zira.Speak "Stoyan, are you a stupid motherfucker?"
David.Rate = 3
David.Speak "Shut the fuck up, you retarded bitch!"
Zira.Rate = 4
Zira.Speak "Go fuck a goat!"
David.Volume = 35
David.Speak "May the jedi run freight train on you!"