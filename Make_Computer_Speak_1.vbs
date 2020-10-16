Option Explicit
Dim objVOC : Set objVOC = CreateObject("SAPI.SpVoice")

objVOC.Rate = 2
objVOC.Speak "My name is Stoyan!"