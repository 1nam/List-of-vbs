Dim talk, sapi
talk=InputBox("Make a Wish, Press ok to send it to a shooting Star...")
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak talk
