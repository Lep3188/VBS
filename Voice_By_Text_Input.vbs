

Dim message, sapi

For x=0 to 2
message=InputBox("What do you want me to say?","Voice by Input")
Set sapi=CreateObject("sapi.spvoice")

sapi.Speak message

Next
