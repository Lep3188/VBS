'Powerball ticket generator.

Function TicketRun()

Dim max, max2,min,rand1,rand2,rand3,rand4,rand5,rand6
max=69
max2=27
min=1
Randomize
rand1 = Int((max-min+1)*Rnd+min)
rand2 = Int((max-min+1)*Rnd+min)
rand3 = Int((max-min+1)*Rnd+min)
rand4 = Int((max-min+1)*Rnd+min)
rand5 = Int((max-min+1)*Rnd+min)
rand6 = Int((max2-min+1)*Rnd+min)
WScript.Echo ("Random Ticket is: " & rand1 & "-" & rand2 & "-" & rand3 & "-" & rand4 & "-" & rand5 & "-" & rand6)

End Function

TicketRun