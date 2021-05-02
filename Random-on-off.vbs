Randomize
Function RandomWithinRange(min, max)
    RandomWithinRange = Int((max - min + 1) * Rnd() + min)
End Function
Function RandItemFromArray(arr)
    RandItemFromArray= arr(RandomWithinRange(LBound(arr), UBound(arr)))
End Function
Dim names
    names = Array("On", "Off", "No Electricity")
msgbox("Light Switch Test.")
msgbox("Press ok to Start.")
Dim LightSwitch
         LightSwitch = RandItemFromArray(names)
If LightSwitch = "On" Then
         msgbox("Light is on!")
Elseif LightSwitch = "Off" Then
         msgbox("Light is Off!")
Else
           msgbox("No Eleictricity!")
End If
msgbox("Test was Successful !!!.")
