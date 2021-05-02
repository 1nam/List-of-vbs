msgbox("Quater Flip Has a 50/50 Chance")
rem .5 means 50/50 chance result Heads or Tails.
randomize()
r = rnd()
if r > .5 then
msgbox ("Heads")
else
msgbox ("Tails")
end if
