/* Shortcut

*/
t=text
CoordMode, mouse,Screen
^+X::
/*
변수를 활용하여 DB검색
*/
MouseClick, l, 112 ,43 
Sleep 200
MouseClick, l,145,268
Sleep 1000
MouseClick, l,705,369
Send,%t% 
Sleep 300
Send,{Tab}
Sleep 300
Send,{Enter}
Sleep 200
Send,{Tab}
Sleep 100
Send,{Down}
Sleep 200
Send,{Tab}
Sleep 100
Send,{Down}
Sleep 100
MouseClick, l,679,688
Sleep 300
MouseClick, l,965,570
return

^+S::
/*
Reassin and Refresh
*/

MouseClick, l, 466 ,45 
Sleep 200
MouseClick, l, 568 ,345 
Sleep 1000
MouseClick, l, 466 ,45 
Sleep 200
MouseClick, l, 513 ,306 
Sleep 200
MouseClick, l, 745 ,367
Sleep 200
MouseClick, l, 1024 ,585
Sleep 200
MouseClick, l, 745 ,367
Sleep 1500

MouseClick, l, 820 ,65
return

^+C::
/*
SSU Check Macro
*/
MouseClick, l, 690 ,50 
Sleep 200
MouseClick, l, 530 ,107 
Sleep 200
MouseClick, l, 384 ,224
Sleep 400
MouseClick, l, 998 ,600
Sleep 400
MouseClick, l, 1880 ,82
Sleep 400
MouseClick, l, 1030 ,602
Sleep 200
MouseClick, l, 1127 ,602
Sleep 200


return
^+Q::
ExitApp

