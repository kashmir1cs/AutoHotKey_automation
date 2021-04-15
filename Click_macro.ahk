/* short cut
  v1.1 : iteration macro added
  
*/
CoordMode, mouse,Screen
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
/* 
iterable auto check macro
*/

^+D::
InputBox, n, ssu check_macro, n=?,,150,125
if n is Integer
{
	MsgBox, iteration n="%n%"
	loop,%n%
	{
		MouseClick, l, 690 ,50 
		Sleep 200
		MouseClick, l, 530 ,107 
		Sleep 200
		MouseClick, l, 384 ,224
		Sleep 400
		MouseClick, l, 1088 ,600
		Sleep 400
		MouseClick, l, 1880 ,82
		Sleep 400
		MouseClick, l, 1030 ,602
		Sleep 200
		MouseClick, l, 1127 ,602
		Sleep 200
	}
	return
}
else
{
	MsgBox, "wrong value or cancled"
	return
}	



^+Q::
ExitApp
