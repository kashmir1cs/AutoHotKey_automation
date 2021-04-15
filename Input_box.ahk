InputBox, n, check_macro, n=?,,150,125
if n is Integer
{
	MsgBox, iteration n="%n%"
	loop,%n%
	{
		MsgBox, your num = %A_Index%
		sleep, 100
	}

	return
}
else
{
	MsgBox, "wrong value or cancled"
	return
}	
