#NoEnv
path := "path"
xl :=ComObjCreate("Excel.Application")
xl.Workbooks.Open(path)
xl.visible := false
CoordMode, mouse,Screen


/*엑셀 이용한 일괄 수정
*/


/*사용자 함수 정의 

*/

pipesearch(x)
{
;파이프 검색 후 윈도우창 활성화	
MouseClick l,936,5
MouseClick l,236,11
MouseClick l,63,128
SendInput, %x%
send,{enter}
Sleep, 100
send,^+P
}
pipename(pname)
{
MouseClick l, 1844,86
sleep, 200
MouseClick l, 1844,86
SendInput, %pname%
send,{enter}

}
pipeinput(matl,di)
{

;배관이 PVDF인경우
MouseClick, l, 1812 ,110
sleep, 500
if (matl == "PVDF")
{
	MouseClickDrag,l,1110 ,490,1110,720
	Sleep, 500
	if (di==20)
	{
	
	MouseClick, l, 837,447
	}
	else if (di==25)
	{
	MouseClick, l, 837,464
	}
	else if (di==32)
	{
	MouseClick, l, 837,481
	}
	else if (di==40)
	{
	MouseClick, l, 837,497
	}
	else if (di==50)
	{
	MouseClick, l, 837,508
	}
	else if (di==63)
	{
	MouseClick, l, 837,528
	}
	else if (di==75)
	{
	MouseClick, l, 837,543
	}
	else if (di==90)
	{
	MouseClick, l, 837,560
	}
	else if (di==110)
	{
	MouseClick, l, 837,574
	}
	else if (di==125)
	{
	MouseClick, l, 837,590
	}
	else if (di==140)
	{
	MouseClick, l, 837,718
	}
	else if (di==160)
	{
	MouseClick, l, 837,609
	}
	else if (di==180)
	{
	MouseClick, l, 837,620
	}
	else if (di==200)
	{
	MouseClick, l, 837,638
	}
	else if (di==225)
	{
	MouseClick, l, 837,654
	}
	else if (di==250)
	{
	MouseClick, l, 837,671
	}
	else if (di==280)
	{
	MouseClick, l, 837,686
	}
	else if (di==315)
	{
	MouseClick, l, 837,702
	}
}
else if (matl== "PP")
{
	MouseClickDrag,l,1110 ,455,1110,463
	Sleep,150
	if (di==20)
	{
	MouseClick, l, 837,447
	}
	else if (di==25)
	{
	MouseClick, l, 837,462
	}
	else if (di==32)
	{
	MouseClick, l, 837,478
	}
	else if (di==40)
	{
	MouseClick, l, 837,494
	}
	else if (di==50)
	{
	MouseClick, l, 837,511
	}
	else if (di==63)
	{
	MouseClick, l, 837,526
	}
	else if (di==75)
	{
	MouseClick, l, 837,543
	}
	else if (di==90)
	{
	MouseClick, l, 837,558
	}
	else if (di==110)
	{
	MouseClick, l, 837,575
	}
	else if (di==125)
	{
	MouseClick, l, 837,590
	}
	else if (di==140)
	{
	MouseClick, l, 837,607
	}
	else if (di==160)
	{
	MouseClick, l, 837,623
	}
	else if (di==180)
	{
	MouseClick, l, 837,640
	}
	else if (di==200)
	{
	MouseClick, l, 837,657
	}
	else if (di==225)
	{
	MouseClick, l, 837,671
	}
	else if (di==250)
	{
	MouseClick, l, 837,687
	}
	else if (di==280)
	{
	MouseClick, l, 837,702
	}
	else if (di==315)
	{
	MouseClick, l, 837,718
	}
}
}

pipelength(lng)
{
;배관길이 입력
MouseClick l,1825,137
Sleep,500
MouseClick l,1825,137
SendInput, %lng%
send,{enter}
Sleep,200
MouseMove 836,530
}

pipetype(typ,ea)
{

MouseClick, l, 1800 ,190
sleep, 300
if (typ=="B")
{
send,{down 3}
sendraw,%ea%
sleep, 50
send,{down 6}
sendraw,%ea%
sleep, 50
send,{down 28}
sleep, 50
sendraw,%ea%
sleep, 50
MouseClick, l, 1117 ,703	
}
else if (typ=="E")
{
send,{down 3}
sendraw,%ea%
sleep, 50
MouseClick, l, 1117 ,703	
}
else if (typ=="D")
{
send,{down 3}
sendraw,%ea%
sleep, 50
send,{down 6}
sendraw,%ea%
sleep, 50
send,{down 7}
sendraw,%ea%
sleep, 50
MouseClick, l, 1117 ,703
}


}


F2::
Loop
{
a:= A_Index+5
pno :=xl.Cells(a ,2).Value ;Pipe No 검색
pchange :=xl.Cells(a ,3).Value ;Pipe No. 변경
matl :=xl.Cells(a ,4).Value ;matl
size :=xl.Cells(a ,5).Value ;size
t :=xl.Cells(a ,6).Value ;type
qty :=xl.Cells(a ,7).Value ;qty
length :=xl.Cells(a ,8).Value ;length
pipesearch(pno)
Sleep, 200
pipelength(length)
Sleep, 200
pipeinput(matl,size)
Sleep, 200
pipetype(t,qty)
if (pno=="")
{
	xl.activeworkbook.Close
	ExitApp
}	


}


F4::
xl.activeworkbook.Close
ExitApp
