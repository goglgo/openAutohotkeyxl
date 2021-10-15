X1 := ComObjCreate("Excel.Application")

X1.Workbooks.Open(A_ScriptDir . "\testtt.xlsx")

X1.Visible := True
Msgbox,1

return

WritingTest:
X1 := ComObjCreate("Excel.Application")

X1.Workbooks.Open(A_ScriptDir . "\aaaa.xlsx")

X1.Visible := True
Msgbox,1
; X1.Sheets("TestSheet1").Range("B1:B200").Value 

Loop,4
{
    timeBefore := A_TickCount
    loop,10000
    {
        X1.Sheets("TestSheet1").Range("B" . A_Index).Value := "aaa"
    }

    ; Msgbox,% A_TickCount - timeBefore
    FileAppend, % A_TickCount -timeBefore, ComObjResult.txt
}

return