

range := "B3:D4"
StringSplit, out, range, :
MSgbox,% out0
return

NumberToRangeColumnCheck:
Loop,16383
{
    Column := NumberToRangeColumn(A_Index)

    column_Number := RangeColumnToNumber(Column)

    Column2 := NumberToRangeColumn(column_Number)

    num1 := RangeColumnToNumber(Column)
    num2 := RangeColumnToNumber(Column2)

    if (num1 - num2) != 0
    {
        Msgbox,% A_Index . "`n not`n" . num1 . "<>" . num2
    }
}

Return

NumberToRangeColumn(columnNumber)
{
    columnName := ""

    while (columnNumber > 0.5)
    {
        modulo := Mod((columnNumber - 1), 26)
        columnName := Chr(65 + modulo) . columnName
        columnNumber := (columnNumber - modulo) / 26
    } 

    return Trim(columnName)
}

RangeColumnToNumber(range)
{
    StringUpper, range, range
    RegExMatch(range, "[a-zA-Z]+", regexString)

    columnNumber := 0
    chars := Array()
    Loop, parse, regexString
    {   
        chars.Push(A_LoopField)
    }

    if chars.length() >= 4 
        throw, "too much column char"

    if chars.Length() = 1
        columnNumber += ord(chars[1]) - 64

    if chars.Length() = 2
        columnNumber := 26 + (ord(chars[2]) - 64) + 26*(ord(chars[1])-64-1)

    if chars.Length() =3
    {
        ; very hard to figure out this formula.
        columnNumber := 702 + (ord(chars[3]) - 64) 
            + 26*(ord(chars[2]) - 64 - (ord(chars[1]) - 64)) 
            + 702*(ord(chars[1]) - 64 - 1)
    }
    if columnNumber > 16384
        throw, "too big column number for excel keeping."
    return columnNumber
}

classTest:
testFunc(["aaaa","bbbb"]*)
return

testFunc(params*)
{
    Msgbox,% params[1] . params[2]
}

selectNodeTest:
doc := LoadXML("sheet1.xml")

root := doc.documentElement


for nodeItem in ( root.selectNodes( "//c" ))
    Msgbox,% nodeItem.getAttribute("r")

; msgbox % RTrim( descList2, "|" )

return
LoadXML(xml_path)
    {
        doc := ComObjCreate("MSXML2.DOMDocument.3.0")
        doc.async := false
        doc.Load(xml_path)

        Err := doc.parseError
        if Err.reason
        {
            msgbox % "Error: " Err.reason . "`n: " . xml_path
            ExitApp
        }
    return doc
    }

