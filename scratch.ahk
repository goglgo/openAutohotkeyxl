; xml := ComObjCreate("MSXML2.DOMDocument.6.0")

; xml.Load("sharedStrings.xml")
; xml.async := false
; xml.setProperty("SelectionLanguage", "XPath")
; ns1 := "xmlns:main='http://schemas.openxmlformats.org/spreadsheetml/2006/main'"
; xml.setProperty("SelectionNamespaces" , ns1)

; Msgbox,% xml.DocumentElement.selectNodes("//main:si").length

; return


; https://docs.microsoft.com/en-us/previous-versions/troubleshoot/msxml/msxml-6-matching-nodes-not-return

ns1 := "xmlns:main='http://schemas.openxmlformats.org/spreadsheetml/2006/main'"
ns2 := "xmlns:x14ac='http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac'"
ns3 := "xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'"
ns4 := "xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006'"

ns := Format("{1} {2} {3} {4}", ns1, ns2, ns3, ns4)
xml := ComObjCreate( "MSXML2.DOMDocument.6.0" )

xml.Load("sheet1.xml")
xml.async := false
xml.setProperty("SelectionLanguage", "XPath")
xml.setProperty("SelectionNamespaces" , ns)

; tt := xml.selectNodes( "//row/c[@r='B2']" )
; tt := xml.DocumentElement.selectNodes("//main:c")
; ee := xml.DocumentElement.selectSingleNode("//main:c[@r='B4']")
ee := xml.DocumentElement.selectSingleNode("//main:sheetData")
; dd := root.selectSingleNode( "//main:row/main:c[@main:r='B3']/main:v" )
Msgbox,% ee.xml
; MSgbox,% tt.item(1).text

; for k in tt
;     MSgbox,% k.xml

Return
; doc := LoadXML("sheet1.xml")

doc := ComObjCreate( "MSXML2.DOMDocument.6.0" )
doc.setProperty("SelectionLanguage", "XPath")
doc.Load("sheet1.xml")
root := doc.documentElement

; //row/c[@r="B2"]
; //row/c[@r="B2"]/v/text()

; tt := root.selectNodes( "//row/c[@r='B3']/v" )
tt := root.selectSingleNode( "//row/c[@r='B3']/v" )
; dd := root.selectSingleNode("//row/c[@r=""B2""]/v")

for nodeItem in tt
    Msgbox,% nodeItem.xml

; for nodeItem in ( rootd.selectNodes( "//row[@r='2']" ), descList := "" )
;     descList .= nodeItem.text "|"


return

a := [1,2,3]
Msgbox,% a[1]
return

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

