; // test code for sheet class below.
; tt := new Sheet("sheet1.xml", "sharedStrings.xml")
; Msgbox,% tt.Range("B3").text
; tt.Range("z5") := "tttt"
; Return
; //

; TODO: change architecture.
; testing for Sheet("B3").Range("B3").value

; tt := new naduretest("zzzz")
; tt.value := "Asdf"
; Msgbox,% tt
; return


; rng := new RangeClass("B3")
; rng.tt := "asdf"


; newSheetXMLFormat : New Sheet XML


; sheet := new Sheet("sheet1.xml", "sharedStrings.xml")
; sheet.Range("b3").value := "asdfas"
; return


class BaseMethod
{

    findNode(xmlnodes, nodename:="")
    {
        for k, v in xmlnodes
        {
            if k.nodeName = nodename
            {
                return k
            }
            
            
            if k.hasChildNodes()
            {
                result := this.findNode(k.childNodes, nodename)
                if result
                    return result
            }
                
        }
        
    }

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

    sheetData()
    {
        doc := this.LoadXML(this.sheetXML)
        this.sheetDataDoc := doc
        found := this.findNode(doc.childNodes, "sheetData")
        if not found
            throw,"There is no found at the Sheet. please check sheet.xml."
        return found
    }

    SharedStrings()
    {
        doc := this.LoadXML(this.sharedStringsXML)
        this.sharedStringsDoc := doc
        tTags:= doc.getElementsByTagName("t") 

        ; it has no __ENum. so rearrange.
        result := Array()
        for k, v in tTags
            result.Push(k)
        return result 
    }

    ; It seems not works when use with method or func.
    ; selectXPathNodes(xmlDoc, XPathString)
    ; {
    ;     found := Array()
    ;     for nodeItem in xmlDoc.selectNodes(XPathString)
    ;     {
    ;         ; Msgbox,% nodeItem.xml
    ;         found.Push(nodeItem)
    ;     }
    ;     return found
    ; }
}


; Sheet class
class Sheet extends BaseMethod
{
    __New(sheetXML:="", sharedStringsXML:="")
    {
        if not FileExist(sheetXML)
            throw, "Can't find sheet.xml file."

        if not FileExist(sharedStringsXML)
            throw, "Can't find sharedStrings.xml file."

        this.sheetXML := sheetXML
        this.sharedStringsXML := sharedStringsXML

        this.isThisSheetDeleted := False
    }

    Range(params*)
    {
        if this.isThisSheetDeleted
            throw, "This sheet is already deleted."

        rangeClass := new RangeClass(this.sheetXML
            , this.sharedStringsXML, params*)
            
        rangeClass.RangeColumnToNumber := this.RangeColumnToNumber
        return rangeClass
    }

    DeleteSheet()
    {
        if this.isThisSheetDeleted
            throw, "This sheet is already deleted."
        ; delete sheetN file
        ; workbook에서 sheetN 제거
        ; app vt:i4에서 숫자 -1
        ; vt:lpstr 에서 시트이름 하나 제거
        ; vt:vector 에서 size -1
        ; ContentType에서 sheetN 제거

        ; Step1: Delete this.sheetXMl file
        WorkSheetsPathList := this.paths.WorkSheetsPathList
        FileDelete,% this.sheetXML

        ; GetSheetNo
        SplitPath, % this.sheetXML, , , , sheetXMLFileName
        RegExMatch(sheetXMLFileName, "\d+$", SheetNo)
        SheetNo -= 1
        
        ; Step2
        app := this.loadXml(this.paths.app)
        ; remove sheet[N] row
        found := app.documentElement.selectNodes("//vt:vector/vt:lpstr")
        ; Msgbox,% found.item(0).xml ; Start from Zero
        deleteSheetRow := found.item(SheetNo)
        deleteSheetRow.ParentNode.removeChild(deleteSheetRow)

        ; touch vt:i4
        found := app.documentElement.selectNodes("//vt:i4")
        foundItem := found.item(0)
        foundItem.text -= 1
        size := foundItem.text

        ; touch vector size
        found := app.documentElement.selectNodes("//TitlesOfParts[0]/vt:vector[0]")
        found.item(0).setAttribute("size", size)

        ; app file done
        app := app.save(this.paths.app)

        ; Reorder sheet N Files
        originSize := WorkSheetsPathList.Length()
        Loop,% WorkSheetsPathList.Length() - 1
        {
            firstElement := WorkSheetsPathList[A_Index]
            secondElement := WorkSheetsPathList[A_Index+1]
            this.ReorederSheetFile(firstElement, secondElement)
        }

        ; Step 3
        ; touch [ContentType] File
        contentType :=  this.loadXML(this.paths.ContentType)
        XPathString := "//Types/Override[@PartName=""/xl/worksheets/sheet" . originSize . ".xml""]"
        found := contentType.documentElement.selectNodes(XPathString)
        foundItem := found.item(0)
        foundItem.ParentNode.removeChild(foundItem)
        contentType.save(this.paths.ContentType)

        ; Step 4
        ; touch workbook.xml
        workbook := this.loadXML(this.paths.workbook)
        XPathString := "//sheet"
        found := workbook.documentElement.selectNodes(XPathString)
        foundItem := found.item(sheetNo)
        foundItem.ParentNode.removeChild(foundItem)

        found := workbook.documentElement.selectNodes(XPathString)
        Loop,% found.Length()
        {
            found.item(A_Index-1).setAttribute("sheetId", A_Index)
            found.item(A_Index-1).setAttribute("r:id", "rId" . A_Index)
        }
        workbook.save(this.paths.workbook)
        
        ; Check sheet deleted. prevent doing duplicate remove.
        this.isThisSheetDeleted := True
    }

    ReorederSheetFile(FirstFile, SecondFile)
    {
        if FileExist(FirstFile) and FileExist(SecondFile)
            return

        if !FileExist(FirstFile) and FileExist(SecondFile)
        {
            FileMove, % SecondFile, % FirstFile
            return
        }
    }

    

}

; Range Class
class RangeClass extends BaseMethod
{
    ; sheetXML - sheetXML path
    ; sharedStringsXML : sharedStringsXML path
    ; params : Cell Address
    __New(sheetXML, sharedStringsXML, params*)
    {
        if not FileExist(sheetXML)
            throw, "Can't find sheet.xml file."

        if not FileExist(sharedStringsXML)
            throw, "Can't find sharedStrings.xml file."

        this.sheetXML := sheetXML
        this.sharedStringsXML := sharedStringsXML
        this.params := params

        ; if this line. error occurs.
        ; sheet := this.LoadXML(sheetXML)
    }


    value
    {
        get {
            if Not this.sheetData
                throw, "there is no sheetDataDoc."  
            if (this.params.length() = 1) and (this.MultiCellCheck(this.params[1]) = False)
            {
                res := this.FindRange(this.params[1])
                if res
                {
                    return res.text
                }
                else
                {
                    return
                }
            }
            
            ; Msgbox,% this.MultiCellCheck(this.params[1])
            ; Get multi cell values
            if (this.params.length() = 1) and (this.MultiCellCheck(this.params[1]) = True)
            {
                ; it saids B3:E5 format
                addressObject := this.GetCellAddresses(this.params[1])
                res := Array()
                for k, row in addressObject
                {
                    rowArray := Array()
                    for j, cell in row
                    {
                        text := this.FindRange(cell).text
                        rowArray.Push(text)
                    }
                    res.Push(rowArray)
                }
                return res
            }
            if(this.params.length() >= 2 )
            {
                res := Array()
                for k, cell in this.params
                {
                    res.Push(this.FindRange(cell).text)
                }
                return res
            }
        }

        set {
            ; takes assigning value to value var
            if IsObject(value)
            {
                ; if value is object(multiple values)
            }
            else
            {
                this.WriteCell(this.params[1], value)
            }
            
        }
    }

    MultiCellCheck(range)
    {
        StringSplit, splitedRange, range, :
        ; Msgbox,% splitedRange0
        if splitedRange0 > 2
            throw, "Invald Range.`n" . A_ThisFunc
        if splitedRange0 = 2
            return True
        if splitedRange0 = 1
            return False
    }
    
    GetCellAddresses(range)
    {
        ; range
        ; it looks like A3:E8 format
        ; output  > array("A3", "B3", "C4" ----) like this format
        
        ; Split range
        StringSplit, splitedRange, range, :
        if splitedRange0 > 2
            throw, "Invald Range.`n" . A_ThisFunc

        ; Range Check
        rangeColumnNum1 := this.RangeColumnToNumber(splitedRange1)
        rangeColumnNum2 := this.RangeColumnToNumber(splitedRange2)

        if rangeColumnNum1 > rangeColumnNum2
            throw, "Invalid Range.`n" . A_ThisFunc
        
        RegExMatch(splitedRange1, "\d+$", rowNumber1)
        RegExMatch(splitedRange2, "\d+$", rowNumber2)

        if rowNumber1 > rowNumber2
            throw, "Invalid Range.`n" . A_ThisFunc
        
        res := Array()
        
        ; Loop Row
        Loop, % (rowNumber2 - rowNumber1) + 1
        {
            currentRow := A_Index + rowNumber1 -1
            rowArray := Array()
            ; Loop Column
            Loop, % (rangeColumnNum2 - rangeColumnNum1) + 1
            {
                cellAddress := this.NumberToRangeColumn(rangeColumnNum1 + A_Index - 1) . currentRow
                rowArray.Push( cellAddress )
            }
            res.Push(rowArray)
        }
        return res
    }

    NumberToRangeColumn(columnNumber)
    {
        columnName := ""

        while (columnNumber > 0.5) ; i don't know is it ok to use 0.5 float. :)
        {
            modulo := Mod((columnNumber - 1), 26)
            columnName := Chr(65 + modulo) . columnName
            columnNumber := (columnNumber - modulo) / 26
        } 

        return Trim(columnName)
    }

    WriteCell(range, value)
    {
        ns := "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        ns2 := "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
        x14acns := "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
        mcns := "http://schemas.openxmlformats.org/markup-compatibility/2006"
        
        sharedDoc := this.LoadXML(this.sharedStringsXML)

        StringUpper, range, range

        chracterElementCheck := this.FindRange(range, rangeOnly:=True)

        if not this.sheetDataDoc.childNodes[1].getAttribute("xmlns:x14ac")
        {
            this.sheetDataDoc.childNodes[1].setAttribute("xmlns:x14ac", x14acns)
            this.sheetDataDoc.childNodes[1].setAttribute("mc:Ignorable", "x14ac")
            this.sheetDataDoc.childNodes[1].setAttribute("xmlns:mc", mcns)
        }

        ; Find range at the sheetDataDoc
        ; if exist found.
        if chracterElementcheck
        {
            ; it from found element
            chracterElement := chracterElementcheck
            chracterElement.removeAttribute("t")
        }
        else
        {
            ; make new "c" elem
            chracterElement := sharedDoc.createNode(1, "c", ns)
            v := sharedDoc.createNode(1, "v", ns)
            chracterElement.setAttribute("r", range)
            chracterElement.appendChild(v)

        }

        if value is integer
        {
            ; check value whether integer or the other.
            chracterElement.childNodes[0].text := value
        }
        else
        {
            ; when string or other(not checked other type yet.)
            chracterElement.setAttribute("t", "s")

            ; .createNode(Type, name, namespaceURI)
            ; 1 : element
            ; 2 : text
            ; Type : https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms766473(v=vs.85)
            
            si := sharedDoc.createNode(1, "si", ns)
            t := sharedDoc.createNode(1, "t", ns) ; text
            phoneticPr := sharedDoc.createNode(1, "phoneticPr", ns) ; text sibling
            phoneticPr.setAttribute("fontId", "1")
            phoneticPr.setAttribute("type", "noConversion")

            t.text := value
            si.appendChild(t), si.appendChild(phoneticPr)

            sst := sharedDoc.getElementsByTagName("sst")

            ; sst has just one.
            for k, v in sst
                {
                    count := k.getAttribute("count")
                    k.setAttribute("count", count+1)
                    k.appendChild(si)
                }
            
            elemCount := sharedDoc.getElementsByTagName("t").length

            chracterElement.childNodes[0].text := elemCount -1

            if not chracterElementcheck
            {
                ; insert to row
                ; if exist, just put there
                ; else make new row, and adjust rowspan value
                if foundRow := this.FindRow(range)
                {
                    foundRow.appendChild(chracterElement)
                }
                else
                {   
                    ; make row node
                    RegExMatch(range, "\d+$", rowNumber)
                    row := sharedDoc.createNode(1, "row", ns)
                    row.setAttribute("spans", spans)
                    row.setAttribute("r", rowNumber)
                    row.setAttribute("x14ac:dyDescent", 0.3)
                    row.appendChild(chracterElement)

                    ; append row to sheetdata node
                    resTag := this.sheetDataDoc.getElementsByTagName("sheetData")
                    for k, v in resTag
                    {
                        k.appendchild(row)
                    }

                }
            }
            sharedDoc.save(this.sharedStringsXML)
        }
        this.sheetDataDoc.save(this.sheetXML)
    }

    FindRow(rangeAddress)
        {
            RegExMatch(rangeAddress, "\d+$", rowNumber)
            found := this.sheetDataDoc.getElementsByTagName("row")
            
            for k, v in found
            {
                if k.getAttribute("r") = rowNumber
                {
                    return k
                }
            }

            return False
        }

    FindRange(rangeAddress, rangeOnly:=False)
    {
        sheetData := this.sheetData()
        found := sheetData.getElementsByTagName("c")
        for k,v in found
        {
            if k.getAttribute("r") = rangeAddress
            {   
                if rangeOnly
                    return k

                if k.getAttribute("t") = "s"
                {
                    temp := this.SharedStrings()
                    return temp[k.text+1]
                }
                else
                {
                    return k
                }

            }
        }
        
    }
}
    


