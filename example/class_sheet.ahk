


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
        this._range := Array()
        this.sheetXML := sheetXML
        this.sharedStringsXML := sharedStringsXML

        this.isThisSheetDeleted := False
    }

    Range(params*)
    {
        if this.isThisSheetDeleted
            throw, "This sheet is already deleted."

        if this._range[params[1]]
            return this._range[params[1]]

        rangeClass := new RangeClass(this.sheetXML
            , this.sharedStringsXML, params*)
            
        rangeClass.RangeColumnToNumber := this.RangeColumnToNumber

        ; for assigning style path
        rangeclass.paths := this.paths

        this._range[params[1]] := rangeclass
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
        ; sheetXML - sheet xml path
        ; sharedStringsXMl - sharedStrings xml path
        ; params - range object value. if intput is "b2" then key is params[1]

        if not FileExist(sheetXML)
            throw, "Can't find sheet.xml file."

        if not FileExist(sharedStringsXML)
            throw, "Can't find sharedStrings.xml file."

        this.sheetXML := sheetXML
        this.sharedStringsXML := sharedStringsXML
        this.params := params
        this.isStyle := ""
        

        this.mainns := "http://schemas.openxmlformats.org/spreadsheetml/2006/main" ; main:
        this.x14acns := "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" ; x14ac:
        this.rns := "http://schemas.openxmlformats.org/officeDocument/2006/relationships" ; r:
        this.mcns := "http://schemas.openxmlformats.org/markup-compatibility/2006" ; mc:

        if not this.sheetDataDoc.childNodes[1].getAttribute("xmlns:x14ac")
        {
            this.sheetDataDoc.childNodes[1].setAttribute("xmlns:x14ac", this.x14acns)
            this.sheetDataDoc.childNodes[1].setAttribute("mc:Ignorable", "x14ac")
            this.sheetDataDoc.childNodes[1].setAttribute("xmlns:mc", this.mcns)
        }

    }

    style
    {
        get {
            if this.isStyle
            {
                return this.isStyle
            }
            else
            {
                this.isStyle := new StyleXMLBuildTool(this.paths.style
                    , this.sheetXMLNameSpace
                    , this.GetRangeForStyle(this.params[1])
                    , this.sheetXML)
                this.isStyle.nameSpace := this.sheetXMLNameSpace
                
                return this.isStyle
            }

        }

        ; set {

        ; }
    }

    sheetXMLNameSpace
    {
        get {
            x := "xmlns:"
            nameSpace := format("{1}main='{2}' {1}x14ac='{3}' {1}r='{4}' {1}mc='{5}'"
                , x
                , this.mainNs
                , this.x14acns
                , this.rns
                , this.mcns )
            return nameSpace
        }
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

                if (this.params.length() = 1) and (this.MultiCellCheck(this.params[1]) = True)
                {
                    addressObject := this.GetCellAddresses(this.params[1])
                    for k, row in addressObject
                    {
                        for j, cell in row
                        {
                            this.WriteCell(cell, value[k][j])
                        }
                    }
                }
            }

            else
            {
                ; if value is not object.
                
                ; if not multi cell
                if (this.params.length() = 1) and (this.MultiCellCheck(this.params[1]) = False)
                {
                    this.WriteCell(this.params[1], value)
                }
                    ; this.WriteCell(this.params[1], value)

                ; write whole range with single value
                if (this.params.length() = 1) and (this.MultiCellCheck(this.params[1]) = True)
                {
                    addressObject := this.GetCellAddresses(this.params[1])
                    for k, row in addressObject
                    {
                        for j, cell in row
                        {
                            this.WriteCell(cell, value)
                        }
                    }
                }
                
                ; write certain cells with single value
                if(this.params.length() >= 2 )
                {
                    for k, cell in this.params
                    {
                        this.WriteCell(cell, value)
                    }
                }

                
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

        while (columnNumber > 0.5) ; i don't know whether it is ok for using 0.5 float. :)
        {
            modulo := Mod((columnNumber - 1), 26)
            columnName := Chr(65 + modulo) . columnName
            columnNumber := (columnNumber - modulo) / 26
        } 

        return Trim(columnName)
    }

    GetRangeForStyle(range)
    {
        sheetDoc := this.LoadXML(this.sheetXML)
        sheetDoc.setProperty("SelectionLanguage", "XPath")
        sheetDoc.setProperty("SelectionNamespaces" , this.sheetXMLNameSpace)
        StringUpper, range, range
        
        ; use selectSingleNode method for the performance
        
        foundRange := sheetDoc.DocumentElement.selectSingleNode("//main:c[@r='" . range . "']")
        if foundRange
            return foundRange
        else
        {
            chracterElement := sheetDoc.createNode(1, "c", this.mainns)

            chracterElement.setAttribute("r", range)
            RegExMatch(range, "\d+$", rowNumber)
            rowElem := sheetDoc.DocumentElement.selectSingleNode("//main:row[@r='" . rowNumber . "']")

            if rowElem
            {
                rowElem.appendChild(chracterElement)
            }
            else
            {
                ; make row node
                row := sheetDoc.createNode(1, "row", this.mainns)
                row.setAttribute("spans", "")
                row.setAttribute("r", rowNumber)
                row.setAttribute("x14ac:dyDescent", 0.3)
                row.appendChild(chracterElement)

                ; append row to sheetdata node
                sheetDataElem := sheetDoc.DocumentElement.selectSingleNode("//main:sheetData")
                sheetDataElem.appendChild(row)
            }
            return chracterElement
        }
    }

    WriteCell(range, value)
    {
        ; TODO: for optimizing.
        ; this.sheetXMLNameSpace
        ; sharedDoc := this.LoadXML(this.sharedStringsXML)
        sheetDoc := this.LoadXML(this.sheetXML)
        sheetDoc.setProperty("SelectionLanguage", "XPath")
        sheetDoc.setProperty("SelectionNamespaces" , this.sheetXMLNameSpace)
        StringUpper, range, range
        
        ; use selectSingleNode method for the performance
        foundRange := sheetDoc.DocumentElement.selectSingleNode("//main:c[@r='" . range . "']")
        if value is not integer
        {
            elemCount := this.WriteTextToSharedDoc(value)
        }

        if foundRange ; if Range is
        {
            if value is integer
            {
                foundRange.removeAttribute("t")
                foundRange.selectSingleNode("//main:v").text := value
            }
            else
            {
                ; attribute for string type to "s"
                foundRange.setAttribute("t", "s")
                foundRange.selectSingleNode("main:v").text := elemCount
            }
        }

        else ; if not 
        {
            ; make new character Element
            chracterElement := sheetDoc.createNode(1, "c", this.mainns)
            v := sheetDoc.createNode(1, "v", this.mainns)

            if value is integer
            {
                v.text := value
            }
            else
            {
                chracterElement.setAttribute("t", "s")
                v.text := elemCount
            }
            
            chracterElement.setAttribute("r", range)
            chracterElement.appendChild(v)

            RegExMatch(range, "\d+$", rowNumber)
            rowElem := sheetDoc.DocumentElement.selectSingleNode("//main:row[@r='" . rowNumber . "']")

            if rowElem
            {
                ; make row node
                rowElem.appendChild(chracterElement)
            }
            else
            {
                row := sheetDoc.createNode(1, "row", this.mainns)
                row.setAttribute("spans", "")
                row.setAttribute("r", rowNumber)
                row.setAttribute("x14ac:dyDescent", 0.3)
                row.appendChild(chracterElement)

                ; append row to sheetdata node
                sheetDataElem := sheetDoc.DocumentElement.selectSingleNode("//main:sheetData")
                sheetDataElem.appendChild(row)
            }
        }
        sheetDoc.save(this.sheetXML)

    }

    WriteTextToSharedDoc(value)
    {
        ; return value : SharedDocStringNumber
        sharedDoc := this.LoadXML(this.sharedStringsXML)

        si := sharedDoc.createNode(1, "si", this.mainns)
        t := sharedDoc.createNode(1, "t", this.mainns) ; text
        phoneticPr := sharedDoc.createNode(1, "phoneticPr", this.mainns) ; text sibling
        phoneticPr.setAttribute("fontId", "1")
        phoneticPr.setAttribute("type", "noConversion")

        t.text := value
        si.appendChild(t), si.appendChild(phoneticPr)

        ; sst := sharedDoc.getElementsByTagName("sst")

        sharedDoc.setProperty("SelectionLanguage", "XPath")
        sharedDoc.setProperty("SelectionNamespaces" 
            , Format("xmlns:main='{1}'", this.mainns))

        elemCount := sharedDoc.DocumentElement.selectNodes("//main:si").length

        sst := sharedDoc.DocumentElement.selectSingleNode("//main:sst")
        
        count := sst.getAttribute("count")
        sst.setAttribute("count", count+1)
        sst.appendChild(si)

        sharedDoc.save(this.sharedStringsXML)
        return elemCount
    }

    WriteCell_legacy(range, value)
    {
        ns := "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        ns2 := "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
        x14acns := "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
        mcns := "http://schemas.openxmlformats.org/markup-compatibility/2006"

        
        sharedDoc := this.LoadXML(this.sharedStringsXML)
        
        StringUpper, range, range

        chracterElementCheck := this.FindRange(range, rangeOnly:=True)

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
                    row.setAttribute("spans", "")
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



class StyleXMLBuildTool
{
    __New(stylePath, nameSpace, rangeXml, sheetPath)
    {
        this.stylePath := stylePath
        this.nameSpace := namespace
        this.rangeXml := rangeXML
        this.sheetPath := sheetPath
        this.mainns := "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        this.prevCellXfNo := ""
        
        ; this.nameSpace  - namespace
        xml := ComObjCreate("MSXML2.DOMDocument.6.0")
        xml.async := false
        xml.Load(this.stylePath)

        xml.setProperty("SelectionLanguage", "XPath")
        xml.setProperty("SelectionNamespaces" , this.nameSpace)
        
        Err := xml.parseError
        if Err.reason
        {
            msgbox % "Error: " Err.reason . "`n: " . this.stylePath
            ExitApp
        }

        this.xml := xml

        this.cellXfs := xml.DocumentElement.selectSingleNode("//main:cellXfs")
        this.cellXfsCount := this.cellXfs.getAttribute("count")

        ; numFmtId, fontId, fillId, borderId, xfId
        cloneNode := xml.DocumentElement.selectSingleNode("//main:cellXfs/main:xf[1]").cloneNode(true)
        this.CellXf := this.cellXfs.appendChild(cloneNode)
    }
    
    loadStyleXML()
    {
        xml := ComObjCreate("MSXML2.DOMDocument.6.0")
        xml.async := false
        xml.Load(this.stylePath)

        xml.setProperty("SelectionLanguage", "XPath")
        xml.setProperty("SelectionNamespaces" , this.nameSpace)
        
        Err := xml.parseError
        if Err.reason
        {
            msgbox % "Error: " Err.reason . "`n: " . this.stylePath
            ExitApp
        }

        return xml
    }

    ReloadXMLAndElement()
    {
        if this.prevCellXfNo
        {
            this.xml := this.loadStyleXML()
            this.cellXfs := this.xml.DocumentElement.selectSingleNode("//main:cellXfs")
            this.cellXf := this.cellXfs.childNodes.item(this.prevCellXfNo)
            this.cellXfsCount := this.prevCellXfNo
        }
    }
    
    Save()
    {
        if not this.prevCellXfNo
        {
            this.rangeXml.ownerDocument.save(this.sheetPath)
            this.cellXfs.setAttribute("count", this.cellXfsCount+1)
            this.prevCellXfNo := this.cellXfsCount
            this.xml.save(this.stylePath)
        }
        else
        {
            this.xml.save(this.stylePath)
        }
    }

    ChangeXfAttribute(AttributeName, value)
    {
        if value = 0
        {
            this.CellXf.removeAttribute("apply" . AttributeName)
        }
        else
        {
            this.CellXf.setAttribute("apply" . AttributeName, 1)
            if not this.prevCellXfNo
                this.rangeXml.setAttribute("s", this.cellXfsCount)
        }
        this.CellXf.setAttribute(AttributeName . "Id", value)
        this.Save()
    }

    _CreateElement(nodeName)
    {
        if nodeName = ""
            throw, A_ThisFunc . "`nnode name is null."
        return this.xml.createNode(1, nodeName, this.mainns)
    }

    ; just for viewing
    _SetAttribute(node, key, value)
    {
        node.setAttribute(key, value)
        return node
    }

    ; just for viewing
    GetAttribute(node)
    {
        node.getAttribute(key)
        return node
    }

    Fill
    {
        set
        {
            if value = ""
            {
                this.ChangeXfAttribute("fill", 0)
            }
            else if value.__class = "FillStyleBuild"
            {
                this.ReloadXMLAndElement()
                ; rgb set only
                fills := this.xml.DocumentElement.selectSingleNode("//main:fills")
                fillsCount := fills.getAttribute("count")
                
                fill := this._CreateElement("fill")
                patternFill := this._CreateElement("patternFill")
                patternFill.setAttribute("patternType", "solid")

                fgColor := this._CreateElement("fgColor")
                fgColor.setAttribute("rgb", value.rgb)

                bgColor := this._CreateElement("bgColor")
                bgColor.setAttribute("indexed", 64)

                patternFill.appendChild(fgColor)
                patternFill.appendChild(bgColor)

                fill.appendChild(patternFill)
                fills.appendChild(fill)
                fills.setAttribute("count", fillsCount + 1)

                ; fills.ownerDocument.save(this.stylePath)
                this.xml.save(this.stylePath)

                this.ChangeXfAttribute("fill", fillsCount)
            }
            else
            {
                throw, "please use Fill function. for building style"
            }
        }
    }

    Font
    {
        set
        {
            if value = ""
            {
                this.ChangeXfAttribute("font", 0)
            }
            else if value.__class = "FontStyleBuild"
            {
                this.ReloadXMLAndElement()
                ; rgb set only
                fonts := this.xml.DocumentElement.selectSingleNode("//main:fonts")
                fontsCount := fonts.getAttribute("count")
                
                font := this.xml.DocumentElement.selectSingleNode("//main:font[1]").cloneNode(true)
                ; font sub nodes below
                
                if value.fontSize
                {
                    sz := font.selectSingleNode("main:sz")
                    sz.setAttribute("val", value.fontSize)
                }

                if value.fontName
                {
                    fontName := font.selectSingleNode("main:name")
                    fontName.setAttribute("val", value.fontName)
                }

                if value.Bold
                {
                    Bold := this._CreateElement("b")
                    font.appendChild(Bold)
                }

                if value.strike
                {
                    strike := this._CreateElement("strike")
                    font.appendChild(strike)
                }

                if value.underline
                {
                    underline := this._CreateElement("u")
                    if value.underline = "double"
                        underline.setAttribute("val", "double")
                    font.appendChild(underline)
                }

                color := font.selectSingleNode("main:color")
                if value.color
                {
                    color.removeAttribute("theme")
                    color.setAttribute("rgb", value.color)
                }
                Else
                {
                    if value.color = 0
                    {
                        color.removeAttribute("theme")
                        color.setAttribute("rgb", "000000")
                    }
                }

                fonts.appendChild(font)
                fonts.setAttribute("count", fontsCount + 1)

                ; fonts.ownerDocument.save(this.stylePath)
                this.xml.save(this.stylePath)
                this.ChangeXfAttribute("font", fontsCount)
            }
            else
            {
                throw, "please use Font function. for building style"
            }

        }
    }

    Border {
        set
        {
            if value = ""
            {
                this.ChangeXfAttribute("border", 0)
            }
            else if value.__class = "BorderStyleBuild"
            {
                this.ReloadXMLAndElement()
                borders := this.xml.DocumentElement.selectSingleNode("//main:borders")
                bordersCount := borders.getAttribute("count")
                
                border := this.xml.DocumentElement.selectSingleNode("//main:border[1]").cloneNode(true)
                value.StyleCheck()
                if value.left["style"]
                {
                    left := border.selectSingleNode("main:left")
                    left.setAttribute("style", value.left["style"])

                    color := this._CreateElement("color")
                    if not value.left["color"]
                    {
                        color.setAttribute("indexed", 64)
                    }
                    else
                    {
                        color.setAttribute("rgb", value.left["color"])
                    }
                    left.appendChild(color)
                }

                if value.right["style"]
                {
                    right := border.selectSingleNode("main:right")
                    right.setAttribute("style", value.right["style"])

                    color := this._CreateElement("color")
                    if not value.right["color"]
                    {
                        color.setAttribute("indexed", 64)
                    }
                    else
                    {
                        color.setAttribute("rgb", value.right["color"])
                    }
                    right.appendChild(color)
                }

                if value.top["style"]
                {
                    top := border.selectSingleNode("main:top")
                    top.setAttribute("style", value.top["style"])

                    color := this._CreateElement("color")
                    if not value.top["color"]
                    {
                        color.setAttribute("indexed", 64)
                    }
                    else
                    {
                        color.setAttribute("rgb", value.top["color"])
                    }
                    top.appendChild(color)
                }

                if value.bottom["style"]
                {
                    bottom := border.selectSingleNode("main:bottom")
                    bottom.setAttribute("style", value.bottom["style"])

                    color := this._CreateElement("color")
                    if not value.bottom["color"]
                    {
                        color.setAttribute("indexed", 64)
                    }
                    else
                    {
                        color.setAttribute("rgb", value.bottom["color"])
                    }
                    bottom.appendChild(color)
                }

                borders.appendChild(border)
                borders.setAttribute("count", bordersCount + 1)

                ; borders.ownerDocument.save(this.stylePath)
                this.xml.save(this.stylePath)
                this.ChangeXfAttribute("border", bordersCount)

                
            }
        }
    }

   
}
    

Font()
{
    return new FontStyleBuild()
}

Fill()
{
    
    return new FillStyleBuild()
}

Border()
{
    return new BorderStyleBuild()
}

class FontStyleBuild
{
    __New()
    {
        this.fontName := "" ; set default font when assigning.
        this.fontSize := ""
        this.color := ""
        this.family := ""
        this.underline := "" ; 1. true, 2. "double", 3. ""
        this.Bold := false
        this.Italic := false
        this.Strike := false
        this.Shadow := false
        this.Outline := false

        this.sz := this.fontSize
        this.name := this.fontName
        this.b := this.Bold
    }
}

class BorderStyleBuild
{
    __New()
    {
        ; style - thin, thick, medium, dotted
        ; color - indexed=64(black default) or rgb=000000
        this.availableStyle := "thin|thick|medium|dotted"

        this.left := Array(), this.right := Array()
        this.top := Array(), this.bottom := Array()
        
        this.left["style"] := "", this.right["style"] := ""
        this.top["style"] := "", this.bottom["style"] := ""

        this.left["color"] := "", this.right["color"] := ""
        this.top["color"] := "", this.bottom["color"] := ""

    }

    StyleCheck()
    {   
        if this.left["style"]
        {
            if not InStr(this.availableStyle, this.left["style"])
                throw, "you pushed invalid "
                . "Border Style. the style is : " . this.left["style"]
        }
        
        if this.right["style"]
        {
            if not InStr(this.availableStyle, this.right["style"])
                throw, "you pushed invalid "
                . "Border Style. the style is : " . this.right["style"]
        }
        
        if this.top["style"]
        {
            if not InStr(this.availableStyle, this.top["style"])
                throw, "you pushed invalid "
                . "Border Style. the style is : " . this.top["style"]
        }

        if this.bottom["style"]
        {
            if not InStr(this.availableStyle, this.bottom["style"])
                throw, "you pushed invalid "
                . "Border Style. the style is : " . this.bottom["style"]
        }
        
    }
}

class FillStyleBuild
{
    __New()
    {
        ; Set only Solid Type..
        ; simple color only..
        this.rgb := ""
        this.bgCoilorIndexed := 64

    }

    

}


/*
Fill
    {
        get {
            fillId := this.CellXf.getAttribute("fillId")
            fillXML := this.xml.DocumentElement.selectSingleNode("//main:fill[" . fillId + 1 . "]")
            patternFill := fillXML.childNodes[0]
            patternType_ := patternFill.getAttributeNode("patternType")
            
            fill := fill()

            fill.patternType := patternType
            
            if patternFill.hasChildNodes()
            {
                if fgColorRGB := patternFill.childNodes[0].getAttributeNode("rgb")
                    fill.fgColorRGB_ := fgColorRGB

                if fgColorTheme := patternFill.childNodes[0].getAttributeNode("theme")
                    fill.fgColorThemes_ := fgColorTheme
                
                if fgColortInt := patternFill.childNodes[0].getAttributeNode("tint")
                    fill.fgColortInts_ := fgColortInt

                if bgcolorIndexed := patternFill.childNodes[1].getAttributeNode("indexed")
                    fill.bgcolorIndexeds_ := bgcolorIndexed
            }
            fill.xml := this.xml
            return fill

        }

*/


/*
; how to returning node names.
; for k, v in CellXf.attributes
; {
;     ; Msgbox,% k.nodeName
;     this[k.nodeName] := k.text
; }
; Msgbox,% this.fontId.nodeName
*/