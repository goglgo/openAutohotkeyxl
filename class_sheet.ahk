; tt := new Sheet("sheet1.xml", "sharedStrings.xml")
; Msgbox,% tt.Range("B3").text
; tt.Range("z5") := "tttt"
; Return


class Sheet
{
    __New(sheetXML:="", sharedStringsXML:="")
    {
        if not FileExist(sheetXML)
            throw, "Can't find sheet.xml file."

        if not FileExist(sharedStringsXML)
            throw, "Can't find sharedStrings.xml file."

        this.sheetXML := sheetXML
        this.sharedStringsXML := sharedStringsXML

        this.sheetData := this.getSheetData()
        this.SharedStrings := this.getSharedStrings()
    }

    getSheetData()
    {
        doc := this.LoadXML(this.sheetXML)
        this.sheetDataDoc := doc
        found := this.findNode(doc.childNodes, "sheetData")
        if not found
            throw,"There is no found at the Sheet. At the getSheetData method."

        return found
    }

    getSharedStrings()
    {
        ; when get value

        doc := this.LoadXML(this.sharedStringsXML)
        this.sharedStringsDoc := doc
        tTags:= doc.getElementsByTagName("t") 

        ; it has no __ENum. so rearrange.
        result := Array()
        for k, v in tTags
            result.Push(k)
        return result 
    }

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
            msgbox % "Error: " Err.reason
            ExitApp
        }
    return doc
    }


    Range[params*]
    {
        get
        {
            if Not this.sheetData
                throw, "there is no sheetDataDoc."

            if params.length() = 1
            {
                return this.FindRange(params[1])
            }

            ; TODO: make when multiple cells
        }

        set
        {
            if Not this.sheetData
                throw, "there is no sheetDataDoc."

            ; fixed value var with assigning.
            if params.length() = 1
            {   
                this.WriteCell(params[1], value)
            }

            else {
                ; TODO: make when multiple cells
            }
        }
    }

    WriteCell(range, value)
    {
        ns := "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        ns2 := "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
        doc := this.sharedStringsDoc
        StringUpper, range, range

        chracterElementcheck := this.FindRange(range, rangeOnly:=True)

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
            chracterElement := doc.createNode(1, "c", ns)
            v := doc.createNode(1, "v", ns)
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

            
            si := doc.createNode(1, "si", ns)
            t := doc.createNode(1, "t", ns) ; text
            phoneticPr := doc.createNode(1, "phoneticPr", ns) ; text sibling
            phoneticPr.setAttribute("fontId", "1")
            phoneticPr.setAttribute("type", "noConversion")

            t.text := value
            si.appendChild(t), si.appendChild(phoneticPr)

            sst := doc.getElementsByTagName("sst")

            ; sst has just one.
            for k, v in sst
                {
                    count := k.getAttribute("count")
                    k.setAttribute("count", count+1)
                    k.appendChild(si)
                }
            
            elemCount := doc.getElementsByTagName("t").length
            chracterElement.childNodes[0].text := elemCount -1
            Msgbox,% chracterElement.xml

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
                    spans := this.RowSpanCheck()
                    row := doc.createNode(1, "row", ns)
                    row.setAttribute("spans", spans)
                    row.setAttribute("r", rowNumber)
                    row.setAttribute("x14ac:dyDescent", 0.3)
                    ; row.setAttribute("xmlns:x14ac", ns2)
                    row.appendChild(chracterElement)

                    ; all row change the rowSpan value
                    resRow := doc.getElementsByTagName("row")
                    for k, v in resRow
                    {
                        row.setAttribute("spans", spans)
                    }

                    ; append row to sheetdata node
                    resTag := this.sheetDataDoc.getElementsByTagName("sheetData")
                    for k, v in resTag
                    {
                        k.appendchild(row)
                    }

                }
            }
            doc.save(this.sharedStringsXML)
        }
        this.sheetDataDoc.save(this.sheetXML)
    }

    RowSpanCheck()
    {
        columnNumberArray := Array()
        found := this.sheetData.getElementsByTagName("c")
        for k,v in found
        {
            columnNumberArray
                .Push(this.RangeColumnToNumber(k.getAttribute("r")))
        }
        ; Msgbox,% Min(columnNumberArray*) . ":Min`n" . Max(columnNumberArray*) . ":Max"
        return Min(columnNumberArray*) . ":" . Max(columnNumberArray*)
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
        found := this.sheetData.getElementsByTagName("c")
        for k,v in found
        {

            if k.getAttribute("r") = rangeAddress
            {
                if rangeOnly
                    return k

                if k.getAttribute("t") = "s"
                {
                    ; it saids string. need sharedStrings data.
                    ; Msgbox,% this.SharedStrings[k.text +1]
                    temp := this.SharedStrings[k.text+1]
                    return temp
                }
                else
                {
                    return k
                }

            }
        }
        
    }

    SetRangeLegacy()
    {
        this.Range := Array()
        ; assigin _Range class for extending built-in Array class.
        this.Range.base := this._Range
        
        this.Range.sheetXML := this.sheetXML
        this.Range.SharedStrings := this.SharedStrings
        this.Range.sheetDataDoc := this.sheetData
        this.Range.sharedStringsXML := this.sharedStringsXML
        this.Range.base.__Set := this._RangeMethodClass.__Set

    }

    

    class _RangeLegacy
    {
        __Get(rangeAddress)
        {
            Msgbox, %rangeAddress%
            if Not this.sheetDataDoc
                throw, "there is no sheetDataDoc."
            return this.FindRange(rangeAddress)
        }

        DeleteStringFromSharedStrings(Value)
        {

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

        FindRange(rangeAddress)
        {
            found := this.sheetDataDoc.getElementsByTagName("c")
            for k,v in found
            {
                ; Msgbox,% k.getAttribute("t")
                if k.getAttribute("r") = rangeAddress
                {

                    if k.getAttribute("t") = "s"
                    {
                        ; it saids string. need sharedStrings data.
                        ; Msgbox,% k.text -1
                        ; Msgbox,% this.SharedStrings[k.text -1]
                        return this.SharedStrings[k.text+1].text
                    }
                    else
                    {
                        return k.text
                    }

                }
            }
            
        }
    }

}







