; // test code for sheet class below.
; tt := new Sheet("sheet1.xml", "sharedStrings.xml")
; Msgbox,% tt.Range("B3").text
; tt.Range("z5") := "tttt"
; Return
; //

; TODO: change architecture.
; testing for Sheet("B3").Range("B3").value

class naduretest
{
    __New(langzz)
    {
        this.lang := langzz
        return lang
    }
    __Get(tt)
    {
        Msgbox,% tt
    }

    __Set(k, v)
    {
        Msgbox,% k . "<>" . v
    }
}

tt := new naduretest("zzzz")
tt.value := "Asdf"
Msgbox,% tt
return


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

        ; this.sheetData := this.getSheetData()
        ; this.SharedStrings := this.getSharedStrings()
    }

    sheetData
    {
        get
        {
            doc := this.LoadXML(this.sheetXML)
            this.sheetDataDoc := doc
            found := this.findNode(doc.childNodes, "sheetData")
            if not found
                throw,"There is no found at the Sheet. please check sheet.xml."
            ; Msgbox,% found.xml
            return found
        }
    }

    SharedStrings
    {
        get
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
                res := this.FindRange(params[1])
                
                this.res.value := res.text
                Msgbox,% res.text
                Msgbox,% res.value
                ; ObjBindMethod()
                return res
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
                    row := doc.createNode(1, "row", ns)
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
            doc.save(this.sharedStringsXML)
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
        found := this.sheetData.getElementsByTagName("c")
        for k,v in found
        {

            if k.getAttribute("r") = rangeAddress
            {   
                if rangeOnly
                    return k

                if k.getAttribute("t") = "s"
                {
                    temp := this.SharedStrings
                    return temp[k.text+1]
                }
                else
                {
                    return k
                }

            }
        }
        
    }

    class RangeClass
    {
        __New()
        {
            return 
        }
        value
        {
            get {
                return this.text
            }
        }
    }

}








