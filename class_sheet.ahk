tt := new Sheet("sheet1.xml", "sharedStrings.xml")
; Msgbox,% tt.range["B3"]

; tt.test()


Return

class Sheet
{
    __New(sheetXML:="", sharedStringsXML:="")
    {
        this.sheetXML := sheetXML
        this.sharedStringsXML := sharedStringsXML

        this.sheetData := this.getSheetData()
        this.SharedStrings := this.getSharedStrings()
        this.SetRange()
    }

    test()
    {
        Msgbox,% this.sharedStringsXML.getElementsByTagName("si").length
    }

    getSheetData()
    {
        doc := this.LoadXML(this.sheetXML)
        found := this.findNode(doc.childNodes, "sheetData")
        if not found
            throw,"There is no found at the Sheet. At the getSheetData method."

        return found
    }

    getSharedStrings()
    {
        ; when get value
        ; doc.getElementsByTagName("t")[0].text

        doc := this.LoadXML(this.sharedStringsXML)
        ; this.sharedStringsXML := doc
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

    SetRange()
    {
        this.Range := Array()
        ; assigin _Range class for extending built-in Array class.
        this.Range.base := this._Range
        
        this.Range.sheetXML := this.sheetXML
        ; this.Range.SharedStrings := this.SharedStrings
        ; this.Range.sheetDataDoc := this.sheetData
        ; this.Range.sharedStringsXML := this.sharedStringsXML

    }

    class _Range
    {
        __Get(rangeAddress)
        {
            Msgbox, %rangeAddress%
            if Not this.sheetDataDoc
                throw, "there is no sheetDataDoc."
            return this.FindRange(rangeAddress)
        }

        __Set(rangeAddress, Value)
        {

            ; TODO:object 세트시, recursion하는 문제가 발생함. 에러가 뜨는데, __Set을 마지막에 obj combine으로 묶어주면 어떨까
            
            ; If IsObject(Value)
            ; {
            ;     Msgbox,111111111111111
            ; }
            ; if this.FindRow(rageAddress)
            ; {
            ;     if Value is integer
            ;     {

            ;     }
            ;     else
            ;     {

            ;     }
            ; }
            ; else
            ; {

            ; }
        }

        DeleteStringFromSharedStrings(Value)
        {

        }

        FindRow(rangeAddress)
        {
            RegExMatch(test, "\d+$", rowNumber)
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







