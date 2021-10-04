tt := new Sheet("sheet1.xml", "sharedStrings.xml")

; MSGbox,% tt.Range("B3").text
tt.Range("B3") := "11111"

Return

returnTrue()
{
    return "True"
}

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
        ; this.Range := new _Range(["b3", "d3"])
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
            ; MSgbox,% value

            if params.length() = 1
            {
                ; um... so complicated.
                ; how about to delete all and recreate all?
                if found := this.FindRange(params[1], rangeOnly:=True)
                {   
                    ; Msgbox,% found.childNodes[0].text
                    ; found.childNodes[0].text := 11111
                    ; Msgbox,% found.childNodes[0].xml
                    ; Msgbox,% this.sheetDataDoc.xml
                    ; this.sheetDataDoc.save(this.sheetXML)

                    if tAttValue := found.getAttribute("t")
                    {
                        if (tAttValue = "t") and (value is integer)
                        {
                            ; remove t attribute and remove element at the sharedStrings.xml
                        }
                        if (tAttValue = "t") and (value is not integer)
                        {
                            ; change sharedStrings value
                        }
                    }

                    else
                    {
                        if value is integer
                        {
                            ; just change sheet value
                        }
                        else
                        {
                            ; change sheet attr 't' to 's' and add value to ...
                        }
                    }

                }

                else if this.FindRow(params[1])
                {
                    ; use existing row elem

                }
                else
                {
                    ; Create new Row Elem
                }

            }

            else{
                ; TODO: make when multiple cells
            }
        }
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

        FindRange(rangeAddress, rangeOnly:=False)
        {
            found := this.sheetData.getElementsByTagName("c")
            for k,v in found
            {
                if rangeOnly
                    return k

                if k.getAttribute("r") = rangeAddress
                {

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







