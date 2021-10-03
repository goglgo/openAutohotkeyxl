; tt := new Sheet("sheet1.xml", "sharedStrings.xml")
; Msgbox,% tt.range["B3"]
; Return

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
        found := this.findNode(doc.childNodes, "sheetData")
        if not found
            throw,"There is no found at the Sheet. getSheetData class."

        return found
    }

    getSharedStrings()
    {
        ; when get value
        ; doc.getElementsByTagName("t")[0].text

        doc := this.LoadXML(this.sharedStringsXML)

        tTags:= doc.getElementsByTagName("t") ; begin 0 (zero)

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

        this.Range.base := this._Range
        this.Range.sheetDataDoc := this.sheetData
        this.Range.SharedStrings := this.SharedStrings
    }

    class _Range
    {
        __Get(rangeAddress)
        {
            if not this.sheetDataDoc
                throw, "there is no sheetDataDoc."
                    . " range() must use after sheet class is being initialized."
            return this.FindRange(rangeAddress)
        }

        __Set(Key, Value)
        {

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







