#Include class_sheet.ahk
#Include const.ahk


xl := new OpenAhkXl()
xl.open("bbbbbbbb.xlsx")

; xl.addSheet("Asdfa")
sheet := xl.GetSheetBySheetNo(1)
; Msgbox
Msgbox,% sheet.Range("B2").value
Msgbox,% sheet.Range("B4").value
; Msgbox,% sheet.Range("B3").value
sheet.Range("B3").value := "Asdfasd"
Msgbox,% sheet.Range("B3").value
xl.save("Ttt.xlsx")


return



; TODO:
; 새 시트 작성시 바꿔야 할 것
; xl\worksheet\sheet[N].xml 추가(빈 시트 xml 따로 필요)
; xl\workbook.xml
;   SheetName 지정(안 겹치게), sheetID 하나 올려서 추가, r:id 올려서 추가
; docProps\app.xml
;   Vector Size 올리고
;   lpStr 에 추가한 sheetName 추가하고
;   Variant > vt:i4 쪽 숫자도 하나 올림

; TODO:
; sharedStrings 정리



class OpenAhkXl
{
    __New()
    {
        ; clear unzipped files in temp folder
        this.Initialize()
        OnExit(ObjBindMethod(this, "ClearTempFolder"))
    }

    Open(ExcelFilePath:="")
    {
        ; Unzip Excel file to Temp folder for treating.
        this.xlsxPath := excelFilePath
        this.CheckValidation()

        SplitPath, % this.xlsxPath, FileName, FileDir, ,FileNoExt
        this.destPath := this.tempFolderBase . FileNoExt . "\"

        this.targetZipPath := FileDir . "\" . FileNoExt . ".zip"
        this.unZipFolderPath := destPath
        this.xmlBase := this.destPath
        this.UnZipXlsx()

        ; Load paths class
        this.paths := new this.PathInfo(this.destPath)
        this.GetSheetInfo()
    }

    addSheet(sheetName)
    {
        ; 새 시트 작성시 바꿔야 할 것
        ; xl\worksheet\sheet[N].xml 추가(빈 시트 xml 따로 필요)
        ; xl\workbook.xml
        ;   SheetName 지정(안 겹치게), sheetID 하나 올려서 추가, r:id 올려서 추가
        ; docProps\app.xml
        ;   Vector Size 올리고
        ;   lpStr 에 추가한 sheetName 추가하고
        ;   Variant > vt:i4 쪽 숫자도 하나 올림
        ; [ContentType].xml
        ;   sheet 추가

        if not sheetName
            throw, "There is no sheet name. it requires."

        ns := ""
        sheetCount := this.Paths.WorkSheetsPathList.Length() + 1

        filePath := this.Paths.workSheetPath . "\sheet" . sheetCount . ".xml"
        FileAppend, %newSheetXMLFormat%, %filePath%

        workBook := this.loadXML(this.Paths.workbook)

        ; check sheetName duplication
        for k, v in workBook.getElementsByTagName("sheet")
        {
            if k.getAttribute("name") = sheetName
            {
                throw, "there is same sheet name."
            }
        }

        ns := "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        nsType := "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
        nsContentType := "http://schemas.openxmlformats.org/package/2006/content-types"

        sheetElement := workBook.createNode(1, "sheet", ns)
        sheetElement.setAttribute("name", sheetName)
        sheetElement.setAttribute("sheetId", sheetCount)
        sheetElement.setAttribute("r:id", "rId" . sheetCount)

        ; sheets has just one.
        for k, v in workBook.getElementsByTagName("sheets")
        {
            k.appendChild(sheetElement)
        }

        workBook.save(this.Paths.workbook)

        app := this.loadXML(this.Paths.app)
        for k, v in app.getElementsByTagName("vt:i4")
        {
            k.text := sheetCount
        }
        vtlpstrElement := app.createNode(1, "vt:lpstr", nsType)
        vtlpstrElement.text := sheetName

        for k, v in app.getElementsByTagName("vt:lpstr")
        {
            if k.parentNode.nodeName = "vt:vector"
            {
                k.parentNode.appendChild(vtlpstrElement)
                k.parentNode.setAttribute("size", sheetCount)
            }
        }
        app.save(this.Paths.app)
        
        contentType := this.LoadXML(this.Paths.ContentType)
        contentTypeAttr := "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
        partName := "/xl/worksheets/sheet" . sheetCount . ".xml"

        overrideElement := contentType.createNode(1, "Override", nsContentType)
        overrideElement.setAttribute("PartName", partName)
        overrideElement.setAttribute("ContentType", contentTypeAttr)

        contentType.childNodes[1].appendChild(overrideElement)
        contentType.save(this.Paths.ContentType)

        this.GetSheetInfo()
    }

    ContentTypeSahredStringsOverrideCheck()
    {
        contentType := this.LoadXML(this.Paths.ContentType)
        nsContentType := "http://schemas.openxmlformats.org/package/2006/content-types"
        contentTypeAttr := "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
        partName := "/xl/sharedStrings.xml"

        for k, v in contentType.getElementsByTagName("Override")
        {
            if(k.getAttribute("PartName") = partName)
            {
                return
            }
        }

        overrideElement := contentType.createNode(1, "Override", nsContentType)
        overrideElement.setAttribute("PartName", partName)
        overrideElement.setAttribute("ContentType", contentTypeAttr)

        contentType.childNodes[1].appendChild(overrideElement)

        contentType.save(this.Paths.ContentType)

    }

    WorkBookRelsRearrange()
    {
        relsTypeNs := "http://schemas.openxmlformats.org/package/2006/relationships"
        workbookRel := this.LoadXML(this.Paths.Workbook_rels)

        idIdx := 0
        for k, v in workbookRel.childNodes[1].childNodes
        {
            k.parentNode.removeChild(workbookRel.childNodes[1].childNodes.item(0))
        }

        relsType := "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
        Loop, Files, % this.Paths.workSheetPath . "\*.xml"
        {
            idIdx += 1
            relElement := this.RelsElementCreator(workbookRel, idIdx, relsType, "worksheets/" . A_LoopFileName, relsTypeNs)
            workbookRel.childNodes[1].appendChild(relElement)
        }

        relsType := "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
        Loop, Files, % this.Paths.theme . "\*.xml"
        {
            idIdx += 1
            relElement := this.RelsElementCreator(workbookRel, idIdx, relsType, "theme/" . A_LoopFileName, relsTypeNs)
            workbookRel.childNodes[1].appendChild(relElement)
        }

        relsType := "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
        idIdx += 1
        relElement := this.RelsElementCreator(workbookRel, idIdx, relsType, "styles.xml", relsTypeNs)
        workbookRel.childNodes[1].appendChild(relElement)
        

        relsType := "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
        idIdx += 1
        relElement := this.RelsElementCreator(workbookRel, idIdx, relsType, "sharedStrings.xml", relsTypeNs)
        workbookRel.childNodes[1].appendChild(relElement)
        workbookRel.save(this.Paths.Workbook_rels)
    }

    RelsElementCreator(doc, idx, relType, Target, ns)
    {
        relsElement := doc.createNode(1, "Relationship", ns)
        relsElement.setAttribute("Id" , "rId" . idx)
        relsElement.setAttribute("Type" , relType)
        relsElement.setAttribute("Target" , Target)
        return relsElement
    }

    CheckValidation()
    {
        if not FileExist(this.xlsxPath)
            throw, "There is no .Xlsx File."

        SplitPath, % this.xlsxPath, FileName, FileDir, ,FileNoExt
        If not FileDir
            this.xlsxPath := A_ScriptDir . "\" . this.xlsxPath

        if this.pidListFromName(FileName).Length()
            throw, FileName . " 파일이 열려있는 중입니다. 확인해주세요."
            
    }

    Initialize()
    {
        this.tempFolderBase := A_Temp . "\NadureExcel\"
    }

    UnZipXlsx()
    {
        FileMove, % this.xlsxPath, % this.targetZipPath
        
        Command := "PowerShell.exe -Command Expand-Archive -LiteralPath '"
            . this.targetZipPath . "' -DestinationPath '" . this.destPath . "'"
                
        RunWait %Command%,, Hide

        FileMove, % this.targetZipPath , % this.xlsxPath

    }

    RearrangeRowSpan()
    {
        ; adjust row span value for all sheet.
        for n, sheetPath in this.paths.WorkSheetsPathList
        {
            ; TODO add if no modified. pass this process

            sheetDoc := this.LoadXML(sheetPath)
            spans := this.RowSpanCheck(sheetDoc)
            
            resRow := sheetDoc.getElementsByTagName("row")

            for row, v in resRow
            {
                row.setAttribute("spans", spans)
            }
            sheetDoc.save(sheetPath)
        }
    }

    save(toSavePath:="")
    {
        this.RearrangeRowSpan()
        this.WorkBookRelsRearrange()
        this.ContentTypeSahredStringsOverrideCheck()

        ; it just for save func.
        if not toSavePath
            toSavePath := this.xlsxPath

        SplitPath, % toSavePath, , FileDir, ,FileNoExt
        SplitPath, % this.xlsxPath, , , ,xlsxFileNoExt

        if not FileNoExt
            FileNoExt := xlsxFileNoExt
        if not FileDir
            FileDir := A_ScriptDir

        toSaveZipPath := FileDir . "\" . FileNoExt . ".zip"

        Command := "PowerShell.exe Compress-Archive -Path "
            . this.destPath . "/* -DestinationPath " . toSaveZipPath . " -Update"

        RunWait %Command%,, Hide
        FileMove, % toSaveZipPath , % toSavePath, 1

    }

    RowSpanCheck(sheetDoc)
    {
        columnNumberArray := Array()
        found := sheetDoc.getElementsByTagName("c")
        res := ""
        for k,v in found
        {
            res := this.RangeColumnToNumber(k.getAttribute("r"))
            columnNumberArray
                .Push(res)
        }
        if columnNumberArray.length() = 1
        {
            ; this.RangeColumnToNumber
            return res . ":" . res
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

    LoadXML(xml_path)
    {
        doc := ComObjCreate("MSXML2.DOMDocument.3.0")
        doc.async := false
        doc.Load(xml_path)

        Err := doc.parseError
        if Err.reason
        {
            msgbox % "Error: " Err.reason 
                . "`n" . A_ThisFunc . "`n" . xml_path
            ExitApp
        }
    return doc
    }

    ClearTempFolder()
    {
        ; Clear Temp folder function when exiting app.
        FileRemoveDir, % this.destPath, 1
    }

    IsSheetAlive()
    {
        ; TODO set. returning is sheet available.
        if not this.paths.workbook
            throw, "the paths are not initialized."

    }

    GetSheetInfo()
    {
        if not this.paths.workbook
            throw, "the paths are not initialized."
        doc := this.LoadXML(this.paths.workbook)
        res := doc.getElementsByTagName("sheet")

        this.sheetNameArray := Array()
        this.sheetNoArray := Array()
        
        for k, v in res
        {
            name := k.getAttribute("name")
            sheetrID := k.getAttribute("r:id")
            RegExMatch(sheetrID, "\d+$", sheetNo)
            
            this.sheetNameArray[name] := sheetNo
            ; this.sheetNoArray.Push(sheetNo)
            this.sheetNoArray["Sheet" . sheetNo] := sheetNo
        }
    }

    GetSheetBySheetName(sheetName)
    {
        if not this.sheetNameArray[sheetName]
            throw, "Not initialized. Must open first."

        sheetNo := this.sheetNameArray[sheetName]
        sheetPath := this.paths.workSheetPath . "\sheet" . sheetNo . ".xml"
        Sheet := new Sheet(sheetPath, this.paths.sharedStrings)
        return Sheet
    }

    GetSheetBySheetNo(sheetNo)
    {   
        if not this.sheetNoArray["Sheet" . sheetNo]
            throw, "Not initialized. Must open first."
        sheetPath := this.paths.workSheetPath . "\sheet" . sheetNo . ".xml"
        sheet := new Sheet(sheetPath, this.paths.sharedStrings)
        return sheet
    }

    ; xml Paths class
    class PathInfo
    {
        __New(basePath:="")
        {
            this.basePath := basePath
            ; newSheetSharedStrings
            if not fileExist(this.sharedStrings)
            {
                FileAppend, %newSheetSharedStrings%, % this.sharedStrings
            }
        }

        app
        {
            get {
                return this.basePath . "\docProps\app.xml"
            }
        }

        core
        {
            get {
                return this.basePath . "\docProps\core.xml"
            }
        }

        sharedStrings
        {
            get {
                return this.basePath . "\xl\sharedStrings.xml"
            }
        }

        workbook
        {
            get {
                return this.basePath . "\xl\workbook.xml"
            }
        }

        workSheetPath
        {
            get {
                return this.basePath . "\xl\worksheets"
            }
        }

        WorkSheetsPathList
        {
            get {
                    pathList := Array()
                    Loop, Files, % this.workSheetPath . "\*.xml"
                    {
                        pathList.Push(A_LoopFileFullPath)
                    }
                    return pathList

                return 
            }
        }

        ContentType
        {
            get {
                return this.basePath . "\[Content_Types].xml"
            }
        }

        Workbook_rels
        {
            get {
                return this.basePath . "\xl\_rels\workbook.xml.rels"
            }
        }

        theme
        {
            get {
                return this.basePath . "\xl\theme"
            }
        }

    }

    pidListFromName(name) {
        static wmi := ComObjGet("winmgmts:\\.\root\cimv2")
        
        if (name == "")
            return

        PIDs := []
        for Process in wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" name "'")
            PIDs.Push(Process.processId)
        return PIDs 
    }
    

}



