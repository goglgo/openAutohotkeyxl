#Include class_sheet.ahk

xl := new OpenAhkXl()
xl.open("aaaa.xlsx")
; xl.GetSheetBySheetNo(1).Range["B3"]
xl.GetSheetBySheetName("TestSheet1")

Msgbox,% xl.Range["B3"]




return





; 새 시트 작성시 바꿔야 할 것
; xl\worksheet\sheet[N].xml 추가(빈 시트 xml 따로 필요)
; xl\workbook.xml
;   SheetName 지정(안 겹치게), sheetID 하나 올려서 추가, r:id 올려서 추가
; docProps\app.xml
;   Vector Size 올리고
;   lpStr 에 추가한 sheetName 추가하고
;   Variant > vt:i4 쪽 숫자도 하나 올림


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

    CheckValidation()
    {
        if not FileExist(this.xlsxPath)
            throw, "There is no .Xlsx File."

        SplitPath, % this.xlsxPath, FileName, FileDir, ,FileNoExt
        If not FileDir
            this.xlsxPath := A_ScriptDir . "\" . this.xlsxPath
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

    ZipXlsx()
    {
        ; it just for save func.

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

    ClearTempFolder()
    {
        ; Clear Temp folder function when exiting app.
        FileRemoveDir, % this.destPath, 1
    }

    IsSheetAlive()
    {
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
            sheetNo := k.getAttribute("sheetId")
            
            this.sheetNameArray[name] := sheetNo
            this.sheetNoArray.Push(sheetNo)
        }
    }

    GetSheetBySheetName(sheetName)
    {
        if not this.sheetNameArray[sheetName]
            throw, "Not initialized. Must open first."

        sheetNo := this.sheetNameArray[sheetName]
        sheetPath := this.paths.workSheetPath . "\sheet" . sheetNo . ".xml"
        sheet := new Sheet(sheetPath, this.paths.sharedStrings)
        this.Range := sheet.Range
    }

    GetSheetBySheetNo(sheetNo)
    {   
        if not this.sheetNoArray[sheetNo]
            throw, "Not initialized. Must open first."
        sheetPath := this.paths.workSheetPath . "\sheet" . sheetNo . ".xml"
        sheet := new Sheet(sheetPath, this.paths.sharedStrings)
        this.Range := sheet.Range
        ; return sheet
    }

    ; xml Paths class
    class PathInfo
    {
        __New(basePath:="")
        {
            this.basePath := basePath

            this.sheetsPath := Array()
            Loop Files, % this.workSheetPath . "\*.*"
            {
                this.sheetsPath.Push(this.workSheetPath . "\" . A_LoopFileName)
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

    }
    

}



GetSheetNames(xml)
{
    result_array := Array()

    doc := LoadXML(xml)
    child := doc.getElementsByTagName("vt:vector").item(1).childNodes

    for v in child
    {
        result_array.push(v.text)
    }

return result_array
}

GetDisplayName(xmldata){

    doc := ComObjCreate("MSXML2.DOMDocument.6.0")
    doc.async := false
    doc.loadXML(xmldata)

    Err := doc.parseError
    if Err.reason
        msgbox % "Error: " Err.reason

    for k, v in tt
    {
        Msgbox,% k.text . "<>" . v
    }

return att_text
}

ZipToTemp(Input_Folder)
{
    Output_Folder := "C:\Temp\aaa\"
    RunWait PowerShell.exe Compress-Archive -Path '%Input_Folder%' -DestinationPath '%Output_Folder%' -Update ,, Hide
}

UnzipToTemp(TargetFile)
{
    SplitPath, TargetFile, FileName, FileDir,,FileNoExt
    DestPath := "C:\Temp\NadureExcel\" . FileNoExt . "\"

    TargetPath := TargetFile
    TargetZipPath := FileDir . "\" . FileNoExt . ".zip"

    FileMove, %TargetPath%, %TargetZipPath%
    RunWait PowerShell.exe -Command Expand-Archive -LiteralPath '%TargetZipPath%' -DestinationPath '%DestPath%',, Hide

    FileMove, %TargetZipPath%, %TargetPath%
    ; PowerShell.exe -NoExit -Command Expand-Archive -LiteralPath 'C:\Users\goglk\Desktop\AutohotkeyXlsx\aaaa.xlsx' -DestinationPath 'C:\Temp\aaa\'
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


findNode(xmlnodes, nodename:="")
{
    for k, v in xmlnodes
    {
        if k.nodeName = nodename
        {
            Msgbox, % k.xml

            return k
        }
        
        
        if k.hasChildNodes()
        {
            result := findNode(k.childNodes, nodename)
            if result
                return result
            ; Msgbox,% k.nodeName . "<>" . nodename
        }
            
    }
    
}


