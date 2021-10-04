#Include class_sheet.ahk

xl := new OpenAhkXl()
xl.open("aaaa.xlsx")
sheet := xl.GetSheetBySheetName("TestSheet1")

Msgbox,% sheet.Range("B3").text
sheet.Range("B3") := "asdfadsf"
sheet.Range("C10") := "asdfadsf"
sheet.Range("D11") := "zzzzz"
Msgbox,% sheet.Range("B3").text

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

    save(toSavePath:="")
    {
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

        Clipboard := Command
        RunWait %Command%,, Hide
        FileMove, % toSaveZipPath , % toSavePath, 1

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
        Sheet := new Sheet(sheetPath, this.paths.sharedStrings)
        return Sheet
    }

    GetSheetBySheetNo(sheetNo)
    {   
        if not this.sheetNoArray[sheetNo]
            throw, "Not initialized. Must open first."
        sheetPath := this.paths.workSheetPath . "\sheet" . sheetNo . ".xml"
        sheet := new Sheet(sheetPath, this.paths.sharedStrings)
        ; this.Range := sheet
        return sheet
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



