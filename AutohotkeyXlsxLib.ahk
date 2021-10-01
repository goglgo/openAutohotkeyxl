; Msgbox,% A_ScriptDir . "\ZipTestExtract\"
TargetFile := A_ScriptDir . "\aaaa.xlsx"

; ExcelUnZip(TargetFile) 
; Input_Folder := "C:\Users\goglk\Desktop\AutohotkeyXlsx\_Excel_UnZip\*"
; Output_Folder := "C:\Users\goglk\Desktop\AutohotkeyXlsx\done.zip"

; RunWait PowerShell.exe -NoExit -Command Compress-Archive -LiteralPath '%Input_Folder%' -CompressionLevel Optimal -DestinationPath '%Output_Folder%',, Hide

UnzipToTemp(TargetFile)

; RunWait PowerShell.exe Compress-Archive -Path '%Input_Folder%' -DestinationPath '%Output_Folder%' -Update ,, Hide
; Filemove, done.zip, done.xlsx

; RunWait PowerShell.exe -NoExit -Command Expand-Archive -LiteralPath '%Output_Folder%' -DestinationPath '%DestPath%',, Hide


return


GetSheetNameTest:
GetSheetNames("app.xml")
file.close()
return

class XlsxLib
{
    __New(ExcelFilePath:="")
    {

    }
    UnZipXlsx()
    {

    }
    ZipXlsx()
    {

    }
    LoadXML(xml_path)
    {

    }
}

LoadXML(xml_path)
{
    doc := ComObjCreate("MSXML2.DOMDocument.6.0")
    doc.async := false
    doc.Load(xml_path)


    Err := doc.parseError
    if Err.reason
        {
            msgbox % "Error: "  Err.reason
            ExitApp
        }

    ; doc.setProperty("SelectionLanguage", "XPath")
    ; doc.setProperty("SelectionNamespaces","xmlns:vt='http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'")

    return doc
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
        msgbox % "Error: "  Err.reason


    ; att_text := doc.selectSingleNode("//Properties").getAttributeNode("Application").text
    ; docNode := doc.selectSingleNode("//Properties")
    ; doc.SetProperty("SelectionLanguage","XPath")

    

    for k, v in tt
    {
        Msgbox,% k.text . "<>" . v
    }


    ; working for  find Sheet Name
    ; tt := doc.getElementsByTagName("vt:lpstr")

    ; MsgBox % doc.selectNodes("//Device/DeviceInfos/DeviceInfo").length
    
    ; att_text := DocNode.item.value
    
    ; DocNode := doc.selectSingleNode("//coreProperties/reg")
    ; att_text := DocNode.attributes.getNamedItem("reg.1.displayName").value
    ; doc.selectSingleNode("//Device/DeviceInfos/DeviceInfo[@Name=""test""]").text
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