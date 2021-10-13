

selectNodeTest:
doc := LoadXML("sheet1.xml")

root := doc.documentElement


for nodeItem in ( root.selectNodes( "//c" ),  descList2 := ""  )
    ; descList2 .= nodeItem.getAttribute( "r" ) "|"
    Msgbox,% nodeItem.getAttribute("r")

msgbox % RTrim( descList2, "|" )

return
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

