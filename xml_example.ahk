xml = 
(
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<!-- 7600 SoundPointIP550 -->
<!-- -->
<polycomConfig xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="polycomConfig.xsd">
<reg
reg.1.displayName="somebody"
reg.1.address="7600-1"
reg.1.label="-7600"
reg.1.type="private"
reg.1.auth.userId="7600-1"
reg.1.auth.password="7600-1"
reg.1.lineKeys="1"
reg.1.callsPerLineKey="4"
/>
</polycomConfig>
)



msgbox % GetDisplayName(xml)
return

GetDisplayName(xmldata){

    doc := ComObjCreate("MSXML2.DOMDocument.6.0")
    doc.async := false
    doc.loadXML(xmldata)
    
    Err := doc.parseError
    if Err.reason
        msgbox % "Error: "  Err.reason
    
    DocNode := doc.selectSingleNode("//polycomConfig/reg")
    att_text := DocNode.attributes.getNamedItem("reg.1.displayName").value
    
    return att_text
}