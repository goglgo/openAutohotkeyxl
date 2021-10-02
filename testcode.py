import win32com.client
# xml = win32com.client.Dispatch("MSXML2.DOMDocument.6.0")
xml = win32com.client.Dispatch("MSXML2.DOMDocument.3.0")

# xml.load("sheet1.xml")
xml.load("sharedStrings.xml")
# xml.setProperty("SelectionLanguage", "XPath")

# 자동 로드 만들 수 있을 듯
# xml.DocumentElement.attributes.item(1).nodeName 
# xml.DocumentElement.attributes.item(1).nodeValue


# xml.selectNodes("//Properties/TitlesOfParts/vt:vector/vt:lpstr[1]")

# xml.setProperty("SelectionNamespaces","xmlns:vt='http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'")
# xml.getElementsByTagName("vt:vector")[1].childNodes.length
# //Properties/TitlesOfParts/vt:vector/vt:lpstr



# XmlNamespaces = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"'
# xml.setProperty("SelectionNamespaces", XmlNamespaces)
# xml.setProperty("SelectionNamespaces",'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
# xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"