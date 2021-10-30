import win32com.client
xml = win32com.client.Dispatch("MSXML2.DOMDocument.6.0")
# xml = win32com.client.Dispatch("MSXML2.DOMDocument.3.0")
xml.setProperty("SelectionLanguage", "XPath")
xml.load("sheet1.xml")
# xml.load("styles.xml")

# xml.setProperty("SelectionNamespaces", "xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'")

ns0 = "xmlns:main='http://schemas.openxmlformats.org/spreadsheetml/2006/main'" # test
ns1 = "xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'"
ns2 = 'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"'
ns3 = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
ns4 = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'

ns = f'{ns0} {ns2} {ns3} {ns4}'

# xml.setProperty("SelectionNamespaces" , ns0)
xml.setProperty("SelectionNamespaces" , ns)
xml.documentElement.selectNodes("//main:c")
# xml.setProperty("SelectionNamespaces" , ns1)
# xml.setProperty("SelectionNamespaces" , ns2)
# xml.setProperty("SelectionNamespaces" , ns3)
# xml.setProperty("SelectionNamespaces" , ns4)




# 
xml.load("mstest.xml")
xml.setProperty("SelectionNamespaces", "xmlns:bk='urn:books'")
xml.selectNodes("//Publisher[. = 'MSPress']/parent::node()/Title")
xml.selectNodes("//bk:Publisher[. = 'MSPress']/parent::node()/bk:Title")
# 
# xml.setProperty("SelectionNamespaces" , ns5)
# xml.setProperty("SelectionLanguage", "XPath")

# xml.load("workbook.xml")

# tt = xml.getElementsByTagName("c")
# xml.load("workbook.xml")

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