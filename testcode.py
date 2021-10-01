import win32com.client 
xml = win32com.client.Dispatch("MSXML2.DOMDocument.6.0")
import openpyxl as op

# f = open("app.xml","r",encoding="UTF-8")
# data = f.read()
# f.close()

xml.load("app.xml")

# xml.selectNodes("//Properties/TitlesOfParts/vt:vector/vt:lpstr[1]")
xml.setProperty("SelectionLanguage", "XPath")
xml.setProperty("SelectionNamespaces","xmlns:vt='http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'")
xml.getElementsByTagName("vt:vector")[1].childNodes.length
# //Properties/TitlesOfParts/vt:vector/vt:lpstr