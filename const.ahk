global newSheetXMLFormat, newSheetSharedStrings

/*
https://social.msdn.microsoft.com/Forums/office/en-US/e27aaf16-b900-4654-8210-83c5774a179c/xlsx-numfmtid-predefined-id-14-doesnt-match?forum=oxmlsdk
default style number
1 0
2 0.00
3 #,##0
4 #,##0.00
5 $#,##0_);($#,##0)
6 $#,##0_);[Red]($#,##0)
7 $#,##0.00_);($#,##0.00)
8 $#,##0.00_);[Red]($#,##0.00)
9 0%
10 0.00%
11 0.00E+00
12 # ?/?
13 # ??/??
14 m/d/yyyy
15 d-mmm-yy
16 d-mmm
17 mmm-yy
18 h:mm AM/PM
19 h:mm:ss AM/PM
20 h:mm
21 h:mm:ss
22 m/d/yyyy h:mm
37 #,##0_);(#,##0)
38 #,##0_);[Red](#,##0)
39 #,##0.00_);(#,##0.00)
40 #,##0.00_);[Red](#,##0.00)
45 mm:ss
46 [h]:mm:ss
47 mm:ss.0
48 ##0.0E+0
49 @
*/

newSheetXMLFormat =
(
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
    <dimension ref="A1"/>
    <sheetViews>
        <sheetView workbookViewId="0"/>
    </sheetViews>
    <sheetFormatPr defaultRowHeight="16.5" x14ac:dyDescent="0.3"/>
    <sheetData/>
    <phoneticPr fontId="1" type="noConversion"/>
    <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
    <pageSetup paperSize="9" orientation="portrait" r:id="rId1"/>
</worksheet>
)

newSheetSharedStrings = 
(
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
</sst>
)