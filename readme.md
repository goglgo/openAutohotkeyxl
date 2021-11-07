# OpenAhkXl
Automate Ms Excel without com excel.(no need to install ms excel.)

## How it works

The Excel file is just zip file which it just contains xml files.
So this library is just unzip .xlsx file and modifying xml files with msxml com. so this is ms excel free.

<br>

## Advantage
* Faster than com library. a little.(with simple test.)
* No need to install ms office for modifying.

<br>

## Limitation
* Excel has so many detailed features. So i can't deal all of it. 

<br>

## How to use
### 1. Open Xlsx file.

```AutoHotkey
xl := new OpenAhkXl()
xl.open("aaaa.xlsx")
```
<br>

### 2. Get Sheet
* GetSheetBySheetNo
```AutoHotkey
sheet := xl.GetSheetBySheetNo(1)
```
* GetSheetBySheetName
```AutoHotkey
sheet := xl.GetSheetBySheetName(1)
```

<br>

### 3. Treating Sheet

* Add Sheet
```AutoHotkey
xl.addSheet("SheetName")
```
* Delete Sheet(assigned sheet)
```AutoHotkey
sheet.DeleteSheet()
```
<br>


### 3. Treating Range
* Treating Single value.
```AutoHotkey
; Get value from range.
Msgbox,% sheet.Range("B4").value

; Set value to range
sheet.Range("B3").value := "Asdfasd"
```

* Treating Multi values.
```AutoHotkey
; Get values from range.
values := sheet.Range("B2:E3").value

; Set values to range.
sheet.Range("B2:B4").value := "aaa"
sheet.Range("C2:C4").value := [["aa"],["bb"],["cc"]]
```

<br>

### 4. Styling Cell

* Fill

make fill object from Fill function.
```AutoHotkey
fil := Fill()
fil.rgb := "963232"
sheet.Range("C7").style.Fill := fil
```
> It can use only rgb option. now.

<br>

* Font
```
fontt := Font()
fontt.color := "0000000"
fontt.fontSize := 15
fontt.Bold := True
sheet.Range("C8").style.Font := fontt
```
<br>

> Font options.
```AutoHotkey
this.fontName := "" ; set default font when assigning.
this.fontSize := ""
this.color := ""
this.underline := "" ; 1. true, 2. "double", 3. ""
this.Bold := false
this.Italic := false
this.Strike := false
this.Shadow := false
this.Outline := false
```
<br>

* Border

```AutoHotkey
bborder := Border()
bborder.left["style"] := "thin"
bborder.right["style"] := "thick"
bborder.bottom["style"] := "medium"
bborder.top["style"] := "thick"
bborder.top["color"] := "963232" ; set border color
sheet.Range("C9").style.Border := bborder
```
> Border available line style
```AutoHotkey
"thin|thick|medium|dotted"
```

<br>

* Style Combination
```AutoHotkey
sheet.Range("C9").style.Border := bborder
sheet.Range("C9").style.Font := fontt
```
<br>

### 5. Save
save to abstract path or current path.
```
xl.save("Ttt.xlsx")
```
