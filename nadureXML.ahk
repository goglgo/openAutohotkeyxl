
fileName := "sheet1.xml"
xml := new NadureXML(fileName)
res := xml.findNodes("c", ["r=""B5"""])
MSgbox,% res[1] . "`nres"



return

class NadureXML
{
    __New(fileName)
    {
        this.xml := this._OpenXML(fileName)

    }
    _OpenXML(fileName)
    {
        FileRead, f, %fileName%
        return f
    }
    FindNodes(nodeName, extraOptions*)
    {
        ; https://www.autohotkey.com/board/topic/117001-regexmatch-multiple-matches/
        ; nodeName := "c"
        ; extraOptions := ["r=""B5""", "t=""s"""]
        p := 1, m := "", nodeName := "c"
        regex := "<" . nodeName . "(.+?)</" . nodeName . ">"
        output := Array()
        while p := RegExMatch(this.xml, regex, m, p + StrLen(m))
        {
            if extraOptions
            {
                for k, option in extraOptions
                {
                    if InStr(m1, option[k])
                    {   
                        output.Push("<" . nodeName . m1 . "</" . nodeName . ">")
                    }
                }
            }
            else
            {
                output.Push("<" . nodeName . m1 . "</" . nodeName . ">")
            }

        }
        return output
    }
}