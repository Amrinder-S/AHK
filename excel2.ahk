#NoEnv
;#SendMode Input
SetWorkingDir %A_ScriptDir%
indx := 2

!f::
FileSelectFile, Path
return

rangeis := "A3"
; !q::
; ex := ComObjCreate("Excel.Application")
; ex.visible := False
; ex.Workbooks.Open(Path)
; Name := ex.Range("C6").Value
; Send %Name%
; return

^q::
ex := ComObjCreate("Excel.Application")
ex.visible := False
ex.Workbooks.Open(Path)

Name1 := ex.Range(rangeis).Value
Pan1 := ex.Range("C3").Value
Qualification1 := 15
degDate1 := ex.Range("E3").Value
Area1 := ex.Range("F3").Value
Desig1 := 2
joinDate1 := ex.Range("H3").Value
CAY1 := Floor(ex.Range("I3").Value)
CAYm11 := Floor(ex.Range("J3").Value)
CAYm21 := Floor(ex.Range("K3").Value)
Assoc1 := Integer("1")
Nature1 := Integer("2")
SendValues(Name1, Pan1, Qualification1, degDate1, Area1, Desig1, joinDate1, CAY1, CAYm11, CAYm21, Assoc1, Nature1)
indx := indx + 1
return

Integer(var)
{
    temp := var*10
    temp := temp/10
    return temp
}

SendValues(Name, Pan, Qualification, degDate, Area, Desig, joinDate, CAY, CAYm1, CAYm2, Assoc, Nature)
{
    Send %Name%
Sleep, 500
Send {tab}
Sleep, 500
Send %Pan%
Sleep, 500
Send {tab}
Sleep, 500
Send {Down %Qualification%}
Sleep, 500
Send {tab}
Sleep, 500
Send %degDate%
Sleep, 500
Send {tab}%Area%
Sleep, 500
Send {tab}
Sleep, 500
Send {down 3}
Sleep, 500
Send {tab}
Sleep, 500
Send {text}%joinDate%
Sleep, 500
Send {tab}%CAY%{tab}
Sleep, 500
Send %CAYm1%{tab}
Sleep, 500
Send %CAYm2%{tab}
Sleep, 500
Send {down %Assoc%}
Sleep, 500
Send {tab}
Sleep, 500
Send {down %Nature%}
Sleep, 500
Send {tab 2}
}