#NoEnv
;#SendMode Input
SetWorkingDir %A_ScriptDir%
Path := "C:\Users\Amrinder\Downloads\Amrinder FINAL FORM OF First year teachers data(2).xlsx"
indx := 2
!f::
FileSelectFile, Path
return

; !q::
; ex := ComObjCreate("Excel.Application")
; ex.visible := False
; ex.Workbooks.Open(Path)
; Name := ex.Range("C" . indx).Value
; Send %Name%
; return

^q::
Loop, 15
{
ex := ComObjCreate("Excel.Application")
ex.visible := False
ex.Workbooks.Open(Path)

Name1 := ex.Range("A" . indx).Value
Pan1 := ex.Range("B" . indx).Value
Qualification1 := ex.Range("C" . indx).Value
degDate1 := ex.Range("D" . indx).Value
degDate1 := dateFormat(degDate1)
Area1 := ex.Range("E" . indx).Value
Desig1 := ex.Range("F" . indx).Value
joinDate1 := ex.Range("G" . indx).Value
joinDate1 := dateFormat(joinDate1)
CAY1 := Floor(ex.Range("J" . indx).Value)
CAYm11 := Floor(ex.Range("I" . indx).Value)
CAYm21 := Floor(ex.Range("H" . indx).Value)
Assoc1 := ex.Range("K" . indx).Value
Nature1 := ex.Range("L" . indx).Value
leaveDate1 := ex.Range("M" . indx).Value
SendValues(Name1, Pan1, Qualification1, degDate1, Area1, Desig1, joinDate1, CAY1, CAYm11, CAYm21, Assoc1, Nature1, leaveDate1)
indx := indx + 1
}
return

Integer(var)
{
    temp := var*10
    temp := temp/10
    return temp
}

SendValues(Name, Pan, Qualification, degDate, Area, Desig, joinDate, CAY, CAYm1, CAYm2, Assoc, Nature, leaveDate)
{
    Sleep, 500
    ; Click, 62, 706
    clipboard := "document.getElementById(""ctl00_MainContent_rptrQuestion_ctl00_txtFacultyMemberName"").value = """ . Name . """`n"
    clipboard := clipboard . "document.getElementById(""ctl00_MainContent_rptrQuestion_ctl00_txtPAN"").value = """ . Pan . """`n"
    clipboard := clipboard . "document.getElementById(""ctl00_MainContent_rptrQuestion_ctl00_ddlQualification"").value = """ . Qualification . """`n"
    clipboard := clipboard . "document.getElementById(""ctl00_MainContent_rptrQuestion_ctl00_txtDateofReceivingDegree"").value = """ . degDate . """`n"
    clipboard := clipboard . "document.getElementById(""ctl00_MainContent_rptrQuestion_ctl00_txtFMSpecialization"").value = """ . Area . """`n"
    
    clipboard := clipboard . "document.getElementById(""ctl00_MainContent_rptrQuestion_ctl00_ddlFMDesignation"").value = """ . Desig . """`n"
    
    clipboard := clipboard . "document.getElementById(""ctl00_MainContent_rptrQuestion_ctl00_txtFMDateofJoining"").value = """ . joinDate . """`n"
    
    clipboard := clipboard . "document.getElementById(""ctl00_MainContent_rptrQuestion_ctl00_txtcay"").value = """ . CAY . """`n"
    
    clipboard := clipboard . "document.getElementById(""ctl00_MainContent_rptrQuestion_ctl00_txtcay1"").value = """ . CAYm1 . """`n"
    
    clipboard := clipboard . "document.getElementById(""ctl00_MainContent_rptrQuestion_ctl00_txtcay2"").value = """ . CAYm2 . """`n"
    
    clipboard := clipboard . "document.getElementById(""ctl00_MainContent_rptrQuestion_ctl00_ddlCurrentlyAssociated"").value = """ . Assoc . """`n"
    
    clipboard := clipboard . "document.getElementById(""ctl00_MainContent_rptrQuestion_ctl00_ddlAssociation"").value = """ . Nature . """`n"
    clipboard := clipboard . "document.getElementById(""ctl00_MainContent_rptrQuestion_ctl00_txtLeavingDate"").value = """ . leaveDate . """`n"
    clipboard := clipboard . "document.getElementById(""ctl00_MainContent_rptrQuestion_ctl00_btnAddFYFaculty"").click()"
    Send ^v{enter}
    Sleep, 10000
    ; Sleep, 500
    ; Click, 1185, 375
    ; Sleep, 1500
    ; Click, 1177, 331

}
; {
; Send %Name%
; Sleep, 500
; Send {tab}
; Sleep, 500
; Send %Pan%
; Sleep, 500
; Send {tab}
; Sleep, 500
; Send {Down %Qualification%}
; Sleep, 500
; Send {tab}
; Sleep, 500
; Send %degDate%
; Sleep, 500
; Send {tab}%Area%
; Sleep, 500
; Send {tab}
; Sleep, 500
; Send {down 3}
; Sleep, 500
; Send {tab}
; Sleep, 500
; Send %joinDate%
; Sleep, 500
; Send {tab}%CAY%{tab}
; Sleep, 500
; Send %CAYm1%{tab}
; Sleep, 500
; Send %CAYm2%{tab}
; Sleep, 500
; Send {down %Assoc%}
; Sleep, 500
; Send {tab}
; Sleep, 500
; Send {down %Nature%}
; Sleep, 500
; Send {tab 2}
; Sleep, 300
; Send {enter}
; }

dateFormat(var)
{
    tempdate := var
    StringReplace, tempdate, tempdate,-,/, All
    StringReplace, tempdate, tempdate,/1/,/01/, All
    StringReplace, tempdate, tempdate,/2/,/02/, All
    StringReplace, tempdate, tempdate,/3/,/03/, All
    StringReplace, tempdate, tempdate,/4/,/04/, All
    StringReplace, tempdate, tempdate,/5/,/05/, All
    StringReplace, tempdate, tempdate,/6/,/06/, All
    StringReplace, tempdate, tempdate,/7/,/07/, All
    StringReplace, tempdate, tempdate,/8/,/08/, All
    StringReplace, tempdate, tempdate,/9/,/09/, All
    return tempdate
}

; ahk
; clipboard := "document.getElementById(""ctl00_MainContent_rptrQuestion_ctl00_ddlAssociation"").value = ""Regular"""
; something := clipboard
; Send %something%