

ЕСЛИ(Настройки!AK102="";ТЕКСТ(Настройки!$J2;"ч:мм");ТЕКСТ(Настройки!AK102;"ч:мм"))

=ЕСЛИ(Настройки!J2<>"";ТЕКСТ(МИН(Настройки!G36:AK36;Настройки!J2);"ч:мм")&"-"&ТЕКСТ(МАКС(Настройки!G36:AK36;Настройки!J2)+ВРЕМЯ(0;МАКС(Январь!E13:AI13);0);"ч:мм");"")
=ЕСЛИ(Настройки!J2<>"";ТЕКСТ(МИН(Настройки!G42:AI42;Настройки!J2);"ч:мм")&"-"&ТЕКСТ(МАКС(Настройки!G42:AI42;Настройки!J2)+ВРЕМЯ(0;МАКС(Февраль!E13:AG13);0);"ч:мм");"")
=ЕСЛИ(Настройки!J2<>"";ТЕКСТ(МИН(Настройки!G48:AK48;Настройки!J2);"ч:мм")&"-"&ТЕКСТ(МАКС(Настройки!G48:AK48;Настройки!J2)+ВРЕМЯ(0;МАКС(Март!E13:AI13);0);"ч:мм");"")
=ЕСЛИ(Настройки!J2<>"";ТЕКСТ(МИН(Настройки!G54:AJ54;Настройки!J2);"ч:мм")&"-"&ТЕКСТ(МАКС(Настройки!G54:AJ54;Настройки!J2)+ВРЕМЯ(0;МАКС(Апрель!E13:AH13);0);"ч:мм");"")
=ЕСЛИ(Настройки!J2<>"";ТЕКСТ(МИН(Настройки!G60:AK60;Настройки!J2);"ч:мм")&"-"&ТЕКСТ(МАКС(Настройки!G60:AK60;Настройки!J2)+ВРЕМЯ(0;МАКС(Май!E13:AI13);0);"ч:мм");"")
=ЕСЛИ(Настройки!J2<>"";ТЕКСТ(МИН(Настройки!G66:AJ66;Настройки!J2);"ч:мм")&"-"&ТЕКСТ(МАКС(Настройки!G66:AJ66;Настройки!J2)+ВРЕМЯ(0;МАКС(Июнь!E13:AH13);0);"ч:мм");"")
=ЕСЛИ(Настройки!J2<>"";ТЕКСТ(МИН(Настройки!G72:AK72;Настройки!J2);"ч:мм")&"-"&ТЕКСТ(МАКС(Настройки!G72:AK72;Настройки!J2)+ВРЕМЯ(0;МАКС(Июль!E13:AI13);0);"ч:мм");"")
=ЕСЛИ(Настройки!J2<>"";ТЕКСТ(МИН(Настройки!G78:AK78;Настройки!J2);"ч:мм")&"-"&ТЕКСТ(МАКС(Настройки!G78:AK78;Настройки!J2)+ВРЕМЯ(0;МАКС(Август!E13:AI13);0);"ч:мм");"")
=ЕСЛИ(Настройки!J2<>"";ТЕКСТ(МИН(Настройки!G84:AJ84;Настройки!J2);"ч:мм")&"-"&ТЕКСТ(МАКС(Настройки!G84:AJ84;Настройки!J2)+ВРЕМЯ(0;МАКС(Сентябрь!E13:AH13);0);"ч:мм");"")
=ЕСЛИ(Настройки!J2<>"";ТЕКСТ(МИН(Настройки!G90:AK90;Настройки!J2);"ч:мм")&"-"&ТЕКСТ(МАКС(Настройки!G90:AK90;Настройки!J2)+ВРЕМЯ(0;МАКС(Октябрь!E13:AI13);0);"ч:мм");"")
=ЕСЛИ(Настройки!J2<>"";ТЕКСТ(МИН(Настройки!G96:AJ96;Настройки!J2);"ч:мм")&"-"&ТЕКСТ(МАКС(Настройки!G96:AJ96;Настройки!J2)+ВРЕМЯ(0;МАКС(Ноябрь!E13:AH13);0);"ч:мм");"")
=ЕСЛИ(Настройки!J2<>"";ТЕКСТ(МИН(Настройки!G102:AK102;Настройки!J2);"ч:мм")&"-"&ТЕКСТ(МАКС(Настройки!G102:AK102;Настройки!J2)+ВРЕМЯ(0;МАКС(Декабрь!E13:AI13);0);"ч:мм");"")





=ЕСЛИ(Настройки!J2<>"";ТЕКСТ(МИН(Настройки!G36:AK36;Настройки!J2);"ч:мм")&"-"&ТЕКСТ(МАКС(Настройки!G36:AK36;Настройки!J2)+ВРЕМЯ(0;МАКС(Январь!E13:AI13);0);"ч:мм");"")

=СУММПРОИЗВ((Январь[№]=2)*Январь[1];СМЕЩ(Январь[Периодичность];-1;0))
=СУММПРОИЗВ((Январь[№]=3)*Январь[1];СМЕЩ(Январь[Периодичность];-2;0))


СМЕЩ(Январь[Периодичность];-1;0)
СМЕЩ(Январь[Периодичность];-2;0)

=ЕСЛИ([@№]=1;ИНДЕКС(Услуги;ПОИСКПОЗ([@Услуга];Услуги[Кратко];0);4);"")
=ЕСЛИ([@№]=1;ИНДЕКС(Услуги;ПОИСКПОЗ([@Услуга];Услуги[Кратко];0);3);"")
=ЕСЛИ([@№]=1;СУММ(Январь[@[1]:[31]])+СУММ(СМЕЩ(Январь[@[1]:[31]];1;0))+СУММ(СМЕЩ(Январь[@[1]:[31]];2;0));"")
=ЕСЛИ([@УСЛУГ]<>"";[@УСЛУГ]*[@Периодичность];"")


'Call EditTable(ActiveSheet.Name, ontabname, outvalue, st, 1)
'Call AddRowsMonth("January", "January", "Example")


Option Explicit

Public selectSocialIndex As Long

Private Sub getItemCount(control As IRibbonControl, ByRef count)
    Dim socialList As ListObject
    Set socialList = Sheets("Настройки").ListObjects("Услуги")
    Dim scRow As Long
    scRow = socialList.Range.Rows.count - 1
    count = scRow
End Sub

Private Sub getItemImage(control As IRibbonControl, index As Integer, ByRef image)
    image = "MailMergeGoToNextRecord"
End Sub

Private Sub getItemLabel(control As IRibbonControl, index As Integer, ByRef label)
    label = Sheets("Настройки").ListObjects("Услуги").DataBodyRange(index + 1, 2).value
End Sub
Private Sub dropDownClick(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    selectSocialIndex = selectedIndex + 1
End Sub

Public Sub EditTable(sheetName As String, tabname As String, value As String, row As Long, col As Long)
    Dim listobj As ListObject
    Set listobj = Sheets(sheetName).ListObjects(tabname)
    listobj.DataBodyRange(row, col).value = value
End Sub

Public Sub AddRowsSocialList()
    If IsNumeric(EditSource.selectSocialIndex) = True And EditSource.selectSocialIndex > 0 Then
        Application.ScreenUpdating = False
        Dim ontabname As String, outvalue As String
        outvalue = Sheets("Настройки").ListObjects("Услуги").DataBodyRange(EditSource.selectSocialIndex, 2).value
        ontabname = "СписокУслуг"
        Dim SocialTable As ListObject
        Set SocialTable = ActiveSheet.ListObjects(ontabname)
        SocialTable.ListRows.Add
        Dim st As Long
        st = SocialTable.Range.Rows.count - 1
        Call EditTable(ActiveSheet.Name, ontabname, outvalue, st, 1)
        SocialTable.Sort.SortFields.Clear
        Dim rng As Range
        Set rng = Range("СписокУслуг[Список услуг получателя]")
        SocialTable.Sort.SortFields.Add Key:=rng, SortOn:=xlSortOnValues, Order:=xlAscending
        SocialTable.Sort.Header = xlYes
        SocialTable.Sort.Apply
        Application.ScreenUpdating = True
    End If
End Sub

Public Sub AddRowsMonth(theListName As String, theTabName As String, thevalue As String)
    Dim MonthTable As ListObject
    Set MonthTable = Sheets(theListName).ListObjects(theTabName)
    MonthTable.ListRows.Add
    Dim monthrw As Long
    monthrw = MonthTable.Range.Rows.count - 1
    Call EditTable(theListName, theTabName, thevalue, monthrw, 1)
    Call EditTable(theListName, theTabName, 1, monthrw, 4)
    MonthTable.ListRows.Add
    monthrw = MonthTable.Range.Rows.count - 1
    Call EditTable(theListName, theTabName, 2, monthrw, 4)
    MonthTable.ListRows.Add
    monthrw = MonthTable.Range.Rows.count - 1
    Call EditTable(theListName, theTabName, 3, monthrw, 4)
End Sub

Public Sub ClearMonth(onSheetName As String, onMonthName As String)
    Dim onMonthTables  As ListObject
    Set onMonthTables = Sheets(onSheetName).ListObjects(onMonthName)
    Dim tbrow As Long, tbcol As Long
    tbrow = onMonthTables.Range.Rows.count - 1
    tbcol = onMonthTables.Range.Columns.count
    Dim searchRng As Range
    Dim sRngRow As Long
    If tbrow > 3 Then
        Set searchRng = onMonthTables.Range.Find(What:="Услуга", LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, MatchCase:=False, MatchByte:=False)
        If Not searchRng Is Nothing Then
            sRngRow = searchRng.row + 4
            Worksheets(onSheetName).Cells.Range(Worksheets(onSheetName).Cells(sRngRow, 1), Worksheets(onSheetName).Cells(tbrow + 25, tbcol)).Delete Shift:=xlShiftUp
        End If
    End If
    tbrow = onMonthTables.Range.Rows.count - 1
    tbcol = onMonthTables.Range.Columns.count - 2
    Set searchRng = onMonthTables.Range.Find(What:="Услуга", LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, MatchCase:=False, MatchByte:=False)
    If Not searchRng Is Nothing Then
        sRngRow = searchRng.row + 1
        Worksheets(onSheetName).Cells.Range(Worksheets(onSheetName).Cells(sRngRow, 5), Worksheets(onSheetName).Cells(tbrow + 25, tbcol)).Clear
    End If
End Sub

Public Sub JanuaryUpdate()
    Application.ScreenUpdating = False
    
    Dim nameSheet As String, nameMonth As String
    nameSheet = "Январь"
    nameMonth = "Январь"
    
    Dim TabMonth  As ListObject, TabSocial  As ListObject
    Set TabMonth = Sheets(nameSheet).ListObjects(nameMonth)
    Set TabSocial = Sheets("СписокУслуг").ListObjects("СписокУслуг")
    Dim trw As Long, tcol As Long, srw As Long, scol As Long
    trw = TabMonth.Range.Rows.count - 1
    tcol = TabMonth.Range.Columns.count
    srw = TabSocial.Range.Rows.count - 1
    scol = TabSocial.Range.Columns.count
    
    Call ClearMonth(nameSheet, nameMonth)
    
    Dim r As Long
    Call EditTable(nameSheet, nameMonth, TabSocial.DataBodyRange(1, 1).value, 1, 1)
    For r = 2 To srw
        Call AddRowsMonth(nameSheet, nameMonth, TabSocial.DataBodyRange(r, 1).value)
    Next r
    
    Application.ScreenUpdating = True
End Sub

Public Sub CopyToMonth(oldSheet As String, oldMonth As String, currSheet As String, currMonth As String)
    Application.ScreenUpdating = False
    
    Dim oldTables As ListObject, currTables As ListObject
    Set oldTables = Sheets(oldSheet).ListObjects(oldMonth)
    Set currTables = Sheets(currSheet).ListObjects(currMonth)
    Dim orw As Long, ocol As Long, crw As Long, ccol As Long
    orw = oldTables.Range.Rows.count - 1
    ocol = oldTables.Range.Columns.count
    crw = currTables.Range.Rows.count - 1
    ccol = currTables.Range.Columns.count
    
    Call ClearMonth(currSheet, currMonth)
    Call EditTable(currSheet, currMonth, oldTables.DataBodyRange(1, 1).value, 1, 1)
    
    Dim outRange As Range
    Dim outRow As Long
    Set outRange = currTables.Range.Find(What:="Услуга", LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, MatchCase:=False, MatchByte:=False)
    If Not outRange Is Nothing Then
        outRow = outRange.row + 1
        Dim inRande As Range
        Dim inRow As Long
        Set inRande = oldTables.Range.Find(What:="Услуга", LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, MatchCase:=False, MatchByte:=False)
        If Not inRande Is Nothing Then
            inRow = inRande.row + 1
            Worksheets(oldSheet).Cells.Range(Worksheets(oldSheet).Cells(inRow, 1), Worksheets(oldSheet).Cells(orw + 25, 1)).Copy Destination:=Worksheets(currSheet).Cells(outRow, 1)
            Worksheets(oldSheet).Cells.Range(Worksheets(oldSheet).Cells(inRow, 4), Worksheets(oldSheet).Cells(orw + 25, 4)).Copy Destination:=Worksheets(currSheet).Cells(outRow, 4)
        End If
    End If
    
    Application.ScreenUpdating = True
End Sub

Public Sub MonthUpdate()
	Call JanuaryUpdate
	MsgBox "Январь Обновлен"
	Dim elementMonth As String, monthGroup As Variant
	monthGroup = Array("Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь")
	For Each elementMonth In monthGroup
		Call CopyToMonth("Январь", "Январь", elementMonth, elementMonth)
		MsgBox elementMonth & " обновлён!"
	Next elementMonth
End Sub

Call JanuaryUpdate
Call CopyToMonth("Январь", "Январь", "Февраль", "Февраль")











