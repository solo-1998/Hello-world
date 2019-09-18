# Hello-world
this is my first repository

Sub 日报表()

Dim myPath$, 清单表$, 日报表$

Dim 清单wb, 日报wb As Workbook

Dim 清单sh, 日报sh As Worksheet

Dim rng As Range

Dim endNumber As Integer


'------------------------

Application.DisplayAlerts = False

Application.ScreenUpdating = False

myPath = ThisWorkbook.Path & "\"

清单表 = Dir(myPath & "*ganzhou_yd*.xls")

日报表 = Dir(myPath & "*日报表*.xls")

Set 清单wb = Workbooks.Open(myPath & 清单表)

Set 日报wb = Workbooks.Open(myPath & 日报表)

Set 清单sh = 清单wb.Sheets(1)

Set 日报sh = 日报wb.Sheets(1)


endNumber = Application.WorksheetFunction.CountA(清单sh.Range("a:a"))

Dim endStr As String

endStr = "A1:H" & endNumber


Set rng = 清单sh.Range(endStr)



'筛选管理机构

rng.AutoFilter Field:=3, Criteria1:="=863400", Operator:=xlFilterValues


'筛选银代

rng.AutoFilter Field:=4, Criteria1:="=银代", Operator:=xlFilterValues

'筛选有效订单

rng.AutoFilter Field:=5, Criteria1:="=有效", Operator:=xlFilterValues

'筛选直销订单
'rng.AutoFilter Field:=6, Criteria1:="赣州", Operator:=xlFilterValues

rng.Sort key1:="姓名", Order1:=xlAscending, Header:=xlYes

Dim maxCol, maxRow As Long

Dim iCol, iRow As Integer

Dim strCurPerson As String

Dim strCurCost As Double

Dim rest As Range

'-----------新建

清单wb.Sheets.Add After:=ActiveSheet


rng.SpecialCells(xlCellTypeVisible).Copy 清单wb.Sheets(2).Range("a1")

iRow = rng.SpecialCells(xlCellTypeVisible).Cells.Count / 8

'获取最后筛选结果

Set rest = 清单wb.Sheets(2).Range("a1:H" & iRow)

清单wb.Sheets(2).Activate

Dim Dic As Object
Set Dic = CreateObject("Scripting.dictionary")

For i = 2 To iRow
    strCurPerson = rest.Cells(i, 7).Value
    strCurCost = rest.Cells(i, 8).Value
    
    If Not Dic.exists(strCurPerson) Then
    Dic.Add strCurPerson, strCurCost
    Else
        Dic(strCurPerson) = Dic(strCurPerson) + strCurCost
    End If
    
    'If Range("a" & i).EntireRow.Hidden = False Then
        'strCurPerson = rest.Cells(i, 7).Value
    'End If
Next i

清单wb.Sheets.Add After:=ActiveSheet

清单wb.Sheets(3).Activate

清单wb.Sheets(3).Cells(1, 1).Value = "销售人员"
清单wb.Sheets(3).Cells(1, 2).Value = "当日渠道保费"
清单wb.Sheets(3).Cells(1, 3).Value = "当月渠道保费"
Dim er As Double

er = Dic("大壮")
b = Dic.Keys
strCurPerson = b(2)

清单wb.Sheets(3).Cells(2, 1).Resize(UBound(Dic.Keys) + 1, 1) = Application.Transpose(Dic.Keys)
清单wb.Sheets(3).Cells(2, 2).Resize(UBound(Dic.Items) + 1, 1) = Application.Transpose(Dic.Items)

maxCol = rng.CurrentArray.Rows.Count


Workbooks(清单表).Close True

Workbooks(日报表).Close True


Application.DisplayAlerts = Ture

Application.DisplayAlerts = True

End Sub
Sub 人员统计(rng As Range)

'筛选管理机构

rng.AutoFilter Field:=3, Criteria1:="=863400", Operator:=xlFilterValues


'筛选银代

rng.AutoFilter Field:=4, Criteria1:="=银代", Operator:=xlFilterValues

'筛选有效订单

rng.AutoFilter Field:=5, Criteria1:="=有效", Operator:=xlFilterValues
    
    
'rng.AutoFilter Field:=6, Criteria1:="赣州", Operator:=xlFilterValues

rng.Sort key1:="姓名", Order1:=xlAscending, Header:=xlYes

'新建sheet2

清单wb.Sheets.Add After:=ActiveSheet

rng.SpecialCells(xlCellTypeVisible).Copy 清单wb.Sheets(2).Range("a1")

iRow = rng.SpecialCells(xlCellTypeVisible).Cells.Count / 8

'获取最后筛选结果

Set rest = 清单wb.Sheets(2).Range("a1:H" & iRow)

清单wb.Sheets(2).Activate

Dim Dic As Object
Set Dic = CreateObject("Scripting.dictionary")

For i = 2 To iRow
    strCurPerson = rest.Cells(i, 7).Value
    strCurCost = rest.Cells(i, 8).Value
    
    If Not Dic.exists(strCurPerson) Then
    Dic.Add strCurPerson, strCurCost
    Else
        Dic(strCurPerson) = Dic(strCurPerson) + strCurCost
    End If
    
    'If Range("a" & i).EntireRow.Hidden = False Then
        'strCurPerson = rest.Cells(i, 7).Value
    'End If
Next i

清单wb.Sheets.Add After:=ActiveSheet

清单wb.Sheets(3).Activate

清单wb.Sheets(3).Cells(1, 1).Value = "销售人员"
清单wb.Sheets(3).Cells(1, 2).Value = "当日渠道保费"
清单wb.Sheets(3).Cells(1, 3).Value = "当月渠道保费"
Dim er As Double

er = Dic("大壮")
b = Dic.Keys
strCurPerson = b(2)

清单wb.Sheets(3).Cells(2, 1).Resize(UBound(Dic.Keys) + 1, 1) = Application.Transpose(Dic.Keys)
清单wb.Sheets(3).Cells(2, 2).Resize(UBound(Dic.Items) + 1, 1) = Application.Transpose(Dic.Items)


End Sub

Sub 当日直销(rng As Range)



End Sub

Sub 当月直销(rng As Range)



End Sub

Sub 当日渠道(rng As Range)



End Sub

Sub 当月渠道(rng As Range)



End Sub

