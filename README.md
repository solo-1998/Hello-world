# Hello-world
this is my first repository for my baby

Public 清单wb As Workbook
Public PersonSum As Integer
Public rng As Range


Sub 日报表()

Dim myPath$, 清单表$

Dim endNumber As Integer

Dim endStr As String

Application.DisplayAlerts = False

Application.ScreenUpdating = False

myPath = ThisWorkbook.Path & "\"

清单表 = Dir(myPath & "*ganzhou_yd*.xls")

Set 清单wb = Workbooks.Open(myPath & 清单表)

endNumber = Application.WorksheetFunction.CountA(清单wb.Sheets(1).Range("a:a"))

endStr = "A1:H" & endNumber

Set rng = 清单wb.Sheets(1).Range(endStr)

Call 初始统计

Call 当日直销

Call 当月直销

Call 当日渠道

Call 当月渠道

Workbooks(清单表).Close True

Application.DisplayAlerts = Ture

Application.ScreenUpdating = True

End Sub

Sub 筛选统计(Index As Integer)

'Index表示sheet3要修改的列

Dim Dic As Object
Dim rest As Range
Dim Person As String
Dim Cost As Double

If 清单wb.Sheets(2) Is Nothing Then
   MsgBox "Sheet2表不存在"
End If

清单wb.Sheets(2).UsedRange.ClearContents

rng.SpecialCells(xlCellTypeVisible).Copy 清单wb.Sheets(2).Range("a1")

iRow = rng.SpecialCells(xlCellTypeVisible).Cells.Count / 8

Set rest = 清单wb.Sheets(2).Range("a1:H" & iRow)

Set Dic = CreateObject("Scripting.dictionary")

For i = 2 To iRow
    strCurPerson = rest.Cells(i, 7).Value
    strCurCost = rest.Cells(i, 8).Value
    
    If Not Dic.exists(strCurPerson) Then
        Dic.Add strCurPerson, strCurCost
    Else
        Dic(strCurPerson) = Dic(strCurPerson) + strCurCost
    End If
Next i

If 清单wb.Sheets(3) Is Nothing Then
   MsgBox "Sheet3表不存在"
End If

d_keys = Dic.Keys
d_items = Dic.Items
For i = 0 To UBound(Dic.Keys)
    Person = d_keys(i)
    Cost = d_items(i)
    For j = 0 To PersonSum
        If 清单wb.Sheets(3).Cells(j + 2, 1).Value = Person Then
            清单wb.Sheets(3).Cells(j + 2, Index).Value = Cost
        End If
    Next j
Next i


End Sub

Sub 初始统计()

Dim Dic As Object
Dim rest As Range

'筛选管理机构

rng.AutoFilter Field:=3, Criteria1:="=863400", Operator:=xlFilterValues


'筛选银代

rng.AutoFilter Field:=4, Criteria1:="=银代", Operator:=xlFilterValues

'筛选有效订单

rng.AutoFilter Field:=5, Criteria1:="=有效", Operator:=xlFilterValues
    
    
'rng.AutoFilter Field:=6, Criteria1:="赣州", Operator:=xlFilterValues

rng.Sort key1:="姓名", Order1:=xlAscending, Header:=xlYes

'新建sheet2

清单wb.Sheets.Add After:=清单wb.Sheets(1)

rng.SpecialCells(xlCellTypeVisible).Copy 清单wb.Sheets(2).Range("a1")

iRow = rng.SpecialCells(xlCellTypeVisible).Cells.Count / 8

'获取最后筛选结果

Set rest = 清单wb.Sheets(2).Range("a1:H" & iRow)

清单wb.Sheets(2).Activate

Set Dic = CreateObject("Scripting.dictionary")

For i = 2 To iRow
    strCurPerson = rest.Cells(i, 7).Value
    strCurCost = 0
    
    If Not Dic.exists(strCurPerson) Then
        Dic.Add strCurPerson, strCurCost
    Else
        Dic(strCurPerson) = 0
    End If
Next i

清单wb.Sheets.Add After:=清单wb.Sheets(2)

清单wb.Sheets(3).Activate

清单wb.Sheets(3).Cells(1, 1).Value = "销售人员"
清单wb.Sheets(3).Cells(1, 2).Value = "当日直销保费"
清单wb.Sheets(3).Cells(1, 3).Value = "当月直销保费"
清单wb.Sheets(3).Cells(1, 4).Value = "当日渠道保费"
清单wb.Sheets(3).Cells(1, 5).Value = "当月渠道保费"

清单wb.Sheets(3).Cells(2, 1).Resize(UBound(Dic.Keys) + 1, 1) = Application.Transpose(Dic.Keys)
清单wb.Sheets(3).Cells(2, 2).Resize(UBound(Dic.Keys) + 1, 1) = Application.Transpose(Dic.Items)
清单wb.Sheets(3).Cells(2, 3).Resize(UBound(Dic.Keys) + 1, 1) = Application.Transpose(Dic.Items)
清单wb.Sheets(3).Cells(2, 4).Resize(UBound(Dic.Keys) + 1, 1) = Application.Transpose(Dic.Items)
清单wb.Sheets(3).Cells(2, 5).Resize(UBound(Dic.Keys) + 1, 1) = Application.Transpose(Dic.Items)

PersonSum = UBound(Dic.Keys) + 1

End Sub

Sub 当日直销()


rng.AutoFilter Field:=6, Criteria1:="=", Operator:=xlFilterValues

rng.AutoFilter Field:=1, Operator:=xlFilterDynamic, Criteria1:=xlFilterToday

Call 筛选统计(2)

rng.AutoFilter Field:=1

rng.AutoFilter Field:=6

End Sub

Sub 当月直销()

rng.AutoFilter Field:=6, Criteria1:="=", Operator:=xlFilterValues

rng.AutoFilter Field:=1, Criteria1:=xlFilterThisMonth, Operator:=xlFilterDynamic

Call 筛选统计(3)

rng.AutoFilter Field:=1

rng.AutoFilter Field:=6

End Sub

Sub 当日渠道()

rng.AutoFilter Field:=6, Criteria1:="=赣州", Operator:=xlFilterValues

rng.AutoFilter Field:=1, Operator:=xlFilterDynamic, Criteria1:=xlFilterToday

Call 筛选统计(4)

rng.AutoFilter Field:=1

rng.AutoFilter Field:=6


End Sub

Sub 当月渠道()

rng.AutoFilter Field:=6, Criteria1:="=赣州", Operator:=xlFilterValues

rng.AutoFilter Field:=1, Criteria1:=xlFilterThisMonth, Operator:=xlFilterDynamic
Call 筛选统计(5)

rng.AutoFilter Field:=1

rng.AutoFilter Field:=6

End Sub


