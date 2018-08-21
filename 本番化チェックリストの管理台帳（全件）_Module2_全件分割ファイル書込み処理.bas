Attribute VB_Name = "全件分割ファイル書込み処理"
Option Explicit

Dim rIdx As Long
Dim rIdy As Long

Sub VgetFILE_all_F()

' ----------------------------------------------------------------------------------------------
'　「全件分割ファイル」シートへの書込み処理処理
'　　「作業用」シートから「全件分割ファイル」シートにデータを貼りつける処理
' ----------------------------------------------------------------------------------------------

rIdx = 1

ThisWorkbook.Worksheets("操作").Range("E6").Value = Now

With ThisWorkbook.Worksheets("全件分割ファイル")
    .Activate
    .Range("A5").Select    '指定の位置にセルを置く
End With

rIdy = ThisWorkbook.Worksheets("全件分割ファイル").Range("A" & Selection.End(xlDown).Row + 1).Row

ThisWorkbook.Worksheets("全件分割ファイル").Range("A1").Value = rIdy

Do While ThisWorkbook.Worksheets("作業用").Cells(rIdx, 1) <> ""

    ThisWorkbook.Worksheets("全件分割ファイル").Range("A" & rIdy).Value = _
                ThisWorkbook.Worksheets("作業用").Cells(rIdx, 1).Value

    ThisWorkbook.Worksheets("全件分割ファイル").Range("B" & rIdy).Value = _
                ThisWorkbook.Worksheets("作業用").Cells(rIdx, 2).Value

    ThisWorkbook.Worksheets("全件分割ファイル").Range("D" & rIdy).Value = _
                ThisWorkbook.Worksheets("作業用").Cells(rIdx, Columns.Count).End(xlToLeft).Value

    ThisWorkbook.Worksheets("全件分割ファイル").Range("C" & rIdy).Value = _
                ThisWorkbook.Worksheets("全件分割ファイル").Range("A" & rIdy).Value & ThisWorkbook.Worksheets("全件分割ファイル").Range("D" & rIdy).Value
    
    rIdx = rIdx + 1
    rIdy = rIdy + 1

Loop

Call VsortSUB_F
    
Call TBL_Del_Add_SQL

ThisWorkbook.Worksheets("操作").Range("F6").Value = Now
MsgBox "終了しました！"

End Sub

Sub VsortSUB_F()
'
    Dim rsort As Long

' ----------------------------------------------------------------------------------------------
'　ソート処理
' ----------------------------------------------------------------------------------------------

    rsort = ThisWorkbook.Worksheets("全件分割ファイル").Range("A" & Selection.End(xlDown).Row).Row
    ThisWorkbook.Worksheets("全件分割ファイル").Range("A2").Value = rsort
    
    With ThisWorkbook.Worksheets("全件分割ファイル").Range("A5:H" & rsort)
        .Sort Key1:=Range("B1"), order1:=xlDescending, _
        Key2:=Range("A1"), order2:=xlDescending, _
        Key3:=Range("D1"), order3:=xlDescending
    End With
    
    ThisWorkbook.Worksheets("全件分割ファイル").Range("A6:H" & rsort).RemoveDuplicates (Array(1, 2, 4))

End Sub



