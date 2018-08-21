Attribute VB_Name = "エラー証憑取得処理"
Option Explicit

Dim Indx1 As Long
Dim Indx2 As Long

Sub 全件分割ファイル_ERR_MAIN_F()

' ----------------------------------------------------------------------------------------------
'　エラー証憑取得処理
'　　「警告」・「要注意」の明細をピックアップし一覧管理する
' ----------------------------------------------------------------------------------------------

'MsgBox "開始します！"

Dim myPath As String
myPath = ThisWorkbook.Path & "\"

Call Err_DgetA_F          'シート「要注意一覧」にエラーを貼り付ける
Call Err_DputD_F          'ファイル「管理台帳_警告・要注意.xlsx」にエラーを貼り付ける

ThisWorkbook.Worksheets("要注意一覧").Range("B1").Value = Now
MsgBox "終了しました！"

End Sub

Sub Err_DgetA_F()

Indx1 = 6
Indx2 = 4

ThisWorkbook.Worksheets("要注意一覧").Range("B5:H2000").ClearContents

Do While ThisWorkbook.Worksheets("全件分割ファイル").Cells(Indx1, 1) <> ""
    
    ThisWorkbook.Worksheets("要注意一覧").Range("A1").Value = Indx1

'   管理台帳から<警告>・<要注意>・<確認>のものを管理台帳にセットする
    If InStr(ThisWorkbook.Worksheets("全件分割ファイル").Cells(Indx1, 7), "<警告>") > 0 Or InStr(ThisWorkbook.Worksheets("全件分割ファイル").Cells(Indx1, 7), "<要注意>") > 0 Or InStr(ThisWorkbook.Worksheets("全件分割ファイル").Cells(Indx1, 7), "<確認>") > 0 Then
    Indx2 = Indx2 + 1
    With Worksheets("要注意一覧")
        .Range("A" & Indx2).Formula = "=ROW()-4"
        .Range("B" & Indx2).Value = Worksheets("全件分割ファイル").Range("A" & Indx1).Value
        .Range("C" & Indx2).Value = Worksheets("全件分割ファイル").Range("B" & Indx1).Value
        .Range("D" & Indx2).Value = Worksheets("全件分割ファイル").Range("D" & Indx1).Value
        .Range("E" & Indx2).Value = Worksheets("全件分割ファイル").Range("E" & Indx1).Value
        .Range("F" & Indx2).Value = Worksheets("全件分割ファイル").Range("F" & Indx1).Value
        .Range("G" & Indx2).Value = Worksheets("全件分割ファイル").Range("G" & Indx1).Value
        .Range("H" & Indx2).Formula = "=vlookup(B" & Indx2 & ",本番化一覧!$A$2:$K$1000,3,FALSE)"
        End With
    Else
    End If

    Indx1 = Indx1 + 1

Loop

End Sub

Sub Err_DputD_F()

Dim myPath As String
myPath = ThisWorkbook.Path & "\"

Dim END_CELL As Integer
Const TempA As String = "管理台帳_警告・要注意.xlsx"

    Workbooks.Open Filename:=myPath & TempA
    END_CELL = Workbooks(TempA).Worksheets("本番化").Cells(4, 1).End(xlDown).Row
    Workbooks(TempA).Worksheets("本番化").Range("A5:E" & END_CELL).ClearContents
    END_CELL = ThisWorkbook.Worksheets("要注意一覧").Cells(4, 1).End(xlDown).Row
    ThisWorkbook.Worksheets("要注意一覧").Range("B5:G" & END_CELL).Copy

    Workbooks(TempA).Worksheets("本番化").Range("B5").PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False

    Workbooks(TempA).Worksheets("本番化").Range("B1").Value = Now

    Workbooks(TempA).Save
    Workbooks(TempA).Close SaveChanges:=False, Filename:=TempA


End Sub












