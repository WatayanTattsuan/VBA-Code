Attribute VB_Name = "Module1"
Option Explicit

Const myPath As String = "D:\SVN\管理台帳\管理台帳_2018\"
Dim Skip_Count As Long
Dim Indx1 As Long

Sub 管理台帳_PRISM_MAIN_F()
'---------------------------------------------------------------------------------------
'管理台帳に関係する情報の取込み処理
'
'---------------------------------------------------------------------------------------

MsgBox "開始します！"

ThisWorkbook.Worksheets("LOG").Range("B2").Value = Now
Call DgetB_F          '各管理台帳から状況の取込
Call DgetA_F          'チェック
Call DputD_F          'TEMPファイルに反映
Call TBL_Del_Add_SQL
ThisWorkbook.Worksheets("LOG").Range("C2").Value = Now

ThisWorkbook.Worksheets("管理台帳_PRISM").Cells(2, 6).Value = Now

Skip_Count = Indx1 - 6 - Skip_Count
MsgBox "終了しました！" & vbLf & "スキップカウント　：　" & Skip_Count

End Sub



Sub DgetA_F()
'
' Macro_check1 Macro
'

Dim IraiNO As Long
Dim Irai1 As Long
Dim Irai2 As Long
Dim Shub1 As Long
Dim UkeNo As Long
Dim Test1 As Long
Dim Test2 As Long
Dim Test3 As Long
Dim Atta1 As Long
Dim Atta2 As Long
Dim Atta3 As Long
Dim Atta4 As Long
Dim Atta5 As Long
Dim Atta6 As Long
Dim HokoNO As Long
Dim Hoko1 As Long
Dim Hoko2 As Long
Dim Hoko3 As Long
Dim Hoko4 As Long
Dim ShUkNo As Long
Dim KnUkNo As Long

Dim Flag1 As Long

Shub1 = 13     '列名：発生種別
UkeNo = 17     '列名：受付No&連番
ShUkNo = 40    '集約用受付No
KnUkNo = 42    '管理用受付No
IraiNO = 45    '依頼書文書番号2
HokoNO = 46    '報告書文書番号2
Hoko3 = 67     '列名：報告書検収承認
Hoko4 = 68     '列名：報告書検収承認者
Indx1 = 6
Skip_Count = 0

ThisWorkbook.Sheets("管理台帳_PRISM").Activate
ActiveSheet.Range("BK6").Select

If ActiveSheet.Range("A1").Value < 6 Then
    MsgBox "値が間違っています"
Else: Indx1 = ActiveSheet.Range("A1").Value
End If

Do While ActiveSheet.Cells(Indx1, 37) <> "-" And ActiveSheet.Range("A2").Value > Indx1

    On Error Resume Next
    
    If (ActiveSheet.Cells(Indx1, 37).Value = "○" Or ActiveSheet.Cells(Indx1, 37).Value = "△" Or ActiveSheet.Cells(Indx1, 37).Value = "×" Or ActiveSheet.Cells(Indx1, 37).Value = "=") And (ActiveSheet.Cells(Indx1, 11).Value = "○" Or ActiveSheet.Cells(Indx1, 11).Value = "△" Or ActiveSheet.Cells(Indx1, 11).Value = "×") And ActiveSheet.Cells(Indx1, 68).Value <> "" Then
        GoTo bypass
    End If
    
    ActiveSheet.Cells(Indx1, IraiNO).Value = 0
    ActiveSheet.Cells(Indx1, IraiNO).Value = _
            Worksheets("台帳管理").Cells(WorksheetFunction.Match(ActiveSheet.Cells(Indx1, ShUkNo), Worksheets("台帳管理").Range("O:O"), 0), 3).Value
    ActiveSheet.Cells(Indx1, HokoNO).Value = 0
    ActiveSheet.Cells(Indx1, HokoNO).Value = _
            Worksheets("台帳管理").Cells(WorksheetFunction.Match(ActiveSheet.Cells(Indx1, ShUkNo), Worksheets("台帳管理").Range("O:O"), 0), 4).Value
    
    
    If ActiveSheet.Cells(Indx1, 30).Value = "-" Or ActiveSheet.Cells(Indx1, 30).Value = "" Then
        ActiveSheet.Cells(Indx1, 47).Value = 0
    Else
        ActiveSheet.Cells(Indx1, 47).Value = ActiveSheet.Cells(Indx1, 30).Value
    End If
    If ActiveSheet.Cells(Indx1, 31).Value = "-" Or ActiveSheet.Cells(Indx1, 31).Value = "" Then
        ActiveSheet.Cells(Indx1, 48).Value = 0
    Else
        ActiveSheet.Cells(Indx1, 48).Value = ActiveSheet.Cells(Indx1, 31).Value
    End If
    If ActiveSheet.Cells(Indx1, 32).Value = "-" Or ActiveSheet.Cells(Indx1, 32).Value = "" Then
        ActiveSheet.Cells(Indx1, 49).Value = 0
    Else
        ActiveSheet.Cells(Indx1, 49).Value = ActiveSheet.Cells(Indx1, 32).Value
    End If
    
    
    
'    If (ActiveSheet.Cells(Indx1, 50).Value = "*" Or ActiveSheet.Cells(Indx1, 52).Value = "*" Or ActiveSheet.Cells(Indx1, 54).Value = "*") Then
'        ActiveSheet.Cells(Indx1, Hoko3).Value = ""
'        ActiveSheet.Cells(Indx1, Hoko3).Value = _
'            Worksheets("台帳管理").Cells(WorksheetFunction.Match(ActiveSheet.Cells(Indx1, ShUkNo), Worksheets("台帳管理").Range("O:O"), 0), 34).Value
'        ActiveSheet.Cells(Indx1, Hoko4).Value = ""
'        ActiveSheet.Cells(Indx1, Hoko4).Value = _
'            Worksheets("台帳管理").Cells(WorksheetFunction.Match(ActiveSheet.Cells(Indx1, ShUkNo), Worksheets("台帳管理").Range("O:O"), 0), 35).Value
'    ElseIf (ActiveSheet.Cells(Indx1, 37).Value = "○" Or ActiveSheet.Cells(Indx1, 37).Value = "△" Or ActiveSheet.Cells(Indx1, 37).Value = "×") Then
'        ActiveSheet.Cells(Indx1, Hoko3).Value = ""
'        ActiveSheet.Cells(Indx1, Hoko3).Value = ActiveSheet.Cells(Indx1, 32).Value
'        ActiveSheet.Cells(Indx1, Hoko4).Value = ""
'        ActiveSheet.Cells(Indx1, Hoko4).Value = "Ｇ長"
'    End If

'   ---------------------------------------------
'   完了済の件数をカウントアップ
    Skip_Count = Skip_Count + 1
    
'   台帳管理から「警告」・「要注意」のものを管理台帳にセットする
'    ActiveSheet.Cells(Indx1, 66).Value = ""
'    ActiveSheet.Cells(Indx1, 66).Value = _
'            ActiveSheet.Cells(Indx1, 39).Value
'    ActiveSheet.Cells(Indx1, 66).Value = _
'            ActiveSheet.Cells(Indx1, 66).Value & _
'            vbLf & _
'            Worksheets("台帳管理").Cells(WorksheetFunction.Match(ActiveSheet.Cells(Indx1, KnUkNo), Worksheets("台帳管理").Range("Q:Q"), 0), 36).Value
'            Worksheets("台帳管理").Cells(WorksheetFunction.Match(ActiveSheet.Cells(Indx1, KnUkNo), Worksheets("台帳管理").Range("Q:Q"), 0), 37).Value & _
'            vbLf & _

'   管理台帳が「○」なら「×」・「△」に変更しセットする
'    If InStr(ActiveSheet.Cells(Indx1, 66), "警告") > 0 And (ActiveSheet.Cells(Indx1, 37) = "○" Or ActiveSheet.Cells(Indx1, 37) = "△") Then
'        ActiveSheet.Cells(Indx1, 37) = "×"
'    ElseIf InStr(ActiveSheet.Cells(Indx1, 66), "要注意") > 0 And ActiveSheet.Cells(Indx1, 37) = "○" Then
'        ActiveSheet.Cells(Indx1, 37) = "△"
'    Else
'    End If
'   ---------------------------------------------

bypass:
    
    On Error GoTo 0

    Indx1 = Indx1 + 1
    Application.StatusBar = "CHK-" & Indx1 - 6

Loop

DgetA_E:

End Sub

Sub DgetB_F()

Const DAICHO As String = "案件集約.xlsx"
Const Version As String = "本番化チェックリストの管理台帳.xlsm"

ThisWorkbook.Worksheets("台帳管理").Activate
ActiveSheet.Cells.Clear
ActiveSheet.Range("A1").Select    '指定の位置にセルを置く

'-- PRISMの管理台帳を取り込む --

Workbooks.Open Filename:=myPath & DAICHO
Workbooks(DAICHO).Worksheets("台帳管理").Activate
 
Workbooks(DAICHO).Worksheets("台帳管理").Range("A:AQ").Copy
ThisWorkbook.Worksheets("台帳管理").Range("A1").PasteSpecial Paste:=xlPasteAll
Application.CutCopyMode = False

'-- バージョン情報(管理No.)を取り込む --

Workbooks.Open Filename:=myPath & Version
 
Workbooks(Version).Worksheets("本番化チェックリスト台帳(管理No)").Range("A:I").Copy
ThisWorkbook.Worksheets("管理No").Range("A1").PasteSpecial Paste:=xlPasteAll
Application.CutCopyMode = False

Workbooks(Version).Close SaveChanges:=False, Filename:=Version

End Sub


Sub DputD_F()

Dim END_CELL1 As Integer
Dim END_CELL2 As Integer

Const TempA As String = "案件集約.xlsx"

Workbooks(TempA).Worksheets("管理台帳_PRISM").Range("A6:BM1000").ClearContents
ThisWorkbook.Worksheets("管理台帳_PRISM").Range("A6:BM1000").Copy


Workbooks(TempA).Worksheets("管理台帳_PRISM").Range("A6").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
Application.CutCopyMode = False



    END_CELL1 = ThisWorkbook.Worksheets("管理台帳_PRISM").Cells(6, 1).End(xlDown).Row
    ThisWorkbook.Worksheets("管理台帳_PRISM").Range("AN6", "AW" & END_CELL1).Copy

    ThisWorkbook.Worksheets("ACCESS").Range("A2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False

'    ThisWorkbook.Worksheets("管理台帳_PRISM").Range("AD6", "AF" & END_CELL1).Copy

'    ThisWorkbook.Worksheets("ACCESS").Range("H2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
'    Application.CutCopyMode = False

'----------

    END_CELL2 = Workbooks(TempA).Worksheets("PRISM_ACCESS").Cells(1, 1).End(xlDown).Row
    Workbooks(TempA).Worksheets("PRISM_ACCESS").Range("A2:J" & END_CELL2).ClearContents

    ThisWorkbook.Worksheets("ACCESS").Range("A2:J" & END_CELL1).Copy
    Workbooks(TempA).Worksheets("PRISM_ACCESS").Range("A2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False



Workbooks(TempA).Save
Workbooks(TempA).Close SaveChanges:=False, Filename:=TempA

ThisWorkbook.Worksheets("進捗確認").Range("A6:BM1000").ClearContents
ThisWorkbook.Worksheets("管理台帳_PRISM").Range("A6:BM1000").Copy

ThisWorkbook.Worksheets("進捗確認").Range("A6").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False


End Sub




Sub CHECK_MAIN()

MsgBox "開始します！"

ThisWorkbook.Worksheets("LOG").Range("B3").Value = Now
Call DcheckB_F
ThisWorkbook.Worksheets("LOG").Range("C3").Value = Now

Skip_Count = Indx1 - 6 - Skip_Count
MsgBox "終了しました！" & vbLf & "スキップカウント　：　" & Skip_Count

End Sub


Sub DcheckB_F()
'
' Macro_check1 Macro
'

Dim IraiNO As Long
Dim Irai1 As Long
Dim Irai2 As Long
Dim Shub1 As Long
Dim UkeNo As Long
Dim Test1 As Long
Dim Test2 As Long
Dim Test3 As Long
Dim Atta1 As Long
Dim Atta2 As Long
Dim Atta3 As Long
Dim Atta4 As Long
Dim Atta5 As Long
Dim Atta6 As Long
Dim HokoNO As Long
Dim Hoko1 As Long
Dim Hoko2 As Long
Dim Hoko3 As Long
Dim Hoko4 As Long
Dim ShUkNo As Long
Dim KnUkNo As Long

Dim Flag1 As Long

Shub1 = 13     '列名：発生種別
UkeNo = 17     '列名：受付No&連番
ShUkNo = 40    '集約用受付No
KnUkNo = 42    '管理用受付No
IraiNO = 45    '依頼書文書番号2
HokoNO = 46    '報告書文書番号2
Hoko3 = 67     '列名：報告書検収承認
Hoko4 = 68     '列名：報告書検収承認者
Indx1 = 6
Skip_Count = 0

ThisWorkbook.Sheets("管理台帳_PRISM").Activate
ActiveSheet.Range("BK6").Select

If ActiveSheet.Range("A1").Value < 6 Then
    MsgBox "値が間違っています"
Else: Indx1 = ActiveSheet.Range("A1").Value
End If

Do While ActiveSheet.Cells(Indx1, 37) <> "-" And ActiveSheet.Range("A2").Value > Indx1

    On Error Resume Next
    
'    If (ActiveSheet.Cells(Indx1, 37).Value = "○" Or ActiveSheet.Cells(Indx1, 37).Value = "△" Or ActiveSheet.Cells(Indx1, 37).Value = "×" Or ActiveSheet.Cells(Indx1, 37).Value = "=") And (ActiveSheet.Cells(Indx1, 11).Value = "○" Or ActiveSheet.Cells(Indx1, 11).Value = "△" Or ActiveSheet.Cells(Indx1, 11).Value = "×") And ActiveSheet.Cells(Indx1, 68).Value <> "" Then
'        GoTo bypass
'    End If
    
    If (ActiveSheet.Cells(Indx1, 50).Value = "*" Or ActiveSheet.Cells(Indx1, 52).Value = "*" Or ActiveSheet.Cells(Indx1, 54).Value = "*") Then
        ActiveSheet.Cells(Indx1, Hoko3).Value = ""
        ActiveSheet.Cells(Indx1, Hoko3).Value = _
            Worksheets("台帳管理").Cells(WorksheetFunction.Match(ActiveSheet.Cells(Indx1, ShUkNo), Worksheets("台帳管理").Range("O:O"), 0), 34).Value
        ActiveSheet.Cells(Indx1, Hoko4).Value = ""
        ActiveSheet.Cells(Indx1, Hoko4).Value = _
            Worksheets("台帳管理").Cells(WorksheetFunction.Match(ActiveSheet.Cells(Indx1, ShUkNo), Worksheets("台帳管理").Range("O:O"), 0), 35).Value
    ElseIf (ActiveSheet.Cells(Indx1, 37).Value = "○" Or ActiveSheet.Cells(Indx1, 37).Value = "△" Or ActiveSheet.Cells(Indx1, 37).Value = "×") Then
        ActiveSheet.Cells(Indx1, Hoko3).Value = ""
        ActiveSheet.Cells(Indx1, Hoko3).Value = ActiveSheet.Cells(Indx1, 32).Value
        ActiveSheet.Cells(Indx1, Hoko4).Value = ""
        ActiveSheet.Cells(Indx1, Hoko4).Value = "Ｇ長"
    End If

'   ---------------------------------------------
'   完了済の件数をカウントアップ
    Skip_Count = Skip_Count + 1
    
'   台帳管理から「警告」・「要注意」のものを管理台帳にセットする
    ActiveSheet.Cells(Indx1, 66).Value = ""
    ActiveSheet.Cells(Indx1, 66).Value = _
            ActiveSheet.Cells(Indx1, 39).Value
    ActiveSheet.Cells(Indx1, 66).Value = _
            ActiveSheet.Cells(Indx1, 66).Value & _
            vbLf & _
            Worksheets("台帳管理").Cells(WorksheetFunction.Match(ActiveSheet.Cells(Indx1, KnUkNo), Worksheets("台帳管理").Range("Q:Q"), 0), 36).Value

'   管理台帳が「○」なら「×」・「△」に変更しセットする
    If InStr(ActiveSheet.Cells(Indx1, 66), "警告") > 0 And (ActiveSheet.Cells(Indx1, 37) = "○" Or ActiveSheet.Cells(Indx1, 37) = "△") Then
        ActiveSheet.Cells(Indx1, 37) = "×"
    ElseIf InStr(ActiveSheet.Cells(Indx1, 66), "要注意") > 0 And ActiveSheet.Cells(Indx1, 37) = "○" Then
        ActiveSheet.Cells(Indx1, 37) = "△"
    Else
    End If
'   ---------------------------------------------

bypass:
    
    On Error GoTo 0

    Indx1 = Indx1 + 1
    Application.StatusBar = "CHK-" & Indx1 - 6

Loop

End Sub


