Attribute VB_Name = "リソースメンテ取得"
Option Explicit

Const myPath As String = "E:\SVN\本番化\リソースメンテ一覧\"
Dim rIdx As Long
Dim rIdy As Long
Dim rIdz As Long
Dim fName As String
Dim attach_version As String
Dim attach_date As Date
Dim attach_matter As String
Dim R_TITLE As String
Dim R_UkeNo As String
Dim R_EdaNo As String
Dim R_Attach_date As String
Dim R_Resource As String
Dim G_COUNT As Long
Dim E_COUNT As Long
Dim WS_操作 As Worksheet

Sub リソースメンテ_MAIN_F()

' ----------------------------------------------------------------------------------------------
'　資源取得処理
'　　リソースメンテ一覧から資源を取得する
' ----------------------------------------------------------------------------------------------

fName = Dir(myPath & "*.xls")
rIdx = 1
rIdz = 2

Set WS_操作 = ThisWorkbook.Worksheets("操作")

WS_操作.Range("E2").Value = Now
ThisWorkbook.Worksheets("リソースメンテ").Range("E2:E50000").Formula = ""

'リソースメンテから資源を取得する
Call リソースメンテ_get_SUB            '本番化バージョン管理台帳作成（管理No）

ThisWorkbook.Worksheets("リソースメンテ").Range("E2").Formula = "=[@管理用受付No]&[@リソース名]"

'ソートをして重複しているデータを削除する
Call SORT_DUP_SUB

WS_操作.Range("F2").Value = Now

MsgBox "終了しました！"

End Sub

Private Sub リソースメンテ_get_SUB()

ThisWorkbook.Worksheets("リソースメンテ").Activate

If ThisWorkbook.Worksheets("リソースメンテ").Range("B2").Value = "" Then
    rIdx = 1
Else
    rIdx = ThisWorkbook.Worksheets("リソースメンテ").Range("B1").End(xlDown).Row
End If

WS_操作.Range("I2:J800").Value = ""
WS_操作.Range("L2:M800").Value = ""
G_COUNT = 2
E_COUNT = 2

Do Until fName = ""
    
    Workbooks.Open Filename:=myPath & fName
    Worksheets("リソース一覧").Activate
    R_TITLE = ActiveSheet.Range("E3").Value
    R_UkeNo = ActiveSheet.Range("E4").Value
    R_EdaNo = ActiveSheet.Range("E5").Value
    R_Attach_date = ActiveSheet.Range("I9").Value
    
    rIdy = 10
    Call リソースメンテ_get_Fun("Java", True)
    Call リソースメンテ_get_Fun("ORACLE", True)
    Call リソースメンテ_get_Fun("PGM", True)
    Call リソースメンテ_get_Fun("バッチ", False)
    Call リソースメンテ_get_Fun("画面", False)
    Call リソースメンテ_get_Fun("DB", False)
    Call リソースメンテ_get_Fun("CL資源", False)
    Call リソースメンテ_get_Fun("帳票", False)
    Call リソースメンテ_get_Fun("SVFフォーム", False)
    Call リソースメンテ_get_Fun("SVFクエリ", False)
    Call リソースメンテ_get_Fun("シェル", False)

Bypass:
    Windows(fName).Close
    fName = Dir

Loop
    
End Sub
    
Private Function リソースメンテ_get_Fun(F_SRC As String, F_FLAG As Boolean)
    

' ----------------------------------------------------------------------------------------------
'　資源取得処理(FUNCTION)
'　　リソースメンテ一覧から資源を取得する
' ----------------------------------------------------------------------------------------------

    Dim FoundCell As Range
    ActiveSheet.Range("B" & rIdy).Select
    Set FoundCell = ActiveSheet.Range("B:B").Find(F_SRC, LookAt:=xlPart)
    
    If FoundCell Is Nothing Then
        If F_FLAG = False Then
        Else
        End If
        With WS_操作
            .Range("L" & E_COUNT).Value = F_SRC
            .Range("M" & E_COUNT).Value = R_UkeNo & "-" & R_EdaNo
        End With
        E_COUNT = E_COUNT + 1
        Exit Function
    Else
        WS_操作.Range("I" & G_COUNT).Value = F_SRC
        WS_操作.Range("J" & G_COUNT).Value = R_UkeNo & "-" & R_EdaNo
        G_COUNT = G_COUNT + 1
        rIdy = FoundCell.Row + 2
        R_Resource = FoundCell.Value
    End If
    
    Do While Workbooks(fName).ActiveSheet.Cells(rIdy, 5) <> ""
            
        rIdx = rIdx + 1
        
'        ThisWorkbook.Worksheets("リソースメンテ").Cells(rIdx + 1, 1).Formula = "=row()-1"
        ThisWorkbook.Worksheets("リソースメンテ").Cells(rIdx, 2).Value = R_UkeNo
        ThisWorkbook.Worksheets("リソースメンテ").Cells(rIdx, 3).Value = R_EdaNo
        ThisWorkbook.Worksheets("リソースメンテ").Cells(rIdx, 4).Formula = R_UkeNo & "-" & R_EdaNo
        ThisWorkbook.Worksheets("リソースメンテ").Cells(rIdx, 6).Value = R_TITLE
        ThisWorkbook.Worksheets("リソースメンテ").Cells(rIdx, 7).Value = Format(R_Attach_date, "YYYY/MM/DD")
            
        ThisWorkbook.Worksheets("リソースメンテ").Cells(rIdx, 8).Value = Trim(ActiveSheet.Range("E" & rIdy))
        ThisWorkbook.Worksheets("リソースメンテ").Cells(rIdx, 9).Value = ActiveSheet.Range("F" & rIdy)
        ThisWorkbook.Worksheets("リソースメンテ").Cells(rIdx, 10).Value = ActiveSheet.Range("C" & rIdy)
        ThisWorkbook.Worksheets("リソースメンテ").Cells(rIdx, 11).Value = ActiveSheet.Range("D" & rIdy)
        ThisWorkbook.Worksheets("リソースメンテ").Cells(rIdx, 12).Value = Trim(R_Resource)
            
        rIdy = rIdy + 1
    
    Loop


End Function

    
    
    
    
Sub VreplSUB_F()

' ----------------------------------------------------------------------------------------------
'　形式を統一する為に特定の文字記号を置換する
'
' ----------------------------------------------------------------------------------------------
    ThisWorkbook.Worksheets("リソースメンテ").Activate
    Columns("A:A").Select
    Selection.Replace What:="_", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="－", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="＿", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" H", Replacement:="-H", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

Sub SORT_DUP_SUB()
    Dim rsort As Long
    rsort = ThisWorkbook.Worksheets("リソースメンテ").Range("B1").End(xlDown).Row
'SORT処理
    ThisWorkbook.Worksheets("リソースメンテ").Range("A2:M" & rsort).Sort Key1:=Range("E1"), order1:=xlDescending
'重複削除T処理
    ThisWorkbook.Worksheets("リソースメンテ").Range("A2:M" & rsort).RemoveDuplicates (Array(5, 4, 3))

End Sub


