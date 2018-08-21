Attribute VB_Name = "本番化チェックリスト取得"
Option Explicit

Const myPath As String = "D:\SVN\本番化\"
Dim rIdx As Long
Dim rIdy As Long
Dim rIdz As Long
Dim fName As String
Dim attach_version As String
Dim attach_date As Date
Dim attach_matter As String
Dim WS As New Worksheet
Dim WS_LOG As New Worksheet

Sub VgetALL_F()

Set WS = ThisWorkbook.Worksheets("資源全反映")
Set WS_LOG = ThisWorkbook.Worksheets("操作")

WS_LOG.Range("E7").Value = Now
' ----------------------------------------------------------------------------------------------
'　資源取得処理
'　　本番化チェックリストから資源を取得する
' ----------------------------------------------------------------------------------------------

fName = Dir(myPath & "*.xls")
rIdx = 1
rIdz = 2

Call VgetAttachSUB_F        '本番化バージョン管理台帳作成（管理No）
Call VreplSUB_F             '誤植統一（置換）

WS_LOG.Range("F7").Value = Now
MsgBox "終了しました！"

End Sub

Sub VgetAttachSUB_F()

WS.Activate
rIdx = WS.Cells(1, 1).End(xlDown).Row

WS.Range("A" & rIdx).Select    '指定の位置にセルを置く

Do Until fName = ""
    
    Workbooks.Open Filename:=myPath & fName
    Worksheets("表紙").Activate
    attach_version = ActiveSheet.Cells(18, 4).Value
    attach_date = ActiveSheet.Cells(22, 4).Value
    attach_matter = ActiveSheet.Cells(17, 4).Value
    Workbooks(fName).Worksheets("差分一覧").Activate

    '********** 差分一覧からデータを抽出してくる *********************************
    
    rIdy = 3
    Do While Workbooks(fName).ActiveSheet.Cells(rIdy, 1) <> ""
            
        If Len(Workbooks(fName).ActiveSheet.Cells(rIdy, 7).Value) > 5 Then
            
            rIdx = rIdx + 1

            With WS
                .Cells(rIdx, 1).Value = ActiveSheet.Cells(rIdy, 7).Value
                .Cells(rIdx, 2).Value = attach_version
                .Cells(rIdx, 3).Value = attach_date
                .Cells(rIdx, 4).Formula = "=workday(C" & rIdx & ",1,祝日設定!$A$2:$A$1000)"
                .Cells(rIdx, 5).Value = attach_matter
                .Cells(rIdx, 9).Value = fName
                .Cells(rIdx, 10).Value = Workbooks(fName).ActiveSheet.Name
                .Cells(rIdx, 11).Value = ActiveSheet.Cells(rIdy, 2).Value
            End With
            
            If Mid(WS.Cells(rIdx, 1).Value, 1, 5) = "PRISM" Then
                WS.Cells(rIdx, 6).Value = ActiveSheet.Cells(rIdy, 9).Value
            ElseIf Mid(WS.Cells(rIdx, 1).Value, 1, 5) = "ASTRA" Then
                WS.Cells(rIdx, 7).Value = ActiveSheet.Cells(rIdy, 9).Value
            ElseIf Mid(WS.Cells(rIdx, 1).Value, 1, 5) = "JINJI" Then
                WS.Cells(rIdx, 7).Value = ActiveSheet.Cells(rIdy, 9).Value
            Else
                WS.Cells(rIdx, 8).Value = ActiveSheet.Cells(rIdy, 9).Value
            End If
            
        End If
        rIdy = rIdy + 1
    Loop

    '********** DB反映一覧からデータを抽出してくる *********************************

    Workbooks(fName).Worksheets("DB反映一覧").Activate
    
    rIdy = 3
    Do While Workbooks(fName).ActiveSheet.Cells(rIdy, 1) <> ""
        If Len(Workbooks(fName).ActiveSheet.Cells(rIdy, 5).Value) > 5 Then
            
            rIdx = rIdx + 1
            With WS
                .Cells(rIdx, 1).Value = ActiveSheet.Cells(rIdy, 5).Value
                .Cells(rIdx, 2).Value = attach_version
                .Cells(rIdx, 3).Value = attach_date
                .Cells(rIdx, 4).Formula = "=workday(C" & rIdx & ",1,祝日設定!$A$2:$A$1000)"
                .Cells(rIdx, 5).Value = attach_matter
                .Cells(rIdx, 9).Value = fName
                .Cells(rIdx, 10).Value = Workbooks(fName).ActiveSheet.Name
                .Cells(rIdx, 11).Value = ActiveSheet.Cells(rIdy, 2).Value
            End With
            
            If Mid(WS.Cells(rIdx, 1).Value, 1, 5) = "PRISM" Then
                WS.Cells(rIdx, 6).Value = ActiveSheet.Cells(rIdy, 3).Value
            ElseIf Mid(WS.Cells(rIdx, 1).Value, 1, 5) = "ASTRA" Then
                WS.Cells(rIdx, 7).Value = ActiveSheet.Cells(rIdy, 3).Value
            ElseIf Mid(WS.Cells(rIdx, 1).Value, 1, 5) = "JINJI" Then
                WS.Cells(rIdx, 7).Value = ActiveSheet.Cells(rIdy, 3).Value
            Else
                WS.Cells(rIdx, 8).Value = ActiveSheet.Cells(rIdy, 3).Value
            End If
            
        End If
        rIdy = rIdy + 1
              
    Loop
    
    Windows(fName).Close
    fName = Dir

Loop
    
End Sub
    
Sub VreplSUB_F()

    '********** 置換
    WS.Activate
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





