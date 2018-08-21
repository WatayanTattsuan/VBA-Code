Attribute VB_Name = "資源比較処理"
Option Explicit

Dim rsort As Long


Sub 資源比較_sub()

' ----------------------------------------------------------------------------------------------
'　資源の突合せを実施する処理
'　　本番化チェックリストの資源とリソースメンテの資源を比較する処理
' ----------------------------------------------------------------------------------------------

Dim C_ShUkNo As String
Dim strFileName As String
Dim myPath As String
myPath = ThisWorkbook.Path & "\"

Dim rIdx As Long

ActiveSheet.Range("C3:G10000").ClearContents
ActiveSheet.Range("I3:L10000").ClearContents

strFileName = myPath & "台帳管理_2018.accdb" 'データベースのファイル名
C_ShUkNo = ActiveSheet.Range("C1")

Dim adoCn As Object
Set adoCn = CreateObject("ADODB.Connection") 'ADODBコネクションオブジェクトを作成
adoCn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFileName & ";" 'Accessファイルに接続

Dim adoRs As Object 'ADOレコードセットオブジェクト
Set adoRs = CreateObject("ADODB.Recordset") 'ADOレコードセットオブジェクトを作成
 
adoRs.Open "資源別本番化バージョン管理", adoCn, adOpenDynamic, adLockOptimistic

rIdx = 3
adoRs.MoveFirst
adoRs.Filter = "受付No LIKE '" & C_ShUkNo & "*'"
Do Until adoRs.EOF = True

    With ActiveSheet
        .Range("C" & rIdx).Value = adoRs!チェックリスト資源
        .Range("D" & rIdx).Value = adoRs!受付No
        .Range("E" & rIdx).Value = adoRs!異なっている箇所
        .Range("F" & rIdx).Value = adoRs!バージョン
        .Range("G" & rIdx).Value = 999
    End With
    rIdx = rIdx + 1
    adoRs.MoveNext

Loop


Dim adoRRs As Object 'ADOレコードセットオブジェクト
Set adoRRs = CreateObject("ADODB.Recordset") 'ADOレコードセットオブジェクトを作成
adoRRs.Open "案件別本番化リソース管理", adoCn, adOpenDynamic, adLockOptimistic

rIdx = 3
adoRRs.MoveFirst
adoRRs.Filter = "管理用受付No LIKE '" & C_ShUkNo & "*'"
Do Until adoRRs.EOF = True

    With ActiveSheet
        .Range("I" & rIdx).Value = adoRRs!リソース名
        .Range("J" & rIdx).Value = adoRRs!管理用受付No
        .Range("K" & rIdx).Value = adoRRs!資源区分
        .Range("L" & rIdx).Value = 999
    End With
    rIdx = rIdx + 1
    adoRRs.MoveNext

Loop

adoRs.Close 'レコードセットのクローズ
adoRRs.Close 'レコードセットのクローズ
adoCn.Close 'コネクションのクローズ
 
rsort = ActiveSheet.Range("K2").End(xlDown).Row
Call リソース_SORT_sub

rIdx = 3
Dim FoundCell As Range
Dim Find_Rsrc As String
Dim bug_chk As Long
bug_chk = 0


Do While ActiveSheet.Range("C" & rIdx) <> ""
    
    Find_Rsrc = ActiveSheet.Range("C" & rIdx).Value
    Set FoundCell = ActiveSheet.Range("I:I").Find(Find_Rsrc, LookAt:=xlWhole)
    If FoundCell Is Nothing Then
        bug_chk = bug_chk + 1
    Else
        ActiveSheet.Range("G" & rIdx).Value = 0
    End If

    rIdx = rIdx + 1
    
Loop
ActiveSheet.Range("E1").Value = bug_chk

rIdx = 3
bug_chk = 0
Do While ActiveSheet.Range("I" & rIdx) <> ""
    
    Find_Rsrc = ActiveSheet.Range("I" & rIdx).Value
    Set FoundCell = ActiveSheet.Range("C:C").Find(Find_Rsrc, LookAt:=xlWhole)
    If FoundCell Is Nothing Then
        bug_chk = bug_chk + 1
    Else
        ActiveSheet.Range("L" & rIdx).Value = 0
    End If

    rIdx = rIdx + 1
    
Loop

ActiveSheet.Range("K1").Value = bug_chk
Call リソース_SORT_sub

Set adoRs = Nothing
Set adoRRs = Nothing
Set adoCn = Nothing  'オブジェクトの破棄
 
MsgBox "終了しました！"

End Sub

Sub リソース_SORT_sub()
    
    ActiveSheet.Range("F1").Value = rsort
    With ActiveSheet.Range("C3:G" & rsort)
        .Sort Key1:=Range("G3"), order1:=xlAscending, _
        Key2:=Range("D3"), order2:=xlDescending, _
        Key3:=Range("C3"), order3:=xlDescending, _
        Header:=xlYes
    End With

    ActiveSheet.Range("I1").Value = rsort
    With ActiveSheet.Range("I3:M" & rsort)
        .Sort Key1:=Range("L3"), order1:=xlAscending, _
        Key2:=Range("J3"), order2:=xlDescending, _
        Key3:=Range("I3"), order3:=xlDescending, _
        Header:=xlYes
    End With

End Sub




