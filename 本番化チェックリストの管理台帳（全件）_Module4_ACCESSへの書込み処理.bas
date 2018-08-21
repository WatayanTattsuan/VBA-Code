Attribute VB_Name = "ACCESSへの書込み処理"
Option Explicit

Sub TBL_Del_Add_SQL()

' ----------------------------------------------------------------------------------------------
'　ACCESSへの書込み処理
'　　本番化チェックリストの資源を全件DELETE & 全件ADD　で登録する
' ----------------------------------------------------------------------------------------------

Const PRJ_NAME As String = "資源別本番化バージョン管理"
Dim WS As Worksheet

Set WS = Worksheets("全件分割ファイル")

Dim strFileName As String

Dim myPath As String
myPath = ThisWorkbook.Path & "\"
    
    strFileName = myPath & "台帳管理_2018.accdb" 'データベースのファイル名

Dim adoCn As Object
    Set adoCn = CreateObject("ADODB.Connection") 'ADODBコネクションオブジェクトを作成
    adoCn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFileName & ";" 'Accessファイルに接続

Dim adoRs As Object 'ADOレコードセットオブジェクト
    Set adoRs = CreateObject("ADODB.Recordset") 'ADOレコードセットオブジェクトを作成
 
Dim mySQL As String
Dim myRecordSet As New ADODB.Recordset
    mySQL = "DELETE * FROM " & PRJ_NAME
    adoRs.Open mySQL, adoCn
    
    mySQL = "SELECT * FROM " & PRJ_NAME & ";"
    adoRs.Open mySQL, adoCn, adOpenDynamic, adLockOptimistic
  
Dim i As Long
    i = 6
'    i = ActiveSheet.Range("A1").End(xlDown).Row
  
    Do While (WS.Cells(i, 1).Value <> "")
        Application.StatusBar = "row() = " & i
        adoRs.AddNew
            adoRs.Fields("受付No") = WS.Cells(i, 1).Value
            adoRs.Fields("バージョン") = WS.Cells(i, 2).Value
            adoRs.Fields("VLOOKUPキー") = WS.Cells(i, 3).Value
            adoRs.Fields("チェックリスト資源") = WS.Cells(i, 4).Value
            adoRs.Fields("リソースメンテ資源") = WS.Cells(i, 5).Value
            adoRs.Fields("比較") = WS.Cells(i, 6).Value
            adoRs.Fields("異なっている箇所") = WS.Cells(i, 7).Value
        adoRs.Update
        i = i + 1

    Loop

    adoRs.Close 'レコードセットのクローズ
    adoCn.Close 'コネクションのクローズ
 
    Set adoRs = Nothing
    Set adoCn = Nothing  'オブジェクトの破棄
 
    MsgBox "終了しました！"

End Sub


Sub RTBL_Del_Add_SQL()

' ----------------------------------------------------------------------------------------------
'　ACCESSへの書込み処理
'　　リソースメンテ一覧の資源を全件DELETE & 全件ADD　で登録する
' ----------------------------------------------------------------------------------------------

Const PRJ_NAME As String = "案件別本番化リソース管理"
Dim WS As Worksheet
Dim WS_操作 As Worksheet

Set WS = Worksheets("リソースメンテ")
Set WS_操作 = Worksheets("操作")

    WS_操作.Range("E4").Value = Now

Dim strFileName As String

Dim myPath As String
myPath = ThisWorkbook.Path & "\"
    
    strFileName = myPath & "台帳管理_2018.accdb" 'データベースのファイル名

Dim adoCn As Object
    Set adoCn = CreateObject("ADODB.Connection") 'ADODBコネクションオブジェクトを作成
    adoCn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFileName & ";" 'Accessファイルに接続

Dim adoRs As Object 'ADOレコードセットオブジェクト
    Set adoRs = CreateObject("ADODB.Recordset") 'ADOレコードセットオブジェクトを作成
 
Dim mySQL As String
Dim myRecordSet As New ADODB.Recordset
    mySQL = "DELETE * FROM " & PRJ_NAME
    adoRs.Open mySQL, adoCn
    
    mySQL = "SELECT * FROM " & PRJ_NAME & ";"
    adoRs.Open mySQL, adoCn, adOpenDynamic, adLockOptimistic
  
Dim i As Long
    i = 2
'    i = ActiveSheet.Range("A1").End(xlDown).Row
  
    Do While (WS.Cells(i, 2).Value <> "")
        Application.StatusBar = "row() = " & i
        adoRs.AddNew
            adoRs.Fields("No") = Format(WS.Cells(i, 1).Value, 0)
            adoRs.Fields("集約用受付No") = WS.Cells(i, 2).Value
            adoRs.Fields("枝番") = WS.Cells(i, 3).Value
            adoRs.Fields("管理用受付No") = WS.Cells(i, 4).Value
            adoRs.Fields("VLOOKUPキー") = WS.Cells(i, 5).Value
            adoRs.Fields("実施・対応概要") = WS.Cells(i, 6).Value
            adoRs.Fields("実施予定日") = WS.Cells(i, 7).Value
            adoRs.Fields("リソース名") = WS.Cells(i, 8).Value
            adoRs.Fields("リソース日本語名") = WS.Cells(i, 9).Value
            adoRs.Fields("リソース区分") = WS.Cells(i, 10).Value
            adoRs.Fields("区分") = WS.Cells(i, 11).Value
            adoRs.Fields("資源区分") = WS.Cells(i, 12).Value
            adoRs.Fields("突合せ結果") = WS.Cells(i, 13).Value
        adoRs.Update
        i = i + 1

    Loop

    adoRs.Close 'レコードセットのクローズ
    adoCn.Close 'コネクションのクローズ
 
    Set adoRs = Nothing
    Set adoCn = Nothing  'オブジェクトの破棄
 
    WS_操作.Range("F4").Value = Now
    MsgBox "終了しました！"

End Sub


Sub キー再作成()

Dim i As Long
    i = 6
  
    Do While (Worksheets("全件分割ファイル").Range("A" & i).Value <> "")
        Application.StatusBar = "row() = " & i
        Worksheets("全件分割ファイル").Range("C" & i) = _
            Worksheets("全件分割ファイル").Range("A" & i) & Worksheets("全件分割ファイル").Range("D" & i)
        i = i + 1

    Loop

End Sub
