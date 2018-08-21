Attribute VB_Name = "ACCESSから抽出処理"
Option Explicit

Sub リソース_SELECT_SQL()

' ----------------------------------------------------------------------------------------------
'　データ抽出処理
'　　ACCESSのリソースメンテの資源を全件抽出する処理
' ----------------------------------------------------------------------------------------------

Dim WS_リソースメンテ As Worksheet
Dim C_ShUkNo As String
Dim strFileName As String
Dim myPath As String
myPath = ThisWorkbook.Path & "\"

strFileName = myPath & "台帳管理_2018.accdb" 'データベースのファイル名

Dim adoCn As Object
Set adoCn = CreateObject("ADODB.Connection") 'ADODBコネクションオブジェクトを作成
adoCn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFileName & ";" 'Accessファイルに接続

Dim adoRs As Object 'ADOレコードセットオブジェクト
Set adoRs = CreateObject("ADODB.Recordset") 'ADOレコードセットオブジェクトを作成

Dim WS_操作 As Worksheet
    Set WS_操作 = ThisWorkbook.Worksheets("操作")

WS_操作.Range("E8").Value = Now

C_ShUkNo = "*"
Dim strSQL As String
If C_ShUkNo = "*" Then
    strSQL = "SELECT * FROM 案件別本番化リソース管理"
Else
    strSQL = "SELECT * FROM 案件別本番化リソース管理 where 管理用受付No like " & """" & C_ShUkNo & "%" & """"
End If
 
adoRs.Open strSQL, adoCn    'SQLを実行して対象をRecordSetへ

'
' 書き出すために新しいbookを作成し、シート名を"全件分割ファイルx"とする
'
Workbooks.Add
With ActiveSheet
        .Name = "全件分割ファイルx"
    .Range("A1") = "No"
    .Range("B1") = "集約用受付No"
    .Range("C1") = "枝番"
    .Range("D1") = "管理用受付No"
    .Range("E1") = "VLOOKUPキー"
    .Range("F1") = "実施・対応概要"
    .Range("G1") = "実施予定日"
    .Range("H1") = "リソース名"
    .Range("I1") = "リソース日本語名"
    .Range("J1") = "リソース区分"
    .Range("K1") = "区分"
    .Range("L1") = "資源区分"
    .Range("M1") = "突合せ結果"
End With

ActiveSheet.Range("A2").CopyFromRecordset adoRs
 
adoRs.Close 'レコードセットのクローズ
adoCn.Close 'コネクションのクローズ
 
Set adoRs = Nothing
Set adoCn = Nothing  'オブジェクトの破棄
 
WS_操作.Range("F8").Value = Now
MsgBox "終了しました！"

End Sub
