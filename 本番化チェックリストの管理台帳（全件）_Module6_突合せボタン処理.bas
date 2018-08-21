Attribute VB_Name = "突合せボタン処理"
Option Explicit

Sub リソースメンテ_突合せボタン_Click()

' ----------------------------------------------------------------------------------------------
'　突合せ処理
'　　リソースメンテ資源を本番化チェックリストの資源と突き合わせる
' ----------------------------------------------------------------------------------------------

Dim rIdx As Long
Dim F_SRC As String
Dim FoundCell As Range
Dim FoundROW As Long
Dim FoundCount As Long
Const PRJ_NAME As String = "資源別本番化バージョン管理"
Dim strFileName As String

' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' ACCESS操作準備処理
' ____________________________________________

Dim myPath As String
myPath = ThisWorkbook.Path & "\"

Dim adoCn As Object
    strFileName = myPath & "台帳管理_2018.accdb" 'データベースのファイル名
    Set adoCn = CreateObject("ADODB.Connection") 'ADODBコネクションオブジェクトを作成
    adoCn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFileName & ";" 'Accessファイルに接続

Dim adoRs As Object 'ADOレコードセットオブジェクト
    Set adoRs = CreateObject("ADODB.Recordset") 'ADOレコードセットオブジェクトを作成
 
Dim mySQL As String
Dim myRecordSet As New ADODB.Recordset
Dim C_adoFind As String
    
' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' ログ書き出し
' ____________________________________________

Dim WS_操作 As Worksheet
    Set WS_操作 = ThisWorkbook.Worksheets("操作")

    With WS_操作
        .Range("E3").Value = Now
        .Range("Q1").Value = 0
        .Range("Q2").Value = 0
        .Range("R1").Value = "突合せ（リソースメンテ)"
        .Range("R2").Value = "突合せ（リソースメンテ)"
    End With
     
adoRs.Open PRJ_NAME, adoCn, adOpenDynamic, adLockOptimistic
    
Worksheets("リソースメンテ").Activate
rIdx = 2
FoundCount = 0

' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' 資源の突合せを<ACCESSを検索>
' ____________________________________________

Do While Worksheets("リソースメンテ").Range("B" & rIdx) <> ""

    If ActiveSheet.Range("M" & rIdx) <> "" Then
        GoTo Bypass
    End If
    Application.StatusBar = "row() = " & rIdx

    F_SRC = Worksheets("リソースメンテ").Range("E" & rIdx)
    C_adoFind = "VLOOKUPキー = '" & F_SRC & "'"
    adoRs.Find C_adoFind
    If adoRs.BOF = True Then
        WS_操作.Range("Q1").Value = WS_操作.Range("Q1").Value + 1
        GoTo Bypass
    End If
    If adoRs.EOF = True Then
        WS_操作.Range("Q2").Value = WS_操作.Range("Q2").Value + 1
        GoTo Bypass
    End If

    Worksheets("リソースメンテ").Range("M" & rIdx).Value = adoRs!チェックリスト資源
    FoundCount = FoundCount + 1
    
Bypass:
    rIdx = rIdx + 1
        
    adoRs.MoveFirst

Loop

WS_操作.Range("F3").Value = Now
MsgBox "今回突合せ件数 ： " & FoundCount & " 件でした"

End Sub

Sub チェックリスト_突合せボタン_Click()

' ----------------------------------------------------------------------------------------------
'　突合せ処理
'　　本番化チェックリストの資源をリソースメンテ資源と資源と突き合わせる
' ----------------------------------------------------------------------------------------------

Dim rIdx As Long
Dim F_SRC As String
Dim FoundCell As Range
Dim FoundROW As Long
Dim FoundCount As Long
Const PRJ_NAME As String = "案件別本番化リソース管理"
Dim strFileName As String

' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' ACCESS操作準備処理
' ____________________________________________

Dim myPath As String
myPath = ThisWorkbook.Path & "\"

Dim adoCn As Object
    strFileName = myPath & "台帳管理_2018.accdb" 'データベースのファイル名
    Set adoCn = CreateObject("ADODB.Connection") 'ADODBコネクションオブジェクトを作成
    adoCn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFileName & ";" 'Accessファイルに接続

Dim adoRs As Object 'ADOレコードセットオブジェクト
    Set adoRs = CreateObject("ADODB.Recordset") 'ADOレコードセットオブジェクトを作成
 
Dim mySQL As String
Dim myRecordSet As New ADODB.Recordset
Dim C_adoFind As String
    
' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' ログ書き出し
' ____________________________________________

Dim WS_操作 As Worksheet
Dim WS_チェックリスト As Worksheet
    Set WS_操作 = ThisWorkbook.Worksheets("操作")
    Set WS_チェックリスト = ThisWorkbook.Worksheets("全件分割ファイル")

    WS_操作.Range("E5").Value = Now
    WS_操作.Range("Q1").Value = 0
    WS_操作.Range("Q2").Value = 0
    WS_操作.Range("R1").Value = "突合せ（チェックリスト)"
    WS_操作.Range("R2").Value = "突合せ（チェックリスト)"
        
        
        
adoRs.Open PRJ_NAME, adoCn, adOpenDynamic, adLockOptimistic
    
WS_チェックリスト.Activate
rIdx = 6
FoundCount = 0

' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' 資源の突合せを<ACCESSを検索>
' ____________________________________________

Do While WS_チェックリスト.Range("B" & rIdx) <> ""

    If ActiveSheet.Range("E" & rIdx) <> "" Then
        GoTo Bypass
    End If
    
'    Application.StatusBar = "row() = " & rIdx

    F_SRC = WS_チェックリスト.Range("C" & rIdx)
    C_adoFind = "VLOOKUPキー = '" & F_SRC & "'"
    adoRs.Find C_adoFind
    If adoRs.BOF = True Then
        WS_操作.Range("Q1").Value = WS_操作.Range("Q1").Value + 1
        GoTo Bypass
    End If
    If adoRs.EOF = True Then
        WS_操作.Range("Q2").Value = WS_操作.Range("Q2").Value + 1
        GoTo Bypass
    End If

    WS_チェックリスト.Range("E" & rIdx).Value = adoRs!リソース名
    WS_チェックリスト.Range("H" & rIdx).Value = Now
    
    FoundCount = FoundCount + 1
    
Bypass:
    rIdx = rIdx + 1
        
    adoRs.MoveFirst

Loop

WS_操作.Range("F5").Value = Now
MsgBox "今回突合せ件数 ： " & FoundCount & " 件でした"

End Sub
