Attribute VB_Name = "集計処理"
Option Explicit

Sub 集計処理_SUB()
'---------------------------------------------------------------------------------------
'明細個数抽出処理
'　条件のあう明細の個数を抽出する処理
'---------------------------------------------------------------------------------------
'
Dim WS1 As Worksheet
Dim WS2 As Worksheet
Dim WS3 As Worksheet

Dim UkeNo() As Variant
Dim CellNo() As Variant
Dim ShName() As Variant
UkeNo = Array("PRISM*", "ASTRA*", "COMMON*", "iFAS*", "JINJI*", "CSJIN*", "CSZIM*", "TMSP*", "FA*", "WEBAP*")
CellNo = Array(3, 15, 27, 51, 63, 75, 87, 99, 111, 123)
ShName = Array("管理台帳_PRISM", "管理台帳_ASTRA", "管理台帳_COMMON", "管理台帳_本社ｻｰﾊﾞ")

Dim i As Long
Dim x As Long

Set WS3 = Worksheets("集計情報___")

Dim M_Count1 As Long
Dim M_Count2 As Long
Dim M_Count3 As Long

Dim M_Flag As Boolean

Dim M_Row_START As Long
Dim M_Row_ENDED As Long

Dim FoundCell As Range
Set FoundCell = WS3.Range("A:A").Find("<", LookAt:=xlWhole)

If FoundCell Is Nothing Then
        M_Row_START = 5
Else
        M_Row_START = FoundCell.Row
End If
Set FoundCell = WS3.Range("A:A").Find(">", LookAt:=xlWhole)
If FoundCell Is Nothing Then
        M_Row_ENDED = 16
Else
        M_Row_ENDED = FoundCell.Row
End If

If M_Row_ENDED < M_Row_START Then
    MsgBox "開始位置および終了位置が矛盾しています" & vbLf & "開始は < " & vbLf & " 終了は > " & vbLf & "でお願いします"
    Exit Sub
End If

MsgBox M_Row_START

For x = 0 To 9
    
    If x > 2 Then
        Set WS1 = Worksheets(ShName(3))
        Set WS2 = Worksheets(ShName(3) & "_2017")
    Else
        Set WS1 = Worksheets(ShName(x))
        Set WS2 = Worksheets(ShName(x) & "_2017")
    End If
    
    
    If InStr(WS1.Name, "本社ｻｰﾊﾞ") <> 0 Then
        M_Flag = True
    Else
        M_Flag = False
    End If
    
    Application.StatusBar = WS1.Name & " - " & UkeNo(x)
    
    i = M_Row_START
    Do While (i < M_Row_ENDED) Or (WS3.Range("B" & i) <> "集計")
        
        M_Count3 = F_発生明細数(WS1, WS2, WS3, i, M_Flag, UkeNo(x))        '引数=>WS1,WS2,WS3,i,M_Flag
            WS3.Cells(i, CellNo(x)) = M_Count3
        
        M_Count3 = F_完了明細数(WS1, WS2, WS3, i, M_Flag, UkeNo(x))
            WS3.Cells(i, CellNo(x) + 1) = M_Count3
        
        M_Count3 = F_進捗中案件明細数(WS1, WS2, WS3, i, M_Flag, UkeNo(x))
            WS3.Cells(i, CellNo(x) + 2) = M_Count3

        M_Count3 = F_資料品質不良_警告(WS1, WS2, WS3, i, M_Flag, UkeNo(x))
            WS3.Cells(i, CellNo(x) + 3) = M_Count3
        
        M_Count3 = F_資料品質不良_要注意(WS1, WS2, WS3, i, M_Flag, UkeNo(x))
            WS3.Cells(i, CellNo(x) + 4) = M_Count3
        
        M_Count3 = F_資料未確認(WS1, WS2, WS3, i, M_Flag, UkeNo(x))
            WS3.Cells(i, CellNo(x) + 5) = M_Count3
                
        M_Count3 = F_S開発(WS1, WS2, WS3, i, M_Flag, UkeNo(x))
            WS3.Cells(i, CellNo(x) + 6) = M_Count3
          
        M_Count3 = F_本番化(WS1, WS2, WS3, i, M_Flag, UkeNo(x))
            WS3.Cells(i, CellNo(x) + 7) = M_Count3
     
        M_Count3 = F_トラブル(WS1, WS2, WS3, i, M_Flag, UkeNo(x))
            WS3.Cells(i, CellNo(x) + 8) = M_Count3
     
        M_Count3 = F_臨時処理(WS1, WS2, WS3, i, M_Flag, UkeNo(x))
            WS3.Cells(i, CellNo(x) + 9) = M_Count3
     
        M_Count3 = F_マスタメンテ(WS1, WS2, WS3, i, M_Flag, UkeNo(x))
            WS3.Cells(i, CellNo(x) + 10) = M_Count3
     
        M_Count3 = F_データ移行(WS1, WS2, WS3, i, M_Flag, UkeNo(x))
            WS3.Cells(i, CellNo(x) + 11) = M_Count3

        i = i + 1
    
    Loop

Next x

MsgBox "終了しました！"

End Sub

        
Private Function F_発生明細数(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String) As Long
'---------------------------------------------------------------------------------------
'発生明細数を集計します
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long
                
        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>-", F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>-", F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>-", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>-", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
        Else
        End If
            
        F_発生明細数 = F_Count1
        
        
End Function

Private Function F_完了明細数(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String) As Long
'---------------------------------------------------------------------------------------
'完了明細数を集計します
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long
        
        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("K:K"), "○", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("K:K"), "○", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("K:K"), "○", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("K:K"), "○", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
        Else
        End If
            
        F_完了明細数 = F_Count1

End Function
Private Function F_進捗中案件明細数(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String) As Long
'---------------------------------------------------------------------------------------
'進捗中案件明細数を集計します
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long
        
        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("K:K"), "<>○", F_WS1.Range("K:K"), "<>×", F_WS1.Range("K:K"), "<>=", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("K:K"), "<>○", F_WS1.Range("K:K"), "<>×", F_WS1.Range("K:K"), "<>=", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("K:K"), "<>○", F_WS_ASTRA.Range("K:K"), "<>×", F_WS_ASTRA.Range("K:K"), "<>=", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("K:K"), "<>○", F_WS_ASTRA.Range("K:K"), "<>×", F_WS_ASTRA.Range("K:K"), "<>=", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Else
        End If
        
        F_進捗中案件明細数 = F_Count1

End Function
Private Function F_資料品質不良_警告(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String) As Long
'---------------------------------------------------------------------------------------
'資料品質不良(警告)件数を集計します
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long
        
        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("BX:BX"), "*警告*", F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("BX:BX"), "*警告*", F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("BX:BX"), "*警告*", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("BX:BX"), "*警告*", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Else
        End If
        
        F_資料品質不良_警告 = F_Count1
        
End Function
Private Function F_資料品質不良_要注意(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String) As Long
'---------------------------------------------------------------------------------------
'資料品質不良(要注意)件数を集計します
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long

        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("BX:BX"), "*要注意*", F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("BX:BX"), "*要注意*", F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("BX:BX"), "*要注意*", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("BX:BX"), "*要注意*", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Else
        End If
        
        F_資料品質不良_要注意 = F_Count1

End Function
Private Function F_資料未確認(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String) As Long
'---------------------------------------------------------------------------------------
'資料未確認件数を集計します
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long

        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>○", F_WS1.Range("AK:AK"), "<>△", F_WS1.Range("AK:AK"), "<>×", F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>○", F_WS1.Range("AK:AK"), "<>△", F_WS1.Range("AK:AK"), "<>×", F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("AN:AN"), F_UkeNo, Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>○", F_WS_ASTRA.Range("AK:AK"), "<>△", F_WS_ASTRA.Range("AK:AK"), "<>×", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>○", F_WS_ASTRA.Range("AK:AK"), "<>△", F_WS_ASTRA.Range("AK:AK"), "<>×", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, Range("AN:AN"), F_WS3.Range("B" & F_i))

        Else
        End If
        
        F_資料未確認 = F_Count1

End Function
Private Function F_S開発(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String) As Long
'---------------------------------------------------------------------------------------
'Ｓ開発件数を集計します
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long

        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "S開発", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "S開発", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "S開発", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "S開発", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Else
        End If
        
        F_S開発 = F_Count1

End Function
Private Function F_本番化(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String) As Long
'---------------------------------------------------------------------------------------
'本番化計画書及びテスト計画書件数を集計します
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long

        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "本番化", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "本番化", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "本番化", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "本番化", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Else
        End If
        
        F_本番化 = F_Count1

End Function
Private Function F_トラブル(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String) As Long
'---------------------------------------------------------------------------------------
'トラブル報告書件数を集計します
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long

        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "トラブル", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "トラブル", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "トラブル", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "トラブル", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Else
        End If
        
        F_トラブル = F_Count1

End Function
Private Function F_臨時処理(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String) As Long
'---------------------------------------------------------------------------------------
'臨時処理実施報告書件数を集計します（※臨時処理依頼票ではありません）
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long

        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "臨時処理", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "臨時処理", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "臨時処理", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "臨時処理", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Else
        End If
        
        F_臨時処理 = F_Count1

End Function
Private Function F_マスタメンテ(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String) As Long
'---------------------------------------------------------------------------------------
'各種マスタ登録依頼票件数を集計します
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long

        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "マスタメンテ", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "マスタメンテ", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "マスタメンテ", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "マスタメンテ", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Else
        End If
        
        F_マスタメンテ = F_Count1

End Function
Private Function F_データ移行(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String) As Long
'---------------------------------------------------------------------------------------
'データ移行計画書件数を集計します
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long

        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "データ移行", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "データ移行", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "データ移行", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("管理台帳_ASTRA_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "データ移行", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Else
        End If
        
        F_データ移行 = F_Count1

End Function


