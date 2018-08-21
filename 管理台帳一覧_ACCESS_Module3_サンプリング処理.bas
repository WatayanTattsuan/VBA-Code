Attribute VB_Name = "サンプリング取得"
Option Explicit

Sub サンプリング取得()

'------------------------------------------------------------------------------------------
'　サンプリング取得処理
'------------------------------------------------------------------------------------------

    Dim i As Long
    Dim x As Long
    Dim y As Long
    
    Dim WS As Worksheet
    Dim WS2 As Worksheet
        
    Set WS = Worksheets("sheet1")
    set WS2 = worksheets("sheet2")
    
    ActiveSheet.Range("C1:C100").Value = ""
    For i = 1 To 30
        x = WS.Range("B6").End(xlDown).Row
        y = Int((x - 6) * Rnd + 6)
        ActiveSheet.Range("c" & i).Value = WS.Range("B" & y).Value
        ActiveSheet.Range("d" & i).Value = y
     
    Next i
    
    MsgBox "終了しました"

End Sub
