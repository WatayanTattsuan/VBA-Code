Attribute VB_Name = "�T���v�����O�擾"
Option Explicit

Sub �T���v�����O�擾()

'------------------------------------------------------------------------------------------
'�@�T���v�����O�擾����
'------------------------------------------------------------------------------------------

    Dim i As Long
    Dim x As Long
    Dim y As Long
    
    Dim WS As Worksheet
        
    Set WS = Worksheets("sheet1")
        
    ActiveSheet.Range("C1:C100").Value = ""
    For i = 1 To 30
        x = WS.Range("B6").End(xlDown).Row
        y = Int((x - 6) * Rnd + 6)
        ActiveSheet.Range("c" & i).Value = WS.Range("B" & y).Value
        ActiveSheet.Range("d" & i).Value = y
     
    Next i
    
    MsgBox "�I�����܂���"

End Sub
