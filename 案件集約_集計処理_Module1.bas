Attribute VB_Name = "�W�v����_ver3"
Option Explicit

Dim j As Long

Sub �W�v����_ver3_SUB()
'---------------------------------------------------------------------------------------
'���׌����o����
'�@�����̂������ׂ̌��𒊏o���鏈��
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
ShName = Array("�Ǘ��䒠_PRISM", "�Ǘ��䒠_ASTRA", "�Ǘ��䒠_COMMON", "�Ǘ��䒠_�{�л���")

Dim i As Long
Dim x As Long

Set WS3 = Worksheets("�W�v���")

Dim M_Count1 As Long
Dim M_Count2 As Long
Dim M_Count3 As Long

Dim M_Flag As Boolean
Dim M_Flag2 As String
Dim M_Col As String

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
    MsgBox "�J�n�ʒu����яI���ʒu���������Ă��܂�" & vbLf & "�J�n�� < " & vbLf & " �I���� > " & vbLf & "�ł��肢���܂�"
    Exit Sub
End If

j = 2
Worksheets("���O").Range("G2:L50000").Value = ""
MsgBox M_Row_START

For x = 0 To 9
    
    Worksheets("���O").Range("B" & x + 2) = Now
    
    If x > 2 Then
        Set WS1 = Worksheets(ShName(3))
        Set WS2 = Worksheets(ShName(3) & "_2017")
    Else
        Set WS1 = Worksheets(ShName(x))
        Set WS2 = Worksheets(ShName(x) & "_2017")
    End If
    
    
    If InStr(WS1.Name, "ASTRA") <> 0 Then
        M_Flag = True
        M_Flag2 = "��"
        M_Col = "AJ:AJ"
    ElseIf InStr(WS1.Name, "�{�л���") <> 0 Then
        M_Flag = False
        M_Flag2 = "��"
        M_Col = "AI:AI"
    Else
        M_Flag = False
        M_Flag2 = "��"
        M_Col = "AJ:AJ"
    End If
    
    Application.StatusBar = WS1.Name & " - " & UkeNo(x)
    
    i = M_Row_START
    Do While (i < M_Row_ENDED) Or (WS3.Range("B" & i) <> "�W�v")
        
        M_Count3 = F_�������א�(WS1, WS2, WS3, i, M_Flag, UkeNo(x), M_Flag2, M_Col)      '����=>WS1,WS2,WS3,i,M_Flag
            WS3.Cells(i, CellNo(x)) = M_Count3
        
        M_Count3 = F_�������א�(WS1, WS2, WS3, i, M_Flag, UkeNo(x), M_Flag2, M_Col)
            WS3.Cells(i, CellNo(x) + 1) = M_Count3
        
        M_Count3 = F_�i�����Č����א�(WS1, WS2, WS3, i, M_Flag, UkeNo(x), M_Flag2, M_Col)
            WS3.Cells(i, CellNo(x) + 2) = M_Count3

        M_Count3 = F_�����i���s��_�x��(WS1, WS2, WS3, i, M_Flag, UkeNo(x), M_Flag2, M_Col)
            WS3.Cells(i, CellNo(x) + 3) = M_Count3
        
        M_Count3 = F_�����i���s��_�v����(WS1, WS2, WS3, i, M_Flag, UkeNo(x), M_Flag2, M_Col)
            WS3.Cells(i, CellNo(x) + 4) = M_Count3
        
        M_Count3 = F_�������m�F(WS1, WS2, WS3, i, M_Flag, UkeNo(x), M_Flag2, M_Col)
            WS3.Cells(i, CellNo(x) + 5) = M_Count3
                
        M_Count3 = F_S�J��(WS1, WS2, WS3, i, M_Flag, UkeNo(x), M_Flag2, M_Col)
            WS3.Cells(i, CellNo(x) + 6) = M_Count3
          
        M_Count3 = F_�{�ԉ�(WS1, WS2, WS3, i, M_Flag, UkeNo(x), M_Flag2, M_Col)
            WS3.Cells(i, CellNo(x) + 7) = M_Count3
     
        M_Count3 = F_�g���u��(WS1, WS2, WS3, i, M_Flag, UkeNo(x), M_Flag2, M_Col)
            WS3.Cells(i, CellNo(x) + 8) = M_Count3
     
        M_Count3 = F_�Վ�����(WS1, WS2, WS3, i, M_Flag, UkeNo(x), M_Flag2, M_Col)
            WS3.Cells(i, CellNo(x) + 9) = M_Count3
     
        M_Count3 = F_�}�X�^�����e(WS1, WS2, WS3, i, M_Flag, UkeNo(x), M_Flag2, M_Col)
            WS3.Cells(i, CellNo(x) + 10) = M_Count3
     
        M_Count3 = F_�f�[�^�ڍs(WS1, WS2, WS3, i, M_Flag, UkeNo(x), M_Flag2, M_Col)
            WS3.Cells(i, CellNo(x) + 11) = M_Count3

        i = i + 1
    
    Loop
    
    Worksheets("���O").Range("C" & x + 2) = Now

Next x

MsgBox "�I�����܂����I"

End Sub

        
Private Function F_�������א�(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String, ByVal F_Flag2 As String, F_Col As String) As Long
'---------------------------------------------------------------------------------------
'�������א����W�v���܂�
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long
                
        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>-", F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>-", F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>-", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>-", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>-", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>-", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>-", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>-", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
        Else
        
        End If

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' ���O�擾
'_____________________________________

        If Worksheets("���O").Range("F2").Value = True Then
            With Worksheets("���O")
                .Range("G" & j) = F_WS1.Name
                .Range("H" & j) = F_UkeNo
                .Range("I" & j) = F_WS3.Range("B" & F_i)
                .Range("J" & j) = "�������א�"
                .Range("K" & j) = F_Count1
                .Range("L" & j) = Now
            End With
        End If
        
        j = j + 1
'_____________________________________
        
            
        F_�������א� = F_Count1
        
        
End Function

Private Function F_�������א�(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String, ByVal F_Flag2 As String, F_Col As String) As Long
'---------------------------------------------------------------------------------------
'�������א����W�v���܂�
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long
        
        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("K:K"), "��", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("K:K"), "��", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("K:K"), "��", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("K:K"), "��", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("K:K"), "��", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("K:K"), "��", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("K:K"), "��", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("K:K"), "��", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
        Else
        End If
            
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' ���O�擾
'_____________________________________

        If Worksheets("���O").Range("F2").Value = True Then
            With Worksheets("���O")
                .Range("G" & j) = F_WS1.Name
                .Range("H" & j) = F_UkeNo
                .Range("I" & j) = F_WS3.Range("B" & F_i)
                .Range("J" & j) = "�������א�"
                .Range("K" & j) = F_Count1
                .Range("L" & j) = Now
            End With
        End If
        
        j = j + 1
'_____________________________________
        
        F_�������א� = F_Count1

End Function
Private Function F_�i�����Č����א�(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String, ByVal F_Flag2 As String, F_Col As String) As Long
'---------------------------------------------------------------------------------------
'�i�����Č����א����W�v���܂�
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long
        
        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("K:K"), "<>��", F_WS1.Range("K:K"), "<>�~", F_WS1.Range("K:K"), "<>=", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("K:K"), "<>��", F_WS1.Range("K:K"), "<>�~", F_WS1.Range("K:K"), "<>=", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("K:K"), "<>��", F_WS_ASTRA.Range("K:K"), "<>�~", F_WS_ASTRA.Range("K:K"), "<>=", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("K:K"), "<>��", F_WS_ASTRA.Range("K:K"), "<>�~", F_WS_ASTRA.Range("K:K"), "<>=", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("K:K"), "<>��", F_WS_ASTRA.Range("K:K"), "<>�~", F_WS_ASTRA.Range("K:K"), "<>=", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("K:K"), "<>��", F_WS_ASTRA.Range("K:K"), "<>�~", F_WS_ASTRA.Range("K:K"), "<>=", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("K:K"), "<>��", F_WS_ASTRA.Range("K:K"), "<>�~", F_WS_ASTRA.Range("K:K"), "<>=", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("K:K"), "<>��", F_WS_ASTRA.Range("K:K"), "<>�~", F_WS_ASTRA.Range("K:K"), "<>=", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Else
        End If
        
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' ���O�擾
'_____________________________________

        If Worksheets("���O").Range("F2").Value = True Then
            With Worksheets("���O")
                .Range("G" & j) = F_WS1.Name
                .Range("H" & j) = F_UkeNo
                .Range("I" & j) = F_WS3.Range("B" & F_i)
                .Range("J" & j) = "�i�����Č����א�"
                .Range("K" & j) = F_Count1
                .Range("L" & j) = Now
            End With
        End If
        
        j = j + 1
'_____________________________________
        
        F_�i�����Č����א� = F_Count1

End Function
Private Function F_�����i���s��_�x��(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String, ByVal F_Flag2 As String, F_Col As String) As Long
'---------------------------------------------------------------------------------------
'�����i���s��(�x��)�������W�v���܂�
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long
        
        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("BX:BX"), "*�x��*", F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("BX:BX"), "*�x��*", F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("BX:BX"), "*�x��*", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("BX:BX"), "*�x��*", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("BX:BX"), "*�x��*", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("BX:BX"), "*�x��*", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("BX:BX"), "*�x��*", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("BX:BX"), "*�x��*", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Else
        End If
        
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' ���O�擾
'_____________________________________

        If Worksheets("���O").Range("F2").Value = True Then
            With Worksheets("���O")
                .Range("G" & j) = F_WS1.Name
                .Range("H" & j) = F_UkeNo
                .Range("I" & j) = F_WS3.Range("B" & F_i)
                .Range("J" & j) = "�����i���s��(�x��)����"
                .Range("K" & j) = F_Count1
                .Range("L" & j) = Now
            End With
        End If
        
        j = j + 1
'_____________________________________
        
        F_�����i���s��_�x�� = F_Count1
        
End Function
Private Function F_�����i���s��_�v����(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String, ByVal F_Flag2 As String, F_Col As String) As Long
'---------------------------------------------------------------------------------------
'�����i���s��(�v����)�������W�v���܂�
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long

        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("BX:BX"), "*�v����*", F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("BX:BX"), "*�v����*", F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("BX:BX"), "*�v����*", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("BX:BX"), "*�v����*", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("BX:BX"), "*�v����*", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("BX:BX"), "*�v����*", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("BX:BX"), "*�v����*", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("BX:BX"), "*�v����*", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Else
        End If
        
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' ���O�擾
'_____________________________________

        If Worksheets("���O").Range("F2").Value = True Then
            With Worksheets("���O")
                .Range("G" & j) = F_WS1.Name
                .Range("H" & j) = F_UkeNo
                .Range("I" & j) = F_WS3.Range("B" & F_i)
                .Range("J" & j) = "�����i���s��(�v����)����"
                .Range("K" & j) = F_Count1
                .Range("L" & j) = Now
            End With
        End If
        
        j = j + 1
'_____________________________________
        
        F_�����i���s��_�v���� = F_Count1

End Function
Private Function F_�������m�F(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String, ByVal F_Flag2 As String, F_Col As String) As Long
'---------------------------------------------------------------------------------------
'�������m�F�������W�v���܂�
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long

        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>��", F_WS1.Range("AK:AK"), "<>��", F_WS1.Range("AK:AK"), "<>�~", F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>��", F_WS1.Range("AK:AK"), "<>��", F_WS1.Range("AK:AK"), "<>�~", F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>��", F_WS_ASTRA.Range("AK:AK"), "<>��", F_WS_ASTRA.Range("AK:AK"), "<>�~", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>��", F_WS_ASTRA.Range("AK:AK"), "<>��", F_WS_ASTRA.Range("AK:AK"), "<>�~", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>��", F_WS_ASTRA.Range("AK:AK"), "<>��", F_WS_ASTRA.Range("AK:AK"), "<>�~", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>��", F_WS_ASTRA.Range("AK:AK"), "<>��", F_WS_ASTRA.Range("AK:AK"), "<>�~", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>��", F_WS_ASTRA.Range("AK:AK"), "<>��", F_WS_ASTRA.Range("AK:AK"), "<>�~", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>��", F_WS_ASTRA.Range("AK:AK"), "<>��", F_WS_ASTRA.Range("AK:AK"), "<>�~", F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, Range("AN:AN"), F_WS3.Range("B" & F_i))

        Else
        End If

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' ���O�擾
'_____________________________________

        If Worksheets("���O").Range("F2").Value = True Then
            With Worksheets("���O")
                .Range("G" & j) = F_WS1.Name
                .Range("H" & j) = F_UkeNo
                .Range("I" & j) = F_WS3.Range("B" & F_i)
                .Range("J" & j) = "�������m�F����"
                .Range("K" & j) = F_Count1
                .Range("L" & j) = Now
            End With
        End If
        
        j = j + 1
'_____________________________________
                
        F_�������m�F = F_Count1

End Function
Private Function F_S�J��(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String, ByVal F_Flag2 As String, F_Col As String) As Long
'---------------------------------------------------------------------------------------
'�r�J���������W�v���܂�
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long

        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "S�J��", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "S�J��", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "S�J��", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "S�J��", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "S�J��", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "S�J��", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "S�J��", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "S�J��", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Else
        End If
        
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' ���O�擾
'_____________________________________

        If Worksheets("���O").Range("F2").Value = True Then
            With Worksheets("���O")
                .Range("G" & j) = F_WS1.Name
                .Range("H" & j) = F_UkeNo
                .Range("I" & j) = F_WS3.Range("B" & F_i)
                .Range("J" & j) = "�r�J������"
                .Range("K" & j) = F_Count1
                .Range("L" & j) = Now
            End With
        End If
        
        j = j + 1
'_____________________________________
        
        F_S�J�� = F_Count1

End Function
Private Function F_�{�ԉ�(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String, ByVal F_Flag2 As String, F_Col As String) As Long
'---------------------------------------------------------------------------------------
'�{�ԉ��v�揑�y�уe�X�g�v�揑�������W�v���܂�
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long

        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "�{�ԉ�", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "�{�ԉ�", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�{�ԉ�", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�{�ԉ�", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�{�ԉ�", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�{�ԉ�", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�{�ԉ�", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�{�ԉ�", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Else
        End If
        
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' ���O�擾
'_____________________________________

        If Worksheets("���O").Range("F2").Value = True Then
            With Worksheets("���O")
                .Range("G" & j) = F_WS1.Name
                .Range("H" & j) = F_UkeNo
                .Range("I" & j) = F_WS3.Range("B" & F_i)
                .Range("J" & j) = "�{�ԉ��v�揑�y�уe�X�g�v�揑����"
                .Range("K" & j) = F_Count1
                .Range("L" & j) = Now
            End With
        End If
        
        j = j + 1
'_____________________________________
        
        F_�{�ԉ� = F_Count1

End Function
Private Function F_�g���u��(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String, ByVal F_Flag2 As String, F_Col As String) As Long
'---------------------------------------------------------------------------------------
'�g���u���񍐏��������W�v���܂�
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long

        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "�g���u��", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "�g���u��", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�g���u��", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�g���u��", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�g���u��", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�g���u��", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�g���u��", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�g���u��", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Else
        End If
        
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' ���O�擾
'_____________________________________

        If Worksheets("���O").Range("F2").Value = True Then
            With Worksheets("���O")
                .Range("G" & j) = F_WS1.Name
                .Range("H" & j) = F_UkeNo
                .Range("I" & j) = F_WS3.Range("B" & F_i)
                .Range("J" & j) = "�g���u���񍐏�����"
                .Range("K" & j) = F_Count1
                .Range("L" & j) = Now
            End With
        End If
        
        j = j + 1
'_____________________________________
        
        F_�g���u�� = F_Count1

End Function
Private Function F_�Վ�����(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String, ByVal F_Flag2 As String, F_Col As String) As Long
'---------------------------------------------------------------------------------------
'�Վ��������{�񍐏��������W�v���܂��i���Վ������˗��[�ł͂���܂���j
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long

        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "�Վ�����", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "�Վ�����", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�Վ�����", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�Վ�����", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�Վ�����", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�Վ�����", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�Վ�����", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�Վ�����", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Else
        End If
        
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' ���O�擾
'_____________________________________

        If Worksheets("���O").Range("F2").Value = True Then
            With Worksheets("���O")
                .Range("G" & j) = F_WS1.Name
                .Range("H" & j) = F_UkeNo
                .Range("I" & j) = F_WS3.Range("B" & F_i)
                .Range("J" & j) = "�Վ��������{�񍐏�����"
                .Range("K" & j) = F_Count1
                .Range("L" & j) = Now
            End With
        End If
        
        j = j + 1
'_____________________________________
        
        F_�Վ����� = F_Count1

End Function
Private Function F_�}�X�^�����e(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String, ByVal F_Flag2 As String, F_Col As String) As Long
'---------------------------------------------------------------------------------------
'�e��}�X�^�o�^�˗��[�������W�v���܂�
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long

        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "�}�X�^�����e", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "�}�X�^�����e", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�}�X�^�����e", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�}�X�^�����e", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�}�X�^�����e", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�}�X�^�����e", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�}�X�^�����e", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�}�X�^�����e", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Else
        End If
        
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' ���O�擾
'_____________________________________

        If Worksheets("���O").Range("F2").Value = True Then
            With Worksheets("���O")
                .Range("G" & j) = F_WS1.Name
                .Range("H" & j) = F_UkeNo
                .Range("I" & j) = F_WS3.Range("B" & F_i)
                .Range("J" & j) = "�e��}�X�^�o�^�˗��[����"
                .Range("K" & j) = F_Count1
                .Range("L" & j) = Now
            End With
        End If
        
        j = j + 1
'_____________________________________
        
        F_�}�X�^�����e = F_Count1

End Function
Private Function F_�f�[�^�ڍs(ByVal F_WS1 As Worksheet, ByVal F_WS2 As Worksheet, ByVal F_WS3 As Worksheet, ByVal F_i As Long, ByVal F_Flag As Boolean, ByVal F_UkeNo As String, ByVal F_Flag2 As String, F_Col As String) As Long
'---------------------------------------------------------------------------------------
'�f�[�^�ڍs�v�揑�������W�v���܂�
'---------------------------------------------------------------------------------------
        Dim F_Count1 As Long
        Dim F_Count2 As Long

        F_Count1 = 0
        F_Count1 = WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "�f�[�^�ڍs", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Set F_WS1 = F_WS2
        F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS1.Range("AK:AK"), "<>=", F_WS1.Range("V:V"), "�f�[�^�ڍs", F_WS1.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS1.Range("AN:AN"), F_WS3.Range("B" & F_i))
        
        If F_Flag = True Then
            
            Dim F_WS_ASTRA As Worksheet
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�f�[�^�ڍs", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�f�[�^�ڍs", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�f�[�^�ڍs", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            
            Set F_WS_ASTRA = Worksheets("�Ǘ��䒠_�{�л���_2017")
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�f�[�^�ڍs", F_WS_ASTRA.Range("AN:AN"), F_UkeNo, F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�f�[�^�ڍs", F_WS_ASTRA.Range("AN:AN"), "iFAS*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))
            F_Count1 = F_Count1 + WorksheetFunction.CountIfs(F_WS_ASTRA.Range("AK:AK"), "<>=", F_WS_ASTRA.Range("V:V"), "�f�[�^�ڍs", F_WS_ASTRA.Range("AN:AN"), "JINJI*", F_WS1.Range(F_Col), F_Flag2, F_WS_ASTRA.Range("AN:AN"), F_WS3.Range("B" & F_i))

        Else
        End If
        
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' ���O�擾
'_____________________________________

        If Worksheets("���O").Range("F2").Value = True Then
            With Worksheets("���O")
                .Range("G" & j) = F_WS1.Name
                .Range("H" & j) = F_UkeNo
                .Range("I" & j) = F_WS3.Range("B" & F_i)
                .Range("J" & j) = "�f�[�^�ڍs�v�揑����"
                .Range("K" & j) = F_Count1
                .Range("L" & j) = Now
            End With
        End If
        
        j = j + 1
'_____________________________________
        
        F_�f�[�^�ڍs = F_Count1

End Function





