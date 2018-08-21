Attribute VB_Name = "���\�[�X�����e�擾"
Option Explicit

Const myPath As String = "E:\SVN\�{�ԉ�\���\�[�X�����e�ꗗ\"
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
Dim WS_���� As Worksheet

Sub ���\�[�X�����e_MAIN_F()

' ----------------------------------------------------------------------------------------------
'�@�����擾����
'�@�@���\�[�X�����e�ꗗ���玑�����擾����
' ----------------------------------------------------------------------------------------------

fName = Dir(myPath & "*.xls")
rIdx = 1
rIdz = 2

Set WS_���� = ThisWorkbook.Worksheets("����")

WS_����.Range("E2").Value = Now
ThisWorkbook.Worksheets("���\�[�X�����e").Range("E2:E50000").Formula = ""

'���\�[�X�����e���玑�����擾����
Call ���\�[�X�����e_get_SUB            '�{�ԉ��o�[�W�����Ǘ��䒠�쐬�i�Ǘ�No�j

ThisWorkbook.Worksheets("���\�[�X�����e").Range("E2").Formula = "=[@�Ǘ��p��tNo]&[@���\�[�X��]"

'�\�[�g�����ďd�����Ă���f�[�^���폜����
Call SORT_DUP_SUB

WS_����.Range("F2").Value = Now

MsgBox "�I�����܂����I"

End Sub

Private Sub ���\�[�X�����e_get_SUB()

ThisWorkbook.Worksheets("���\�[�X�����e").Activate

If ThisWorkbook.Worksheets("���\�[�X�����e").Range("B2").Value = "" Then
    rIdx = 1
Else
    rIdx = ThisWorkbook.Worksheets("���\�[�X�����e").Range("B1").End(xlDown).Row
End If

WS_����.Range("I2:J800").Value = ""
WS_����.Range("L2:M800").Value = ""
G_COUNT = 2
E_COUNT = 2

Do Until fName = ""
    
    Workbooks.Open Filename:=myPath & fName
    Worksheets("���\�[�X�ꗗ").Activate
    R_TITLE = ActiveSheet.Range("E3").Value
    R_UkeNo = ActiveSheet.Range("E4").Value
    R_EdaNo = ActiveSheet.Range("E5").Value
    R_Attach_date = ActiveSheet.Range("I9").Value
    
    rIdy = 10
    Call ���\�[�X�����e_get_Fun("Java", True)
    Call ���\�[�X�����e_get_Fun("ORACLE", True)
    Call ���\�[�X�����e_get_Fun("PGM", True)
    Call ���\�[�X�����e_get_Fun("�o�b�`", False)
    Call ���\�[�X�����e_get_Fun("���", False)
    Call ���\�[�X�����e_get_Fun("DB", False)
    Call ���\�[�X�����e_get_Fun("CL����", False)
    Call ���\�[�X�����e_get_Fun("���[", False)
    Call ���\�[�X�����e_get_Fun("SVF�t�H�[��", False)
    Call ���\�[�X�����e_get_Fun("SVF�N�G��", False)
    Call ���\�[�X�����e_get_Fun("�V�F��", False)

Bypass:
    Windows(fName).Close
    fName = Dir

Loop
    
End Sub
    
Private Function ���\�[�X�����e_get_Fun(F_SRC As String, F_FLAG As Boolean)
    

' ----------------------------------------------------------------------------------------------
'�@�����擾����(FUNCTION)
'�@�@���\�[�X�����e�ꗗ���玑�����擾����
' ----------------------------------------------------------------------------------------------

    Dim FoundCell As Range
    ActiveSheet.Range("B" & rIdy).Select
    Set FoundCell = ActiveSheet.Range("B:B").Find(F_SRC, LookAt:=xlPart)
    
    If FoundCell Is Nothing Then
        If F_FLAG = False Then
        Else
        End If
        With WS_����
            .Range("L" & E_COUNT).Value = F_SRC
            .Range("M" & E_COUNT).Value = R_UkeNo & "-" & R_EdaNo
        End With
        E_COUNT = E_COUNT + 1
        Exit Function
    Else
        WS_����.Range("I" & G_COUNT).Value = F_SRC
        WS_����.Range("J" & G_COUNT).Value = R_UkeNo & "-" & R_EdaNo
        G_COUNT = G_COUNT + 1
        rIdy = FoundCell.Row + 2
        R_Resource = FoundCell.Value
    End If
    
    Do While Workbooks(fName).ActiveSheet.Cells(rIdy, 5) <> ""
            
        rIdx = rIdx + 1
        
'        ThisWorkbook.Worksheets("���\�[�X�����e").Cells(rIdx + 1, 1).Formula = "=row()-1"
        ThisWorkbook.Worksheets("���\�[�X�����e").Cells(rIdx, 2).Value = R_UkeNo
        ThisWorkbook.Worksheets("���\�[�X�����e").Cells(rIdx, 3).Value = R_EdaNo
        ThisWorkbook.Worksheets("���\�[�X�����e").Cells(rIdx, 4).Formula = R_UkeNo & "-" & R_EdaNo
        ThisWorkbook.Worksheets("���\�[�X�����e").Cells(rIdx, 6).Value = R_TITLE
        ThisWorkbook.Worksheets("���\�[�X�����e").Cells(rIdx, 7).Value = Format(R_Attach_date, "YYYY/MM/DD")
            
        ThisWorkbook.Worksheets("���\�[�X�����e").Cells(rIdx, 8).Value = Trim(ActiveSheet.Range("E" & rIdy))
        ThisWorkbook.Worksheets("���\�[�X�����e").Cells(rIdx, 9).Value = ActiveSheet.Range("F" & rIdy)
        ThisWorkbook.Worksheets("���\�[�X�����e").Cells(rIdx, 10).Value = ActiveSheet.Range("C" & rIdy)
        ThisWorkbook.Worksheets("���\�[�X�����e").Cells(rIdx, 11).Value = ActiveSheet.Range("D" & rIdy)
        ThisWorkbook.Worksheets("���\�[�X�����e").Cells(rIdx, 12).Value = Trim(R_Resource)
            
        rIdy = rIdy + 1
    
    Loop


End Function

    
    
    
    
Sub VreplSUB_F()

' ----------------------------------------------------------------------------------------------
'�@�`���𓝈ꂷ��ׂɓ���̕����L����u������
'
' ----------------------------------------------------------------------------------------------
    ThisWorkbook.Worksheets("���\�[�X�����e").Activate
    Columns("A:A").Select
    Selection.Replace What:="_", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="�|", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="�Q", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" H", Replacement:="-H", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

Sub SORT_DUP_SUB()
    Dim rsort As Long
    rsort = ThisWorkbook.Worksheets("���\�[�X�����e").Range("B1").End(xlDown).Row
'SORT����
    ThisWorkbook.Worksheets("���\�[�X�����e").Range("A2:M" & rsort).Sort Key1:=Range("E1"), order1:=xlDescending
'�d���폜T����
    ThisWorkbook.Worksheets("���\�[�X�����e").Range("A2:M" & rsort).RemoveDuplicates (Array(5, 4, 3))

End Sub


