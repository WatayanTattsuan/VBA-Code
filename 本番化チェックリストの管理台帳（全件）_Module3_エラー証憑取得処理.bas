Attribute VB_Name = "�G���[�؜ߎ擾����"
Option Explicit

Dim Indx1 As Long
Dim Indx2 As Long

Sub �S�������t�@�C��_ERR_MAIN_F()

' ----------------------------------------------------------------------------------------------
'�@�G���[�؜ߎ擾����
'�@�@�u�x���v�E�u�v���Ӂv�̖��ׂ��s�b�N�A�b�v���ꗗ�Ǘ�����
' ----------------------------------------------------------------------------------------------

'MsgBox "�J�n���܂��I"

Dim myPath As String
myPath = ThisWorkbook.Path & "\"

Call Err_DgetA_F          '�V�[�g�u�v���ӈꗗ�v�ɃG���[��\��t����
Call Err_DputD_F          '�t�@�C���u�Ǘ��䒠_�x���E�v����.xlsx�v�ɃG���[��\��t����

ThisWorkbook.Worksheets("�v���ӈꗗ").Range("B1").Value = Now
MsgBox "�I�����܂����I"

End Sub

Sub Err_DgetA_F()

Indx1 = 6
Indx2 = 4

ThisWorkbook.Worksheets("�v���ӈꗗ").Range("B5:H2000").ClearContents

Do While ThisWorkbook.Worksheets("�S�������t�@�C��").Cells(Indx1, 1) <> ""
    
    ThisWorkbook.Worksheets("�v���ӈꗗ").Range("A1").Value = Indx1

'   �Ǘ��䒠����<�x��>�E<�v����>�E<�m�F>�̂��̂��Ǘ��䒠�ɃZ�b�g����
    If InStr(ThisWorkbook.Worksheets("�S�������t�@�C��").Cells(Indx1, 7), "<�x��>") > 0 Or InStr(ThisWorkbook.Worksheets("�S�������t�@�C��").Cells(Indx1, 7), "<�v����>") > 0 Or InStr(ThisWorkbook.Worksheets("�S�������t�@�C��").Cells(Indx1, 7), "<�m�F>") > 0 Then
    Indx2 = Indx2 + 1
    With Worksheets("�v���ӈꗗ")
        .Range("A" & Indx2).Formula = "=ROW()-4"
        .Range("B" & Indx2).Value = Worksheets("�S�������t�@�C��").Range("A" & Indx1).Value
        .Range("C" & Indx2).Value = Worksheets("�S�������t�@�C��").Range("B" & Indx1).Value
        .Range("D" & Indx2).Value = Worksheets("�S�������t�@�C��").Range("D" & Indx1).Value
        .Range("E" & Indx2).Value = Worksheets("�S�������t�@�C��").Range("E" & Indx1).Value
        .Range("F" & Indx2).Value = Worksheets("�S�������t�@�C��").Range("F" & Indx1).Value
        .Range("G" & Indx2).Value = Worksheets("�S�������t�@�C��").Range("G" & Indx1).Value
        .Range("H" & Indx2).Formula = "=vlookup(B" & Indx2 & ",�{�ԉ��ꗗ!$A$2:$K$1000,3,FALSE)"
        End With
    Else
    End If

    Indx1 = Indx1 + 1

Loop

End Sub

Sub Err_DputD_F()

Dim myPath As String
myPath = ThisWorkbook.Path & "\"

Dim END_CELL As Integer
Const TempA As String = "�Ǘ��䒠_�x���E�v����.xlsx"

    Workbooks.Open Filename:=myPath & TempA
    END_CELL = Workbooks(TempA).Worksheets("�{�ԉ�").Cells(4, 1).End(xlDown).Row
    Workbooks(TempA).Worksheets("�{�ԉ�").Range("A5:E" & END_CELL).ClearContents
    END_CELL = ThisWorkbook.Worksheets("�v���ӈꗗ").Cells(4, 1).End(xlDown).Row
    ThisWorkbook.Worksheets("�v���ӈꗗ").Range("B5:G" & END_CELL).Copy

    Workbooks(TempA).Worksheets("�{�ԉ�").Range("B5").PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False

    Workbooks(TempA).Worksheets("�{�ԉ�").Range("B1").Value = Now

    Workbooks(TempA).Save
    Workbooks(TempA).Close SaveChanges:=False, Filename:=TempA


End Sub












