Attribute VB_Name = "�S�������t�@�C�������ݏ���"
Option Explicit

Dim rIdx As Long
Dim rIdy As Long

Sub VgetFILE_all_F()

' ----------------------------------------------------------------------------------------------
'�@�u�S�������t�@�C���v�V�[�g�ւ̏����ݏ�������
'�@�@�u��Ɨp�v�V�[�g����u�S�������t�@�C���v�V�[�g�Ƀf�[�^��\����鏈��
' ----------------------------------------------------------------------------------------------

rIdx = 1

ThisWorkbook.Worksheets("����").Range("E6").Value = Now

With ThisWorkbook.Worksheets("�S�������t�@�C��")
    .Activate
    .Range("A5").Select    '�w��̈ʒu�ɃZ����u��
End With

rIdy = ThisWorkbook.Worksheets("�S�������t�@�C��").Range("A" & Selection.End(xlDown).Row + 1).Row

ThisWorkbook.Worksheets("�S�������t�@�C��").Range("A1").Value = rIdy

Do While ThisWorkbook.Worksheets("��Ɨp").Cells(rIdx, 1) <> ""

    ThisWorkbook.Worksheets("�S�������t�@�C��").Range("A" & rIdy).Value = _
                ThisWorkbook.Worksheets("��Ɨp").Cells(rIdx, 1).Value

    ThisWorkbook.Worksheets("�S�������t�@�C��").Range("B" & rIdy).Value = _
                ThisWorkbook.Worksheets("��Ɨp").Cells(rIdx, 2).Value

    ThisWorkbook.Worksheets("�S�������t�@�C��").Range("D" & rIdy).Value = _
                ThisWorkbook.Worksheets("��Ɨp").Cells(rIdx, Columns.Count).End(xlToLeft).Value

    ThisWorkbook.Worksheets("�S�������t�@�C��").Range("C" & rIdy).Value = _
                ThisWorkbook.Worksheets("�S�������t�@�C��").Range("A" & rIdy).Value & ThisWorkbook.Worksheets("�S�������t�@�C��").Range("D" & rIdy).Value
    
    rIdx = rIdx + 1
    rIdy = rIdy + 1

Loop

Call VsortSUB_F
    
Call TBL_Del_Add_SQL

ThisWorkbook.Worksheets("����").Range("F6").Value = Now
MsgBox "�I�����܂����I"

End Sub

Sub VsortSUB_F()
'
    Dim rsort As Long

' ----------------------------------------------------------------------------------------------
'�@�\�[�g����
' ----------------------------------------------------------------------------------------------

    rsort = ThisWorkbook.Worksheets("�S�������t�@�C��").Range("A" & Selection.End(xlDown).Row).Row
    ThisWorkbook.Worksheets("�S�������t�@�C��").Range("A2").Value = rsort
    
    With ThisWorkbook.Worksheets("�S�������t�@�C��").Range("A5:H" & rsort)
        .Sort Key1:=Range("B1"), order1:=xlDescending, _
        Key2:=Range("A1"), order2:=xlDescending, _
        Key3:=Range("D1"), order3:=xlDescending
    End With
    
    ThisWorkbook.Worksheets("�S�������t�@�C��").Range("A6:H" & rsort).RemoveDuplicates (Array(1, 2, 4))

End Sub



