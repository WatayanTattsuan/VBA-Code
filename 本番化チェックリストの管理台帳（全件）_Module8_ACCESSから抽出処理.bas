Attribute VB_Name = "ACCESS���璊�o����"
Option Explicit

Sub ���\�[�X_SELECT_SQL()

' ----------------------------------------------------------------------------------------------
'�@�f�[�^���o����
'�@�@ACCESS�̃��\�[�X�����e�̎�����S�����o���鏈��
' ----------------------------------------------------------------------------------------------

Dim WS_���\�[�X�����e As Worksheet
Dim C_ShUkNo As String
Dim strFileName As String
Dim myPath As String
myPath = ThisWorkbook.Path & "\"

strFileName = myPath & "�䒠�Ǘ�_2018.accdb" '�f�[�^�x�[�X�̃t�@�C����

Dim adoCn As Object
Set adoCn = CreateObject("ADODB.Connection") 'ADODB�R�l�N�V�����I�u�W�F�N�g���쐬
adoCn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFileName & ";" 'Access�t�@�C���ɐڑ�

Dim adoRs As Object 'ADO���R�[�h�Z�b�g�I�u�W�F�N�g
Set adoRs = CreateObject("ADODB.Recordset") 'ADO���R�[�h�Z�b�g�I�u�W�F�N�g���쐬

Dim WS_���� As Worksheet
    Set WS_���� = ThisWorkbook.Worksheets("����")

WS_����.Range("E8").Value = Now

C_ShUkNo = "*"
Dim strSQL As String
If C_ShUkNo = "*" Then
    strSQL = "SELECT * FROM �Č��ʖ{�ԉ����\�[�X�Ǘ�"
Else
    strSQL = "SELECT * FROM �Č��ʖ{�ԉ����\�[�X�Ǘ� where �Ǘ��p��tNo like " & """" & C_ShUkNo & "%" & """"
End If
 
adoRs.Open strSQL, adoCn    'SQL�����s���đΏۂ�RecordSet��

'
' �����o�����߂ɐV����book���쐬���A�V�[�g����"�S�������t�@�C��x"�Ƃ���
'
Workbooks.Add
With ActiveSheet
        .Name = "�S�������t�@�C��x"
    .Range("A1") = "No"
    .Range("B1") = "�W��p��tNo"
    .Range("C1") = "�}��"
    .Range("D1") = "�Ǘ��p��tNo"
    .Range("E1") = "VLOOKUP�L�["
    .Range("F1") = "���{�E�Ή��T�v"
    .Range("G1") = "���{�\���"
    .Range("H1") = "���\�[�X��"
    .Range("I1") = "���\�[�X���{�ꖼ"
    .Range("J1") = "���\�[�X�敪"
    .Range("K1") = "�敪"
    .Range("L1") = "�����敪"
    .Range("M1") = "�ˍ�������"
End With

ActiveSheet.Range("A2").CopyFromRecordset adoRs
 
adoRs.Close '���R�[�h�Z�b�g�̃N���[�Y
adoCn.Close '�R�l�N�V�����̃N���[�Y
 
Set adoRs = Nothing
Set adoCn = Nothing  '�I�u�W�F�N�g�̔j��
 
WS_����.Range("F8").Value = Now
MsgBox "�I�����܂����I"

End Sub
