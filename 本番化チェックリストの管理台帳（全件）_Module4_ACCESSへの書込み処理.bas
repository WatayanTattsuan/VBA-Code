Attribute VB_Name = "ACCESS�ւ̏����ݏ���"
Option Explicit

Sub TBL_Del_Add_SQL()

' ----------------------------------------------------------------------------------------------
'�@ACCESS�ւ̏����ݏ���
'�@�@�{�ԉ��`�F�b�N���X�g�̎�����S��DELETE & �S��ADD�@�œo�^����
' ----------------------------------------------------------------------------------------------

Const PRJ_NAME As String = "�����ʖ{�ԉ��o�[�W�����Ǘ�"
Dim WS As Worksheet

Set WS = Worksheets("�S�������t�@�C��")

Dim strFileName As String

Dim myPath As String
myPath = ThisWorkbook.Path & "\"
    
    strFileName = myPath & "�䒠�Ǘ�_2018.accdb" '�f�[�^�x�[�X�̃t�@�C����

Dim adoCn As Object
    Set adoCn = CreateObject("ADODB.Connection") 'ADODB�R�l�N�V�����I�u�W�F�N�g���쐬
    adoCn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFileName & ";" 'Access�t�@�C���ɐڑ�

Dim adoRs As Object 'ADO���R�[�h�Z�b�g�I�u�W�F�N�g
    Set adoRs = CreateObject("ADODB.Recordset") 'ADO���R�[�h�Z�b�g�I�u�W�F�N�g���쐬
 
Dim mySQL As String
Dim myRecordSet As New ADODB.Recordset
    mySQL = "DELETE * FROM " & PRJ_NAME
    adoRs.Open mySQL, adoCn
    
    mySQL = "SELECT * FROM " & PRJ_NAME & ";"
    adoRs.Open mySQL, adoCn, adOpenDynamic, adLockOptimistic
  
Dim i As Long
    i = 6
'    i = ActiveSheet.Range("A1").End(xlDown).Row
  
    Do While (WS.Cells(i, 1).Value <> "")
        Application.StatusBar = "row() = " & i
        adoRs.AddNew
            adoRs.Fields("��tNo") = WS.Cells(i, 1).Value
            adoRs.Fields("�o�[�W����") = WS.Cells(i, 2).Value
            adoRs.Fields("VLOOKUP�L�[") = WS.Cells(i, 3).Value
            adoRs.Fields("�`�F�b�N���X�g����") = WS.Cells(i, 4).Value
            adoRs.Fields("���\�[�X�����e����") = WS.Cells(i, 5).Value
            adoRs.Fields("��r") = WS.Cells(i, 6).Value
            adoRs.Fields("�قȂ��Ă���ӏ�") = WS.Cells(i, 7).Value
        adoRs.Update
        i = i + 1

    Loop

    adoRs.Close '���R�[�h�Z�b�g�̃N���[�Y
    adoCn.Close '�R�l�N�V�����̃N���[�Y
 
    Set adoRs = Nothing
    Set adoCn = Nothing  '�I�u�W�F�N�g�̔j��
 
    MsgBox "�I�����܂����I"

End Sub


Sub RTBL_Del_Add_SQL()

' ----------------------------------------------------------------------------------------------
'�@ACCESS�ւ̏����ݏ���
'�@�@���\�[�X�����e�ꗗ�̎�����S��DELETE & �S��ADD�@�œo�^����
' ----------------------------------------------------------------------------------------------

Const PRJ_NAME As String = "�Č��ʖ{�ԉ����\�[�X�Ǘ�"
Dim WS As Worksheet
Dim WS_���� As Worksheet

Set WS = Worksheets("���\�[�X�����e")
Set WS_���� = Worksheets("����")

    WS_����.Range("E4").Value = Now

Dim strFileName As String

Dim myPath As String
myPath = ThisWorkbook.Path & "\"
    
    strFileName = myPath & "�䒠�Ǘ�_2018.accdb" '�f�[�^�x�[�X�̃t�@�C����

Dim adoCn As Object
    Set adoCn = CreateObject("ADODB.Connection") 'ADODB�R�l�N�V�����I�u�W�F�N�g���쐬
    adoCn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFileName & ";" 'Access�t�@�C���ɐڑ�

Dim adoRs As Object 'ADO���R�[�h�Z�b�g�I�u�W�F�N�g
    Set adoRs = CreateObject("ADODB.Recordset") 'ADO���R�[�h�Z�b�g�I�u�W�F�N�g���쐬
 
Dim mySQL As String
Dim myRecordSet As New ADODB.Recordset
    mySQL = "DELETE * FROM " & PRJ_NAME
    adoRs.Open mySQL, adoCn
    
    mySQL = "SELECT * FROM " & PRJ_NAME & ";"
    adoRs.Open mySQL, adoCn, adOpenDynamic, adLockOptimistic
  
Dim i As Long
    i = 2
'    i = ActiveSheet.Range("A1").End(xlDown).Row
  
    Do While (WS.Cells(i, 2).Value <> "")
        Application.StatusBar = "row() = " & i
        adoRs.AddNew
            adoRs.Fields("No") = Format(WS.Cells(i, 1).Value, 0)
            adoRs.Fields("�W��p��tNo") = WS.Cells(i, 2).Value
            adoRs.Fields("�}��") = WS.Cells(i, 3).Value
            adoRs.Fields("�Ǘ��p��tNo") = WS.Cells(i, 4).Value
            adoRs.Fields("VLOOKUP�L�[") = WS.Cells(i, 5).Value
            adoRs.Fields("���{�E�Ή��T�v") = WS.Cells(i, 6).Value
            adoRs.Fields("���{�\���") = WS.Cells(i, 7).Value
            adoRs.Fields("���\�[�X��") = WS.Cells(i, 8).Value
            adoRs.Fields("���\�[�X���{�ꖼ") = WS.Cells(i, 9).Value
            adoRs.Fields("���\�[�X�敪") = WS.Cells(i, 10).Value
            adoRs.Fields("�敪") = WS.Cells(i, 11).Value
            adoRs.Fields("�����敪") = WS.Cells(i, 12).Value
            adoRs.Fields("�ˍ�������") = WS.Cells(i, 13).Value
        adoRs.Update
        i = i + 1

    Loop

    adoRs.Close '���R�[�h�Z�b�g�̃N���[�Y
    adoCn.Close '�R�l�N�V�����̃N���[�Y
 
    Set adoRs = Nothing
    Set adoCn = Nothing  '�I�u�W�F�N�g�̔j��
 
    WS_����.Range("F4").Value = Now
    MsgBox "�I�����܂����I"

End Sub


Sub �L�[�č쐬()

Dim i As Long
    i = 6
  
    Do While (Worksheets("�S�������t�@�C��").Range("A" & i).Value <> "")
        Application.StatusBar = "row() = " & i
        Worksheets("�S�������t�@�C��").Range("C" & i) = _
            Worksheets("�S�������t�@�C��").Range("A" & i) & Worksheets("�S�������t�@�C��").Range("D" & i)
        i = i + 1

    Loop

End Sub
