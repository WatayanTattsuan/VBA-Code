Attribute VB_Name = "�ˍ����{�^������"
Option Explicit

Sub ���\�[�X�����e_�ˍ����{�^��_Click()

' ----------------------------------------------------------------------------------------------
'�@�ˍ�������
'�@�@���\�[�X�����e������{�ԉ��`�F�b�N���X�g�̎����Ɠ˂����킹��
' ----------------------------------------------------------------------------------------------

Dim rIdx As Long
Dim F_SRC As String
Dim FoundCell As Range
Dim FoundROW As Long
Dim FoundCount As Long
Const PRJ_NAME As String = "�����ʖ{�ԉ��o�[�W�����Ǘ�"
Dim strFileName As String

' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' ACCESS���쏀������
' ____________________________________________

Dim myPath As String
myPath = ThisWorkbook.Path & "\"

Dim adoCn As Object
    strFileName = myPath & "�䒠�Ǘ�_2018.accdb" '�f�[�^�x�[�X�̃t�@�C����
    Set adoCn = CreateObject("ADODB.Connection") 'ADODB�R�l�N�V�����I�u�W�F�N�g���쐬
    adoCn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFileName & ";" 'Access�t�@�C���ɐڑ�

Dim adoRs As Object 'ADO���R�[�h�Z�b�g�I�u�W�F�N�g
    Set adoRs = CreateObject("ADODB.Recordset") 'ADO���R�[�h�Z�b�g�I�u�W�F�N�g���쐬
 
Dim mySQL As String
Dim myRecordSet As New ADODB.Recordset
Dim C_adoFind As String
    
' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' ���O�����o��
' ____________________________________________

Dim WS_���� As Worksheet
    Set WS_���� = ThisWorkbook.Worksheets("����")

    With WS_����
        .Range("E3").Value = Now
        .Range("Q1").Value = 0
        .Range("Q2").Value = 0
        .Range("R1").Value = "�ˍ����i���\�[�X�����e)"
        .Range("R2").Value = "�ˍ����i���\�[�X�����e)"
    End With
     
adoRs.Open PRJ_NAME, adoCn, adOpenDynamic, adLockOptimistic
    
Worksheets("���\�[�X�����e").Activate
rIdx = 2
FoundCount = 0

' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' �����̓ˍ�����<ACCESS������>
' ____________________________________________

Do While Worksheets("���\�[�X�����e").Range("B" & rIdx) <> ""

    If ActiveSheet.Range("M" & rIdx) <> "" Then
        GoTo Bypass
    End If
    Application.StatusBar = "row() = " & rIdx

    F_SRC = Worksheets("���\�[�X�����e").Range("E" & rIdx)
    C_adoFind = "VLOOKUP�L�[ = '" & F_SRC & "'"
    adoRs.Find C_adoFind
    If adoRs.BOF = True Then
        WS_����.Range("Q1").Value = WS_����.Range("Q1").Value + 1
        GoTo Bypass
    End If
    If adoRs.EOF = True Then
        WS_����.Range("Q2").Value = WS_����.Range("Q2").Value + 1
        GoTo Bypass
    End If

    Worksheets("���\�[�X�����e").Range("M" & rIdx).Value = adoRs!�`�F�b�N���X�g����
    FoundCount = FoundCount + 1
    
Bypass:
    rIdx = rIdx + 1
        
    adoRs.MoveFirst

Loop

WS_����.Range("F3").Value = Now
MsgBox "����ˍ������� �F " & FoundCount & " ���ł���"

End Sub

Sub �`�F�b�N���X�g_�ˍ����{�^��_Click()

' ----------------------------------------------------------------------------------------------
'�@�ˍ�������
'�@�@�{�ԉ��`�F�b�N���X�g�̎��������\�[�X�����e�����Ǝ����Ɠ˂����킹��
' ----------------------------------------------------------------------------------------------

Dim rIdx As Long
Dim F_SRC As String
Dim FoundCell As Range
Dim FoundROW As Long
Dim FoundCount As Long
Const PRJ_NAME As String = "�Č��ʖ{�ԉ����\�[�X�Ǘ�"
Dim strFileName As String

' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' ACCESS���쏀������
' ____________________________________________

Dim myPath As String
myPath = ThisWorkbook.Path & "\"

Dim adoCn As Object
    strFileName = myPath & "�䒠�Ǘ�_2018.accdb" '�f�[�^�x�[�X�̃t�@�C����
    Set adoCn = CreateObject("ADODB.Connection") 'ADODB�R�l�N�V�����I�u�W�F�N�g���쐬
    adoCn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFileName & ";" 'Access�t�@�C���ɐڑ�

Dim adoRs As Object 'ADO���R�[�h�Z�b�g�I�u�W�F�N�g
    Set adoRs = CreateObject("ADODB.Recordset") 'ADO���R�[�h�Z�b�g�I�u�W�F�N�g���쐬
 
Dim mySQL As String
Dim myRecordSet As New ADODB.Recordset
Dim C_adoFind As String
    
' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' ���O�����o��
' ____________________________________________

Dim WS_���� As Worksheet
Dim WS_�`�F�b�N���X�g As Worksheet
    Set WS_���� = ThisWorkbook.Worksheets("����")
    Set WS_�`�F�b�N���X�g = ThisWorkbook.Worksheets("�S�������t�@�C��")

    WS_����.Range("E5").Value = Now
    WS_����.Range("Q1").Value = 0
    WS_����.Range("Q2").Value = 0
    WS_����.Range("R1").Value = "�ˍ����i�`�F�b�N���X�g)"
    WS_����.Range("R2").Value = "�ˍ����i�`�F�b�N���X�g)"
        
        
        
adoRs.Open PRJ_NAME, adoCn, adOpenDynamic, adLockOptimistic
    
WS_�`�F�b�N���X�g.Activate
rIdx = 6
FoundCount = 0

' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
' �����̓ˍ�����<ACCESS������>
' ____________________________________________

Do While WS_�`�F�b�N���X�g.Range("B" & rIdx) <> ""

    If ActiveSheet.Range("E" & rIdx) <> "" Then
        GoTo Bypass
    End If
    
'    Application.StatusBar = "row() = " & rIdx

    F_SRC = WS_�`�F�b�N���X�g.Range("C" & rIdx)
    C_adoFind = "VLOOKUP�L�[ = '" & F_SRC & "'"
    adoRs.Find C_adoFind
    If adoRs.BOF = True Then
        WS_����.Range("Q1").Value = WS_����.Range("Q1").Value + 1
        GoTo Bypass
    End If
    If adoRs.EOF = True Then
        WS_����.Range("Q2").Value = WS_����.Range("Q2").Value + 1
        GoTo Bypass
    End If

    WS_�`�F�b�N���X�g.Range("E" & rIdx).Value = adoRs!���\�[�X��
    WS_�`�F�b�N���X�g.Range("H" & rIdx).Value = Now
    
    FoundCount = FoundCount + 1
    
Bypass:
    rIdx = rIdx + 1
        
    adoRs.MoveFirst

Loop

WS_����.Range("F5").Value = Now
MsgBox "����ˍ������� �F " & FoundCount & " ���ł���"

End Sub
