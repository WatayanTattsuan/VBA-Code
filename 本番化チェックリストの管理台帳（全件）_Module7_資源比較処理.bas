Attribute VB_Name = "������r����"
Option Explicit

Dim rsort As Long


Sub ������r_sub()

' ----------------------------------------------------------------------------------------------
'�@�����̓ˍ��������{���鏈��
'�@�@�{�ԉ��`�F�b�N���X�g�̎����ƃ��\�[�X�����e�̎������r���鏈��
' ----------------------------------------------------------------------------------------------

Dim C_ShUkNo As String
Dim strFileName As String
Dim myPath As String
myPath = ThisWorkbook.Path & "\"

Dim rIdx As Long

ActiveSheet.Range("C3:G10000").ClearContents
ActiveSheet.Range("I3:L10000").ClearContents

strFileName = myPath & "�䒠�Ǘ�_2018.accdb" '�f�[�^�x�[�X�̃t�@�C����
C_ShUkNo = ActiveSheet.Range("C1")

Dim adoCn As Object
Set adoCn = CreateObject("ADODB.Connection") 'ADODB�R�l�N�V�����I�u�W�F�N�g���쐬
adoCn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFileName & ";" 'Access�t�@�C���ɐڑ�

Dim adoRs As Object 'ADO���R�[�h�Z�b�g�I�u�W�F�N�g
Set adoRs = CreateObject("ADODB.Recordset") 'ADO���R�[�h�Z�b�g�I�u�W�F�N�g���쐬
 
adoRs.Open "�����ʖ{�ԉ��o�[�W�����Ǘ�", adoCn, adOpenDynamic, adLockOptimistic

rIdx = 3
adoRs.MoveFirst
adoRs.Filter = "��tNo LIKE '" & C_ShUkNo & "*'"
Do Until adoRs.EOF = True

    With ActiveSheet
        .Range("C" & rIdx).Value = adoRs!�`�F�b�N���X�g����
        .Range("D" & rIdx).Value = adoRs!��tNo
        .Range("E" & rIdx).Value = adoRs!�قȂ��Ă���ӏ�
        .Range("F" & rIdx).Value = adoRs!�o�[�W����
        .Range("G" & rIdx).Value = 999
    End With
    rIdx = rIdx + 1
    adoRs.MoveNext

Loop


Dim adoRRs As Object 'ADO���R�[�h�Z�b�g�I�u�W�F�N�g
Set adoRRs = CreateObject("ADODB.Recordset") 'ADO���R�[�h�Z�b�g�I�u�W�F�N�g���쐬
adoRRs.Open "�Č��ʖ{�ԉ����\�[�X�Ǘ�", adoCn, adOpenDynamic, adLockOptimistic

rIdx = 3
adoRRs.MoveFirst
adoRRs.Filter = "�Ǘ��p��tNo LIKE '" & C_ShUkNo & "*'"
Do Until adoRRs.EOF = True

    With ActiveSheet
        .Range("I" & rIdx).Value = adoRRs!���\�[�X��
        .Range("J" & rIdx).Value = adoRRs!�Ǘ��p��tNo
        .Range("K" & rIdx).Value = adoRRs!�����敪
        .Range("L" & rIdx).Value = 999
    End With
    rIdx = rIdx + 1
    adoRRs.MoveNext

Loop

adoRs.Close '���R�[�h�Z�b�g�̃N���[�Y
adoRRs.Close '���R�[�h�Z�b�g�̃N���[�Y
adoCn.Close '�R�l�N�V�����̃N���[�Y
 
rsort = ActiveSheet.Range("K2").End(xlDown).Row
Call ���\�[�X_SORT_sub

rIdx = 3
Dim FoundCell As Range
Dim Find_Rsrc As String
Dim bug_chk As Long
bug_chk = 0


Do While ActiveSheet.Range("C" & rIdx) <> ""
    
    Find_Rsrc = ActiveSheet.Range("C" & rIdx).Value
    Set FoundCell = ActiveSheet.Range("I:I").Find(Find_Rsrc, LookAt:=xlWhole)
    If FoundCell Is Nothing Then
        bug_chk = bug_chk + 1
    Else
        ActiveSheet.Range("G" & rIdx).Value = 0
    End If

    rIdx = rIdx + 1
    
Loop
ActiveSheet.Range("E1").Value = bug_chk

rIdx = 3
bug_chk = 0
Do While ActiveSheet.Range("I" & rIdx) <> ""
    
    Find_Rsrc = ActiveSheet.Range("I" & rIdx).Value
    Set FoundCell = ActiveSheet.Range("C:C").Find(Find_Rsrc, LookAt:=xlWhole)
    If FoundCell Is Nothing Then
        bug_chk = bug_chk + 1
    Else
        ActiveSheet.Range("L" & rIdx).Value = 0
    End If

    rIdx = rIdx + 1
    
Loop

ActiveSheet.Range("K1").Value = bug_chk
Call ���\�[�X_SORT_sub

Set adoRs = Nothing
Set adoRRs = Nothing
Set adoCn = Nothing  '�I�u�W�F�N�g�̔j��
 
MsgBox "�I�����܂����I"

End Sub

Sub ���\�[�X_SORT_sub()
    
    ActiveSheet.Range("F1").Value = rsort
    With ActiveSheet.Range("C3:G" & rsort)
        .Sort Key1:=Range("G3"), order1:=xlAscending, _
        Key2:=Range("D3"), order2:=xlDescending, _
        Key3:=Range("C3"), order3:=xlDescending, _
        Header:=xlYes
    End With

    ActiveSheet.Range("I1").Value = rsort
    With ActiveSheet.Range("I3:M" & rsort)
        .Sort Key1:=Range("L3"), order1:=xlAscending, _
        Key2:=Range("J3"), order2:=xlDescending, _
        Key3:=Range("I3"), order3:=xlDescending, _
        Header:=xlYes
    End With

End Sub




