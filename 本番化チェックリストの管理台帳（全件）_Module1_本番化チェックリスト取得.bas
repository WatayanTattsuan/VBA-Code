Attribute VB_Name = "�{�ԉ��`�F�b�N���X�g�擾"
Option Explicit

Const myPath As String = "D:\SVN\�{�ԉ�\"
Dim rIdx As Long
Dim rIdy As Long
Dim rIdz As Long
Dim fName As String
Dim attach_version As String
Dim attach_date As Date
Dim attach_matter As String
Dim WS As New Worksheet
Dim WS_LOG As New Worksheet

Sub VgetALL_F()

Set WS = ThisWorkbook.Worksheets("�����S���f")
Set WS_LOG = ThisWorkbook.Worksheets("����")

WS_LOG.Range("E7").Value = Now
' ----------------------------------------------------------------------------------------------
'�@�����擾����
'�@�@�{�ԉ��`�F�b�N���X�g���玑�����擾����
' ----------------------------------------------------------------------------------------------

fName = Dir(myPath & "*.xls")
rIdx = 1
rIdz = 2

Call VgetAttachSUB_F        '�{�ԉ��o�[�W�����Ǘ��䒠�쐬�i�Ǘ�No�j
Call VreplSUB_F             '��A����i�u���j

WS_LOG.Range("F7").Value = Now
MsgBox "�I�����܂����I"

End Sub

Sub VgetAttachSUB_F()

WS.Activate
rIdx = WS.Cells(1, 1).End(xlDown).Row

WS.Range("A" & rIdx).Select    '�w��̈ʒu�ɃZ����u��

Do Until fName = ""
    
    Workbooks.Open Filename:=myPath & fName
    Worksheets("�\��").Activate
    attach_version = ActiveSheet.Cells(18, 4).Value
    attach_date = ActiveSheet.Cells(22, 4).Value
    attach_matter = ActiveSheet.Cells(17, 4).Value
    Workbooks(fName).Worksheets("�����ꗗ").Activate

    '********** �����ꗗ����f�[�^�𒊏o���Ă��� *********************************
    
    rIdy = 3
    Do While Workbooks(fName).ActiveSheet.Cells(rIdy, 1) <> ""
            
        If Len(Workbooks(fName).ActiveSheet.Cells(rIdy, 7).Value) > 5 Then
            
            rIdx = rIdx + 1

            With WS
                .Cells(rIdx, 1).Value = ActiveSheet.Cells(rIdy, 7).Value
                .Cells(rIdx, 2).Value = attach_version
                .Cells(rIdx, 3).Value = attach_date
                .Cells(rIdx, 4).Formula = "=workday(C" & rIdx & ",1,�j���ݒ�!$A$2:$A$1000)"
                .Cells(rIdx, 5).Value = attach_matter
                .Cells(rIdx, 9).Value = fName
                .Cells(rIdx, 10).Value = Workbooks(fName).ActiveSheet.Name
                .Cells(rIdx, 11).Value = ActiveSheet.Cells(rIdy, 2).Value
            End With
            
            If Mid(WS.Cells(rIdx, 1).Value, 1, 5) = "PRISM" Then
                WS.Cells(rIdx, 6).Value = ActiveSheet.Cells(rIdy, 9).Value
            ElseIf Mid(WS.Cells(rIdx, 1).Value, 1, 5) = "ASTRA" Then
                WS.Cells(rIdx, 7).Value = ActiveSheet.Cells(rIdy, 9).Value
            ElseIf Mid(WS.Cells(rIdx, 1).Value, 1, 5) = "JINJI" Then
                WS.Cells(rIdx, 7).Value = ActiveSheet.Cells(rIdy, 9).Value
            Else
                WS.Cells(rIdx, 8).Value = ActiveSheet.Cells(rIdy, 9).Value
            End If
            
        End If
        rIdy = rIdy + 1
    Loop

    '********** DB���f�ꗗ����f�[�^�𒊏o���Ă��� *********************************

    Workbooks(fName).Worksheets("DB���f�ꗗ").Activate
    
    rIdy = 3
    Do While Workbooks(fName).ActiveSheet.Cells(rIdy, 1) <> ""
        If Len(Workbooks(fName).ActiveSheet.Cells(rIdy, 5).Value) > 5 Then
            
            rIdx = rIdx + 1
            With WS
                .Cells(rIdx, 1).Value = ActiveSheet.Cells(rIdy, 5).Value
                .Cells(rIdx, 2).Value = attach_version
                .Cells(rIdx, 3).Value = attach_date
                .Cells(rIdx, 4).Formula = "=workday(C" & rIdx & ",1,�j���ݒ�!$A$2:$A$1000)"
                .Cells(rIdx, 5).Value = attach_matter
                .Cells(rIdx, 9).Value = fName
                .Cells(rIdx, 10).Value = Workbooks(fName).ActiveSheet.Name
                .Cells(rIdx, 11).Value = ActiveSheet.Cells(rIdy, 2).Value
            End With
            
            If Mid(WS.Cells(rIdx, 1).Value, 1, 5) = "PRISM" Then
                WS.Cells(rIdx, 6).Value = ActiveSheet.Cells(rIdy, 3).Value
            ElseIf Mid(WS.Cells(rIdx, 1).Value, 1, 5) = "ASTRA" Then
                WS.Cells(rIdx, 7).Value = ActiveSheet.Cells(rIdy, 3).Value
            ElseIf Mid(WS.Cells(rIdx, 1).Value, 1, 5) = "JINJI" Then
                WS.Cells(rIdx, 7).Value = ActiveSheet.Cells(rIdy, 3).Value
            Else
                WS.Cells(rIdx, 8).Value = ActiveSheet.Cells(rIdy, 3).Value
            End If
            
        End If
        rIdy = rIdy + 1
              
    Loop
    
    Windows(fName).Close
    fName = Dir

Loop
    
End Sub
    
Sub VreplSUB_F()

    '********** �u��
    WS.Activate
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





