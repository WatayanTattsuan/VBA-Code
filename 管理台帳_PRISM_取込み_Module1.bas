Attribute VB_Name = "Module1"
Option Explicit

Const myPath As String = "D:\SVN\�Ǘ��䒠\�Ǘ��䒠_2018\"
Dim Skip_Count As Long
Dim Indx1 As Long

Sub �Ǘ��䒠_PRISM_MAIN_F()
'---------------------------------------------------------------------------------------
'�Ǘ��䒠�Ɋ֌W������̎捞�ݏ���
'
'---------------------------------------------------------------------------------------

MsgBox "�J�n���܂��I"

ThisWorkbook.Worksheets("LOG").Range("B2").Value = Now
Call DgetB_F          '�e�Ǘ��䒠����󋵂̎捞
Call DgetA_F          '�`�F�b�N
Call DputD_F          'TEMP�t�@�C���ɔ��f
Call TBL_Del_Add_SQL
ThisWorkbook.Worksheets("LOG").Range("C2").Value = Now

ThisWorkbook.Worksheets("�Ǘ��䒠_PRISM").Cells(2, 6).Value = Now

Skip_Count = Indx1 - 6 - Skip_Count
MsgBox "�I�����܂����I" & vbLf & "�X�L�b�v�J�E���g�@�F�@" & Skip_Count

End Sub



Sub DgetA_F()
'
' Macro_check1 Macro
'

Dim IraiNO As Long
Dim Irai1 As Long
Dim Irai2 As Long
Dim Shub1 As Long
Dim UkeNo As Long
Dim Test1 As Long
Dim Test2 As Long
Dim Test3 As Long
Dim Atta1 As Long
Dim Atta2 As Long
Dim Atta3 As Long
Dim Atta4 As Long
Dim Atta5 As Long
Dim Atta6 As Long
Dim HokoNO As Long
Dim Hoko1 As Long
Dim Hoko2 As Long
Dim Hoko3 As Long
Dim Hoko4 As Long
Dim ShUkNo As Long
Dim KnUkNo As Long

Dim Flag1 As Long

Shub1 = 13     '�񖼁F�������
UkeNo = 17     '�񖼁F��tNo&�A��
ShUkNo = 40    '�W��p��tNo
KnUkNo = 42    '�Ǘ��p��tNo
IraiNO = 45    '�˗��������ԍ�2
HokoNO = 46    '�񍐏������ԍ�2
Hoko3 = 67     '�񖼁F�񍐏��������F
Hoko4 = 68     '�񖼁F�񍐏��������F��
Indx1 = 6
Skip_Count = 0

ThisWorkbook.Sheets("�Ǘ��䒠_PRISM").Activate
ActiveSheet.Range("BK6").Select

If ActiveSheet.Range("A1").Value < 6 Then
    MsgBox "�l���Ԉ���Ă��܂�"
Else: Indx1 = ActiveSheet.Range("A1").Value
End If

Do While ActiveSheet.Cells(Indx1, 37) <> "-" And ActiveSheet.Range("A2").Value > Indx1

    On Error Resume Next
    
    If (ActiveSheet.Cells(Indx1, 37).Value = "��" Or ActiveSheet.Cells(Indx1, 37).Value = "��" Or ActiveSheet.Cells(Indx1, 37).Value = "�~" Or ActiveSheet.Cells(Indx1, 37).Value = "=") And (ActiveSheet.Cells(Indx1, 11).Value = "��" Or ActiveSheet.Cells(Indx1, 11).Value = "��" Or ActiveSheet.Cells(Indx1, 11).Value = "�~") And ActiveSheet.Cells(Indx1, 68).Value <> "" Then
        GoTo bypass
    End If
    
    ActiveSheet.Cells(Indx1, IraiNO).Value = 0
    ActiveSheet.Cells(Indx1, IraiNO).Value = _
            Worksheets("�䒠�Ǘ�").Cells(WorksheetFunction.Match(ActiveSheet.Cells(Indx1, ShUkNo), Worksheets("�䒠�Ǘ�").Range("O:O"), 0), 3).Value
    ActiveSheet.Cells(Indx1, HokoNO).Value = 0
    ActiveSheet.Cells(Indx1, HokoNO).Value = _
            Worksheets("�䒠�Ǘ�").Cells(WorksheetFunction.Match(ActiveSheet.Cells(Indx1, ShUkNo), Worksheets("�䒠�Ǘ�").Range("O:O"), 0), 4).Value
    
    
    If ActiveSheet.Cells(Indx1, 30).Value = "-" Or ActiveSheet.Cells(Indx1, 30).Value = "" Then
        ActiveSheet.Cells(Indx1, 47).Value = 0
    Else
        ActiveSheet.Cells(Indx1, 47).Value = ActiveSheet.Cells(Indx1, 30).Value
    End If
    If ActiveSheet.Cells(Indx1, 31).Value = "-" Or ActiveSheet.Cells(Indx1, 31).Value = "" Then
        ActiveSheet.Cells(Indx1, 48).Value = 0
    Else
        ActiveSheet.Cells(Indx1, 48).Value = ActiveSheet.Cells(Indx1, 31).Value
    End If
    If ActiveSheet.Cells(Indx1, 32).Value = "-" Or ActiveSheet.Cells(Indx1, 32).Value = "" Then
        ActiveSheet.Cells(Indx1, 49).Value = 0
    Else
        ActiveSheet.Cells(Indx1, 49).Value = ActiveSheet.Cells(Indx1, 32).Value
    End If
    
    
    
'    If (ActiveSheet.Cells(Indx1, 50).Value = "*" Or ActiveSheet.Cells(Indx1, 52).Value = "*" Or ActiveSheet.Cells(Indx1, 54).Value = "*") Then
'        ActiveSheet.Cells(Indx1, Hoko3).Value = ""
'        ActiveSheet.Cells(Indx1, Hoko3).Value = _
'            Worksheets("�䒠�Ǘ�").Cells(WorksheetFunction.Match(ActiveSheet.Cells(Indx1, ShUkNo), Worksheets("�䒠�Ǘ�").Range("O:O"), 0), 34).Value
'        ActiveSheet.Cells(Indx1, Hoko4).Value = ""
'        ActiveSheet.Cells(Indx1, Hoko4).Value = _
'            Worksheets("�䒠�Ǘ�").Cells(WorksheetFunction.Match(ActiveSheet.Cells(Indx1, ShUkNo), Worksheets("�䒠�Ǘ�").Range("O:O"), 0), 35).Value
'    ElseIf (ActiveSheet.Cells(Indx1, 37).Value = "��" Or ActiveSheet.Cells(Indx1, 37).Value = "��" Or ActiveSheet.Cells(Indx1, 37).Value = "�~") Then
'        ActiveSheet.Cells(Indx1, Hoko3).Value = ""
'        ActiveSheet.Cells(Indx1, Hoko3).Value = ActiveSheet.Cells(Indx1, 32).Value
'        ActiveSheet.Cells(Indx1, Hoko4).Value = ""
'        ActiveSheet.Cells(Indx1, Hoko4).Value = "�f��"
'    End If

'   ---------------------------------------------
'   �����ς̌������J�E���g�A�b�v
    Skip_Count = Skip_Count + 1
    
'   �䒠�Ǘ�����u�x���v�E�u�v���Ӂv�̂��̂��Ǘ��䒠�ɃZ�b�g����
'    ActiveSheet.Cells(Indx1, 66).Value = ""
'    ActiveSheet.Cells(Indx1, 66).Value = _
'            ActiveSheet.Cells(Indx1, 39).Value
'    ActiveSheet.Cells(Indx1, 66).Value = _
'            ActiveSheet.Cells(Indx1, 66).Value & _
'            vbLf & _
'            Worksheets("�䒠�Ǘ�").Cells(WorksheetFunction.Match(ActiveSheet.Cells(Indx1, KnUkNo), Worksheets("�䒠�Ǘ�").Range("Q:Q"), 0), 36).Value
'            Worksheets("�䒠�Ǘ�").Cells(WorksheetFunction.Match(ActiveSheet.Cells(Indx1, KnUkNo), Worksheets("�䒠�Ǘ�").Range("Q:Q"), 0), 37).Value & _
'            vbLf & _

'   �Ǘ��䒠���u���v�Ȃ�u�~�v�E�u���v�ɕύX���Z�b�g����
'    If InStr(ActiveSheet.Cells(Indx1, 66), "�x��") > 0 And (ActiveSheet.Cells(Indx1, 37) = "��" Or ActiveSheet.Cells(Indx1, 37) = "��") Then
'        ActiveSheet.Cells(Indx1, 37) = "�~"
'    ElseIf InStr(ActiveSheet.Cells(Indx1, 66), "�v����") > 0 And ActiveSheet.Cells(Indx1, 37) = "��" Then
'        ActiveSheet.Cells(Indx1, 37) = "��"
'    Else
'    End If
'   ---------------------------------------------

bypass:
    
    On Error GoTo 0

    Indx1 = Indx1 + 1
    Application.StatusBar = "CHK-" & Indx1 - 6

Loop

DgetA_E:

End Sub

Sub DgetB_F()

Const DAICHO As String = "�Č��W��.xlsx"
Const Version As String = "�{�ԉ��`�F�b�N���X�g�̊Ǘ��䒠.xlsm"

ThisWorkbook.Worksheets("�䒠�Ǘ�").Activate
ActiveSheet.Cells.Clear
ActiveSheet.Range("A1").Select    '�w��̈ʒu�ɃZ����u��

'-- PRISM�̊Ǘ��䒠����荞�� --

Workbooks.Open Filename:=myPath & DAICHO
Workbooks(DAICHO).Worksheets("�䒠�Ǘ�").Activate
 
Workbooks(DAICHO).Worksheets("�䒠�Ǘ�").Range("A:AQ").Copy
ThisWorkbook.Worksheets("�䒠�Ǘ�").Range("A1").PasteSpecial Paste:=xlPasteAll
Application.CutCopyMode = False

'-- �o�[�W�������(�Ǘ�No.)����荞�� --

Workbooks.Open Filename:=myPath & Version
 
Workbooks(Version).Worksheets("�{�ԉ��`�F�b�N���X�g�䒠(�Ǘ�No)").Range("A:I").Copy
ThisWorkbook.Worksheets("�Ǘ�No").Range("A1").PasteSpecial Paste:=xlPasteAll
Application.CutCopyMode = False

Workbooks(Version).Close SaveChanges:=False, Filename:=Version

End Sub


Sub DputD_F()

Dim END_CELL1 As Integer
Dim END_CELL2 As Integer

Const TempA As String = "�Č��W��.xlsx"

Workbooks(TempA).Worksheets("�Ǘ��䒠_PRISM").Range("A6:BM1000").ClearContents
ThisWorkbook.Worksheets("�Ǘ��䒠_PRISM").Range("A6:BM1000").Copy


Workbooks(TempA).Worksheets("�Ǘ��䒠_PRISM").Range("A6").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
Application.CutCopyMode = False



    END_CELL1 = ThisWorkbook.Worksheets("�Ǘ��䒠_PRISM").Cells(6, 1).End(xlDown).Row
    ThisWorkbook.Worksheets("�Ǘ��䒠_PRISM").Range("AN6", "AW" & END_CELL1).Copy

    ThisWorkbook.Worksheets("ACCESS").Range("A2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False

'    ThisWorkbook.Worksheets("�Ǘ��䒠_PRISM").Range("AD6", "AF" & END_CELL1).Copy

'    ThisWorkbook.Worksheets("ACCESS").Range("H2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
'    Application.CutCopyMode = False

'----------

    END_CELL2 = Workbooks(TempA).Worksheets("PRISM_ACCESS").Cells(1, 1).End(xlDown).Row
    Workbooks(TempA).Worksheets("PRISM_ACCESS").Range("A2:J" & END_CELL2).ClearContents

    ThisWorkbook.Worksheets("ACCESS").Range("A2:J" & END_CELL1).Copy
    Workbooks(TempA).Worksheets("PRISM_ACCESS").Range("A2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False



Workbooks(TempA).Save
Workbooks(TempA).Close SaveChanges:=False, Filename:=TempA

ThisWorkbook.Worksheets("�i���m�F").Range("A6:BM1000").ClearContents
ThisWorkbook.Worksheets("�Ǘ��䒠_PRISM").Range("A6:BM1000").Copy

ThisWorkbook.Worksheets("�i���m�F").Range("A6").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False


End Sub




Sub CHECK_MAIN()

MsgBox "�J�n���܂��I"

ThisWorkbook.Worksheets("LOG").Range("B3").Value = Now
Call DcheckB_F
ThisWorkbook.Worksheets("LOG").Range("C3").Value = Now

Skip_Count = Indx1 - 6 - Skip_Count
MsgBox "�I�����܂����I" & vbLf & "�X�L�b�v�J�E���g�@�F�@" & Skip_Count

End Sub


Sub DcheckB_F()
'
' Macro_check1 Macro
'

Dim IraiNO As Long
Dim Irai1 As Long
Dim Irai2 As Long
Dim Shub1 As Long
Dim UkeNo As Long
Dim Test1 As Long
Dim Test2 As Long
Dim Test3 As Long
Dim Atta1 As Long
Dim Atta2 As Long
Dim Atta3 As Long
Dim Atta4 As Long
Dim Atta5 As Long
Dim Atta6 As Long
Dim HokoNO As Long
Dim Hoko1 As Long
Dim Hoko2 As Long
Dim Hoko3 As Long
Dim Hoko4 As Long
Dim ShUkNo As Long
Dim KnUkNo As Long

Dim Flag1 As Long

Shub1 = 13     '�񖼁F�������
UkeNo = 17     '�񖼁F��tNo&�A��
ShUkNo = 40    '�W��p��tNo
KnUkNo = 42    '�Ǘ��p��tNo
IraiNO = 45    '�˗��������ԍ�2
HokoNO = 46    '�񍐏������ԍ�2
Hoko3 = 67     '�񖼁F�񍐏��������F
Hoko4 = 68     '�񖼁F�񍐏��������F��
Indx1 = 6
Skip_Count = 0

ThisWorkbook.Sheets("�Ǘ��䒠_PRISM").Activate
ActiveSheet.Range("BK6").Select

If ActiveSheet.Range("A1").Value < 6 Then
    MsgBox "�l���Ԉ���Ă��܂�"
Else: Indx1 = ActiveSheet.Range("A1").Value
End If

Do While ActiveSheet.Cells(Indx1, 37) <> "-" And ActiveSheet.Range("A2").Value > Indx1

    On Error Resume Next
    
'    If (ActiveSheet.Cells(Indx1, 37).Value = "��" Or ActiveSheet.Cells(Indx1, 37).Value = "��" Or ActiveSheet.Cells(Indx1, 37).Value = "�~" Or ActiveSheet.Cells(Indx1, 37).Value = "=") And (ActiveSheet.Cells(Indx1, 11).Value = "��" Or ActiveSheet.Cells(Indx1, 11).Value = "��" Or ActiveSheet.Cells(Indx1, 11).Value = "�~") And ActiveSheet.Cells(Indx1, 68).Value <> "" Then
'        GoTo bypass
'    End If
    
    If (ActiveSheet.Cells(Indx1, 50).Value = "*" Or ActiveSheet.Cells(Indx1, 52).Value = "*" Or ActiveSheet.Cells(Indx1, 54).Value = "*") Then
        ActiveSheet.Cells(Indx1, Hoko3).Value = ""
        ActiveSheet.Cells(Indx1, Hoko3).Value = _
            Worksheets("�䒠�Ǘ�").Cells(WorksheetFunction.Match(ActiveSheet.Cells(Indx1, ShUkNo), Worksheets("�䒠�Ǘ�").Range("O:O"), 0), 34).Value
        ActiveSheet.Cells(Indx1, Hoko4).Value = ""
        ActiveSheet.Cells(Indx1, Hoko4).Value = _
            Worksheets("�䒠�Ǘ�").Cells(WorksheetFunction.Match(ActiveSheet.Cells(Indx1, ShUkNo), Worksheets("�䒠�Ǘ�").Range("O:O"), 0), 35).Value
    ElseIf (ActiveSheet.Cells(Indx1, 37).Value = "��" Or ActiveSheet.Cells(Indx1, 37).Value = "��" Or ActiveSheet.Cells(Indx1, 37).Value = "�~") Then
        ActiveSheet.Cells(Indx1, Hoko3).Value = ""
        ActiveSheet.Cells(Indx1, Hoko3).Value = ActiveSheet.Cells(Indx1, 32).Value
        ActiveSheet.Cells(Indx1, Hoko4).Value = ""
        ActiveSheet.Cells(Indx1, Hoko4).Value = "�f��"
    End If

'   ---------------------------------------------
'   �����ς̌������J�E���g�A�b�v
    Skip_Count = Skip_Count + 1
    
'   �䒠�Ǘ�����u�x���v�E�u�v���Ӂv�̂��̂��Ǘ��䒠�ɃZ�b�g����
    ActiveSheet.Cells(Indx1, 66).Value = ""
    ActiveSheet.Cells(Indx1, 66).Value = _
            ActiveSheet.Cells(Indx1, 39).Value
    ActiveSheet.Cells(Indx1, 66).Value = _
            ActiveSheet.Cells(Indx1, 66).Value & _
            vbLf & _
            Worksheets("�䒠�Ǘ�").Cells(WorksheetFunction.Match(ActiveSheet.Cells(Indx1, KnUkNo), Worksheets("�䒠�Ǘ�").Range("Q:Q"), 0), 36).Value

'   �Ǘ��䒠���u���v�Ȃ�u�~�v�E�u���v�ɕύX���Z�b�g����
    If InStr(ActiveSheet.Cells(Indx1, 66), "�x��") > 0 And (ActiveSheet.Cells(Indx1, 37) = "��" Or ActiveSheet.Cells(Indx1, 37) = "��") Then
        ActiveSheet.Cells(Indx1, 37) = "�~"
    ElseIf InStr(ActiveSheet.Cells(Indx1, 66), "�v����") > 0 And ActiveSheet.Cells(Indx1, 37) = "��" Then
        ActiveSheet.Cells(Indx1, 37) = "��"
    Else
    End If
'   ---------------------------------------------

bypass:
    
    On Error GoTo 0

    Indx1 = Indx1 + 1
    Application.StatusBar = "CHK-" & Indx1 - 6

Loop

End Sub


