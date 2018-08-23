Attribute VB_Name = "�O���t�`�揈��1"
Option Explicit

Sub MG_��������()
'---------------------------------------------------------------------------------------
'�O���[�v���̔����������O���t�ɂ��鏈��
'
'---------------------------------------------------------------------------------------
'
    Call MG_��������1
    Call MG_��������2
    Call MG_��������3
    Call MG_��������4
    Call MG_��������5
        
End Sub


Sub MG_��������1()
Attribute MG_��������1.VB_ProcData.VB_Invoke_Func = " \n14"
'---------------------------------------------------------------------------------------
'�O���[�v���̔����������O���t�ɂ��鏈��
'�@�Č��W�񂩂�O���t���쐬����
'---------------------------------------------------------------------------------------
'
    Dim chart1 As Chart
    Set chart1 = Charts.Add(after:=ActiveSheet)
    ActiveChart.ChartType = xlColumnClustered
    
    '�f�[�^�͈͐ݒ�
    ActiveChart.SetSourceData Source:=Range("�W�v���!$C$47:$C$59,�W�v���!$D$47:$D$59,�W�v���!$R$47:$R$59,�W�v���!$E$47:$E$59,�W�v���!$EE$17:$EE$28,�W�v���!$F$47:$F$59,�W�v���!$T$47:$T$59")
    ActiveChart.PlotArea.Select
    ActiveChart.SeriesCollection(1).Name = "=�W�v���!$C$34"
    ActiveChart.SeriesCollection(1).Select
    
    ActiveChart.SeriesCollection(2).Name = "=�W�v���!$D$33"
    ActiveChart.SeriesCollection(2).Select
    ActiveChart.SeriesCollection(2).AxisGroup = 2
       
    ActiveChart.SeriesCollection(3).Name = "=�W�v���!$R$33"
    ActiveChart.SeriesCollection(3).Select
    ActiveChart.SeriesCollection(3).AxisGroup = 2
       
    ActiveChart.SeriesCollection(4).Name = "=�W�v���!$E$33"
    ActiveChart.SeriesCollection(4).Select
    ActiveChart.SeriesCollection(4).AxisGroup = 2
    
       
    ActiveChart.SeriesCollection(5).Name = "=�W�v���!$S$33"
    ActiveChart.SeriesCollection(5).Select
    ActiveChart.SeriesCollection(5).AxisGroup = 2
       
    ActiveChart.SeriesCollection(6).Name = "=�W�v���!$F$33"
    ActiveChart.SeriesCollection(6).Select
    ActiveChart.SeriesCollection(6).AxisGroup = 2
           
    ActiveChart.SeriesCollection(7).Name = "=�W�v���!$T$33"
    ActiveChart.SeriesCollection(7).Select
    ActiveChart.SeriesCollection(7).AxisGroup = 2
    
    'X���������ݒ�
    ActiveChart.SeriesCollection(1).XValues = "=�W�v���!$B$47:$B$58"
    

End Sub
Sub MG_��������2()
Attribute MG_��������2.VB_ProcData.VB_Invoke_Func = " \n14"
'---------------------------------------------------------------------------------------
'�O���[�v���̔����������O���t�ɂ��鏈��
'�@�S�����������̏�����ύX���đ̍ق𐮂���
'---------------------------------------------------------------------------------------
'
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.ChartGroups(1).Overlap = 0
    ActiveChart.ChartGroups(1).GapWidth = 0
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Solid
    End With
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.400000006
        .Transparency = 0.650000006
        .Solid
    End With

End Sub
Sub MG_��������3()
Attribute MG_��������3.VB_ProcData.VB_Invoke_Func = " \n14"
'---------------------------------------------------------------------------------------
'�O���[�v���̔����������O���t�ɂ��鏈��
'�@�c���̏���l��ύX����
'---------------------------------------------------------------------------------------
'
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MaximumScale = 200
    ActiveChart.Axes(xlValue, xlSecondary).MaximumScale = 200

End Sub
Sub MG_��������4()
'---------------------------------------------------------------------------------------
'�O���[�v���̔����������O���t�ɂ��鏈��
'�@�����������Ȗ͗l�ɂ���
'---------------------------------------------------------------------------------------
'
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).ApplyDataLabels
    ActiveChart.FullSeriesCollection(3).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.25
        .Transparency = 0
        .Solid
        .Patterned msoPatternWideUpwardDiagonal
    End With
    ActiveChart.FullSeriesCollection(5).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .Patterned msoPatternWideUpwardDiagonal
    End With
    ActiveChart.FullSeriesCollection(7).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .Patterned msoPatternWideUpwardDiagonal
    End With

End Sub

Sub MG_��������5()
'---------------------------------------------------------------------------------------
'�O���[�v���̔����������O���t�ɂ��鏈��
'�@�}������ɔz�u����
'---------------------------------------------------------------------------------------
'
    ActiveChart.Legend.Select
    ActiveChart.SetElement (msoElementLegendBottom)

End Sub

