Attribute VB_Name = "Module2"
Option Explicit

Sub MG_��������()

    Call MG_��������1
    Call MG_��������2
    Call MG_��������3
        
End Sub


Sub MG_��������1()
Attribute MG_��������1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' g_G Macro
'
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlColumnClustered
    
    '�f�[�^�͈͐ݒ�
    ActiveChart.SetSourceData Source:=Range("�W�v���!$D$47:$F$59,�W�v���!$R$47:$R$59,�W�v���!$T$47:$T$59,�W�v���!$EE$17:$EE$28")
    ActiveChart.PlotArea.Select
    
    'X���������ݒ�
    ActiveChart.SeriesCollection(1).XValues = "=�W�v���!$B$47:$B$58"
    
    '�}��1 & �f�[�^�G���A1�@�ݒ�
    ActiveChart.SeriesCollection(1).Name = "=�W�v���!$D$33"
    ActiveChart.SeriesCollection(1).Values = "=�W�v���!$D$47:$D$58"
    
    '�}��2 & �f�[�^�G���A2�@�ݒ�
    ActiveChart.SeriesCollection(2).Name = "=�W�v���!$R$33"
    ActiveChart.SeriesCollection(2).Values = "=�W�v���!$R$47:$R$58"
    
    '�}��3 & �f�[�^�G���A3�@�ݒ�
    ActiveChart.SeriesCollection(3).Name = "=�W�v���!$E$33"
    ActiveChart.SeriesCollection(3).Values = "=�W�v���!$E$47:$E$58"
    
    '�}��2 & �f�[�^�G���A2�@�ݒ�
    ActiveChart.SeriesCollection(4).Name = "=�W�v���!$S$33"
    ActiveChart.SeriesCollection(4).Values = "=�W�v���!$EE$17:$EE$28"
    
    '�}��4 & �f�[�^�G���A4�@�ݒ�
    ActiveChart.SeriesCollection(5).Name = "=�W�v���!$F$33"
    ActiveChart.SeriesCollection(5).Values = "=�W�v���!$F$47:$F$58"

    '�}��2 & �f�[�^�G���A2�@�ݒ�
    ActiveChart.SeriesCollection(6).Name = "=�W�v���!$R$33"
    ActiveChart.SeriesCollection(6).Values = "=�W�v���!$T$47:$T$58"

'    ActiveChart.Name = "���Ԕ�������"

End Sub
Sub MG_��������2()
Attribute MG_��������2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(7).Name = "=�W�v���!$C$34"
    ActiveChart.SeriesCollection(7).Values = "=�W�v���!$C$47:$C$58"
    ActiveChart.SeriesCollection(7).Select
    ActiveChart.SeriesCollection(7).AxisGroup = 2
    ActiveChart.ChartGroups(2).Overlap = 0
    ActiveChart.ChartGroups(2).GapWidth = 0
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
'
' Macro3 Macro
'
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MaximumScale = 200

End Sub
