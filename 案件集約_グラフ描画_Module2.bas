Attribute VB_Name = "Module2"
Option Explicit

Sub MG_発生件数()

    Call MG_発生件数1
    Call MG_発生件数2
    Call MG_発生件数3
        
End Sub


Sub MG_発生件数1()
Attribute MG_発生件数1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' g_G Macro
'
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlColumnClustered
    
    'データ範囲設定
    ActiveChart.SetSourceData Source:=Range("集計情報!$D$47:$F$59,集計情報!$R$47:$R$59,集計情報!$T$47:$T$59,集計情報!$EE$17:$EE$28")
    ActiveChart.PlotArea.Select
    
    'X軸メモリ設定
    ActiveChart.SeriesCollection(1).XValues = "=集計情報!$B$47:$B$58"
    
    '凡例1 & データエリア1　設定
    ActiveChart.SeriesCollection(1).Name = "=集計情報!$D$33"
    ActiveChart.SeriesCollection(1).Values = "=集計情報!$D$47:$D$58"
    
    '凡例2 & データエリア2　設定
    ActiveChart.SeriesCollection(2).Name = "=集計情報!$R$33"
    ActiveChart.SeriesCollection(2).Values = "=集計情報!$R$47:$R$58"
    
    '凡例3 & データエリア3　設定
    ActiveChart.SeriesCollection(3).Name = "=集計情報!$E$33"
    ActiveChart.SeriesCollection(3).Values = "=集計情報!$E$47:$E$58"
    
    '凡例2 & データエリア2　設定
    ActiveChart.SeriesCollection(4).Name = "=集計情報!$S$33"
    ActiveChart.SeriesCollection(4).Values = "=集計情報!$EE$17:$EE$28"
    
    '凡例4 & データエリア4　設定
    ActiveChart.SeriesCollection(5).Name = "=集計情報!$F$33"
    ActiveChart.SeriesCollection(5).Values = "=集計情報!$F$47:$F$58"

    '凡例2 & データエリア2　設定
    ActiveChart.SeriesCollection(6).Name = "=集計情報!$R$33"
    ActiveChart.SeriesCollection(6).Values = "=集計情報!$T$47:$T$58"

'    ActiveChart.Name = "期間発生件数"

End Sub
Sub MG_発生件数2()
Attribute MG_発生件数2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(7).Name = "=集計情報!$C$34"
    ActiveChart.SeriesCollection(7).Values = "=集計情報!$C$47:$C$58"
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
Sub MG_発生件数3()
Attribute MG_発生件数3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MaximumScale = 200

End Sub
