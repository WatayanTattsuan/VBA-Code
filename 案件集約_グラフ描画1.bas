Attribute VB_Name = "グラフ描画処理1"
Option Explicit

Sub MG_発生件数()
'---------------------------------------------------------------------------------------
'グループ毎の発生件数をグラフにする処理
'
'---------------------------------------------------------------------------------------
'
    Call MG_発生件数1
    Call MG_発生件数2
    Call MG_発生件数3
    Call MG_発生件数4
    Call MG_発生件数5
        
End Sub


Sub MG_発生件数1()
Attribute MG_発生件数1.VB_ProcData.VB_Invoke_Func = " \n14"
'---------------------------------------------------------------------------------------
'グループ毎の発生件数をグラフにする処理
'　案件集約からグラフを作成する
'---------------------------------------------------------------------------------------
'
    Dim chart1 As Chart
    Set chart1 = Charts.Add(after:=ActiveSheet)
    ActiveChart.ChartType = xlColumnClustered
    
    'データ範囲設定
    ActiveChart.SetSourceData Source:=Range("集計情報!$C$47:$C$59,集計情報!$D$47:$D$59,集計情報!$R$47:$R$59,集計情報!$E$47:$E$59,集計情報!$EE$17:$EE$28,集計情報!$F$47:$F$59,集計情報!$T$47:$T$59")
    ActiveChart.PlotArea.Select
    ActiveChart.SeriesCollection(1).Name = "=集計情報!$C$34"
    ActiveChart.SeriesCollection(1).Select
    
    ActiveChart.SeriesCollection(2).Name = "=集計情報!$D$33"
    ActiveChart.SeriesCollection(2).Select
    ActiveChart.SeriesCollection(2).AxisGroup = 2
       
    ActiveChart.SeriesCollection(3).Name = "=集計情報!$R$33"
    ActiveChart.SeriesCollection(3).Select
    ActiveChart.SeriesCollection(3).AxisGroup = 2
       
    ActiveChart.SeriesCollection(4).Name = "=集計情報!$E$33"
    ActiveChart.SeriesCollection(4).Select
    ActiveChart.SeriesCollection(4).AxisGroup = 2
    
       
    ActiveChart.SeriesCollection(5).Name = "=集計情報!$S$33"
    ActiveChart.SeriesCollection(5).Select
    ActiveChart.SeriesCollection(5).AxisGroup = 2
       
    ActiveChart.SeriesCollection(6).Name = "=集計情報!$F$33"
    ActiveChart.SeriesCollection(6).Select
    ActiveChart.SeriesCollection(6).AxisGroup = 2
           
    ActiveChart.SeriesCollection(7).Name = "=集計情報!$T$33"
    ActiveChart.SeriesCollection(7).Select
    ActiveChart.SeriesCollection(7).AxisGroup = 2
    
    'X軸メモリ設定
    ActiveChart.SeriesCollection(1).XValues = "=集計情報!$B$47:$B$58"
    

End Sub
Sub MG_発生件数2()
Attribute MG_発生件数2.VB_ProcData.VB_Invoke_Func = " \n14"
'---------------------------------------------------------------------------------------
'グループ毎の発生件数をグラフにする処理
'　全件発生件数の書式を変更して体裁を整える
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
Sub MG_発生件数3()
Attribute MG_発生件数3.VB_ProcData.VB_Invoke_Func = " \n14"
'---------------------------------------------------------------------------------------
'グループ毎の発生件数をグラフにする処理
'　縦軸の上限値を変更する
'---------------------------------------------------------------------------------------
'
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MaximumScale = 200
    ActiveChart.Axes(xlValue, xlSecondary).MaximumScale = 200

End Sub
Sub MG_発生件数4()
'---------------------------------------------------------------------------------------
'グループ毎の発生件数をグラフにする処理
'　完了件数を縞模様にする
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

Sub MG_発生件数5()
'---------------------------------------------------------------------------------------
'グループ毎の発生件数をグラフにする処理
'　凡例を下に配置する
'---------------------------------------------------------------------------------------
'
    ActiveChart.Legend.Select
    ActiveChart.SetElement (msoElementLegendBottom)

End Sub

