Sub Main()
    UserForm1.Show
End Sub
--------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub getInfo(cn As ADODB.Connection, strategy As Integer, strategyName As String)
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    Dim firstCell As Range
    'Display model name
    Range("A2").Value = strategyName
    Set firstCell = Sheet1.Range("C3")
    
    SQL = "SELECT REVENUE, VARIABLEX, FIXEDX, PROFIT, DEVR, DEVF FROM Categories C INNER JOIN Strategies S ON C.CategoryID = S.STRATEGY WHERE S.STRATEGY=" & strategy
    With rs
        .Open SQL, cn
        Do While Not .EOF
            firstCell.Offset(0, 0).Value = .Fields("REVENUE")
            firstCell.Offset(0, 1).Value = .Fields("VARIABLEX")
            firstCell.Offset(0, 2).Value = .Fields("FIXEDX")
            firstCell.Offset(0, 3).Value = .Fields("PROFIT")
            firstCell.Offset(1, 0).Value = .Fields("DEVR")
            firstCell.Offset(1, 1).Value = .Fields("DEVF")
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Sheet1.UsedRange.EntireColumn.AutoFit

End Sub
----------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub generate_ScatterPlot(oChartObj As ChartObject, oChart As Chart)
    Dim oChartRange As Range
    Dim ser As Series
    
    
    
    Set oChartRange = Range("B9", Range("C9").End(xlDown))
    oChart.ChartType = xlXYScatterLinesNoMarkers
    oChart.SetSourceData Source:=oChartRange
    'Formatting the titles
    oChart.Axes(xlCategory).HasTitle = True
    oChart.Axes(xlCategory).AxisTitle.Caption = "Profit/Loss"
    oChart.Axes(xlCategory).HasTitle = True
    oChart.Axes(xlCategory).AxisTitle.Caption = Range("B8")
    
    
End Sub
----------------------------------------------------------------------------------------------------------------------------------------------------------
Sub generateWordDocument(oChartObj As ChartObject)
    Dim wdDoc As Word.Document
    Dim wdSel As Word.Selection
    
    'create new instance of word
    Dim wdApp As New Word.Application
    
    wdApp.Visible = True
    
    Set wdDoc = wdApp.Documents.Add
    Set wdSel = wdDoc.ActiveWindow.Selection
    
    wdSel.TypeText "Monte Carlo Simulation Results Explained: "
    wdSel.TypeText "After receiving values from variables such as: revenue, variable and fixed expenses, and standard deviation"
    
    wdSel.TypeText " the program completes its first simulation using normal distribution. On the excel file, after the program takes its time to run"
    wdSel.TypeText " the net profit/loss of each simulation is displayed beside its assigned iteration number. Values in red represent simulations that resulted in a net loss,"
    wdSel.TypeText " values in green represent iterations that resulted in a profit. The program also calculates the likelihood of a net loss as a percentage, taking into consideration the number of simulations performed by the program."
    wdSel.TypeText " Through the user form, the user is asked what their profit goal is, after running the program, the likelihood of having a profit value greater than their desired goal is calculated and displayed as a percentage on the excel file"
    wdSel.TypeText " Below you will find the chart that is displayed to the user on the excel file, it is a lined scatter plot that plots the results of each iteration. The larger the number of iterations defined by the user, the more accurate the results of the simulation"
    wdSel.TypeText " It is easy to identify a consistent pattern/region of values that the simulation results in. "
    
    oChartObj.Copy
    
    
    With wdApp.Selection
        .EndKey Unit:=wdStory
        .TypeParagraph
        .PasteSpecial Link:=False, DataType:=wdPasteBitmap, _
        Placement:=wdInLine, DisplayAsIcon:=False
    End With
    
End Sub
