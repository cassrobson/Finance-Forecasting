Option Explicit
Private cancel As Boolean

-------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub CommandButton1_Click()
    Dim cn As New ADODB.Connection
    Dim strategy As Integer, strategyName As String
    Dim rev As Double, devr As Double, var As Double, devx As Double, fixedx As Double
    Dim numberSims As Integer, i As Integer
    Dim goal As Double
    Dim iter As Range
    Dim cell As Range
    Dim negatives As Integer
    Dim aboveGoals As Integer
    Dim loss As Double
    Dim gain As Double
    Dim oChartObj As ChartObject
    Dim oChart As Chart
    Dim existingChart As ChartObject
    
    Sheet1.Activate
    
    
    
    For Each existingChart In ActiveSheet.ChartObjects
        existingChart.Delete
    Next
    
    
    
    Set oChartObj = ActiveSheet.ChartObjects.Add(Top:=250, Left:=325, Width:=600, Height:=300)
    Set oChart = oChartObj.Chart
    
    Worksheets("MonteCarlo").Range("A2").Value = ""
    Worksheets("MonteCarlo").Range("toClear").ClearContents
    
    With cn
        .ConnectionString = "Data Source=" & ThisWorkbook.Path & "\MonteCarlo.accdb"
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    
    strategy = ListBox1.List(ListBox1.ListIndex, 0)
    strategyName = ListBox1.List(ListBox1.ListIndex, 1)
    
    
    Call getInfo(cn, strategy, strategyName)
    rev = Range("C3").Value
    devr = Range("C4").Value
    var = Range("D3").Value
    devx = Range("D4").Value
    fixedx = Range("E3").Value
    
    'First sim
    Range("C6") = WorksheetFunction.Norm_Inv(Rnd, rev, devr)
    Range("D6") = WorksheetFunction.Norm_Inv(Rnd, var, devx)
    Range("E6") = fixedx
    Range("F6") = Range("C6").Value - Range("D6").Value - Range("E6").Value
    Range("C9") = Range("F6").Value
    
    Worksheets("MonteCarlo").Range("B9", Range("B9").End(xlDown)).ClearContents
    Worksheets("MonteCarlo").Range("C9", Range("C9").End(xlDown)).ClearContents
    
    
    'Iterations
    With Range("B8")
        .Value = "Iterations"
        .Font.Bold = True
        .Font.Italic = True
    End With
    goal = CDbl(TextBox2.Value)
    With Range("H3")
        .Value = "Likelihood of losing money"
        .Font.Bold = True
        .Font.Italic = True
        .Offset(1, 0).Value = "Likelihood of hitting profit goal of: $" & goal
        .Offset(1, 0).Font.Bold = True
        .Offset(1, 0).Font.Italic = True
    End With
    
    numberSims = CInt(TextBox1.Value)
    
    With Range("B9")
        For i = 0 To numberSims - 1
            .Offset(i, 0).Value = i + 1
            .Offset(i, 1) = (WorksheetFunction.Norm_Inv(Rnd, rev, devr) - WorksheetFunction.Norm_Inv(Rnd, var, devx) - fixedx)
        Next i
    End With
    
    
    Set iter = Worksheets("MonteCarlo").Range("C9", Range("C9").End(xlDown))
    
    For Each cell In iter
        If Sgn(cell.Value) = -1 Then
            negatives = negatives + 1
            cell.Font.ColorIndex = 3
        ElseIf cell.Value > goal Then
            aboveGoals = aboveGoals + 1
            cell.Font.ColorIndex = 4
        End If
    Next
    
    loss = negatives / numberSims
    gain = aboveGoals / numberSims
    
    With Range("I3")
        .NumberFormat = "0.00%"
        .Value = loss
        .Font.Bold = True
        .Offset(1, 0).NumberFormat = "0.00%"
        .Offset(1, 0).Value = gain
        .Offset(1, 0).Font.Bold = True
    End With
    
    Call generate_ScatterPlot(oChartObj, oChart)
    Call generateWordDocument(oChartObj)
        
    
        
    
    
    
    
    
    
    
    Sheet1.UsedRange.EntireColumn.AutoFit
    Me.Hide
    cancel = False
    MsgBox "You have selected the " & ListBox1.List(ListBox1.ListIndex, 1) & " finance forecasting issue from the MonteCarlo Database. We are using normal distribution with variables like: probability, mean, and standard deviation. This forecasting problem consists of Revenue, Variable and Fixed Expenses. Fixed Expenses are cunk costs in plant and equipment, so no distribution curve is assumed. Distribution curves are assumed for Revenue and Variable Expenses", vbInformation
    MsgBox "The likelihood of losing money: " & Range("I3").Value * 100 & "%" & vbCrLf & "The change of reaching profit goal: " & Range("I4").Value * 100 & "%"
    
End Sub

------------------------------------------------------------------------------------------------------------------------------------------------------------

Private Sub UserForm_Initialize()
    Dim count As Integer
    ListBox1.ColumnCount = 2
    ListBox1.ColumnWidths = "50;50"
    
    For count = 0 To 5
        ListBox1.AddItem
    Next count
    
    ListBox1.List(0, 0) = "1"
    ListBox1.List(1, 0) = "2"
    ListBox1.List(2, 0) = "3"
    ListBox1.List(0, 1) = "High Revenue"
    ListBox1.List(1, 1) = "Balanced"
    ListBox1.List(2, 1) = "Low Revenue"
    
    
    UserForm1.ListBox1.ListIndex = 0
End Sub
Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then CommandButton2_Click
End Sub
Private Sub CommandButton2_Click()
    Me.Hide
    cancel = True
End Sub
