Attribute VB_Name = "SensitivityAnalysis"
'==============================================================================
' SENSITIVITY ANALYSIS ADD-IN
' Tornado Chart & Spider Chart Generator
' Free to use and distribute
'==============================================================================
Option Explicit

'--- Constants ----------------------------------------------------------------
Private Const ADDIN_NAME    As String = "Sensitivity Analysis"
Private Const TORNADO_SHEET As String = "Tornado Chart"
Private Const SPIDER_SHEET  As String = "Spider Chart"
Private Const MAX_INPUTS    As Integer = 20

'--- Public Types (must be Public to share across modules) --------------------
Public Type InputDef
    Label       As String
    cell        As Range
    BaseValue   As Double
    LowPct      As Double
    HighPct     As Double
End Type

Public Type SensitivityConfig
    OutputCell      As Range
    BaseOutput      As Double
    NumInputs       As Integer
    Inputs(1 To 20) As InputDef
    NumPoints       As Integer
End Type

Public Type SensitivityResult
    Label                 As String
    LowOutput             As Double
    HighOutput            As Double
    Swing                 As Double
    LowPct                As Double
    HighPct               As Double
    SpiderPcts(0 To 8)    As Double
    SpiderOutputs(0 To 8) As Double
End Type

'==============================================================================
' ENTRY POINT
'==============================================================================
Public Sub RunSensitivityAnalysis(control As IRibbonControl)
    Dim cfg As SensitivityConfig
    If Not GetUserInputs(cfg) Then Exit Sub
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrHandler
    Dim results(1 To 20) As SensitivityResult
    Call RunSensitivity(cfg, results)
    Call BuildTornadoChart(results, cfg)
    Call BuildSpiderChart(results, cfg)
    Call RestoreInputs(cfg)
    Application.Calculate
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Sensitivity analysis complete!" & vbCrLf & _
           "- Tornado Chart tab created" & vbCrLf & _
           "- Spider Chart tab created", vbInformation, ADDIN_NAME
    Exit Sub
ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, ADDIN_NAME
End Sub

Public Sub RunFromRibbon(control As IRibbonControl)
    RunSensitivityAnalysis
End Sub

'==============================================================================
' USER INPUT
'==============================================================================
Private Function GetUserInputs(cfg As SensitivityConfig) As Boolean
    GetUserInputs = False
    Dim outRng As Range
    On Error Resume Next
    Set outRng = Application.InputBox( _
        Prompt:="Select the OUTPUT cell (the result your model calculates):", _
        Title:=ADDIN_NAME & " - Step 1 of 3", _
        Type:=8)
    On Error GoTo 0
    If outRng Is Nothing Then Exit Function
    If outRng.Count <> 1 Then
        MsgBox "Please select exactly one output cell.", vbExclamation, ADDIN_NAME
        Exit Function
    End If
    Set cfg.OutputCell = outRng
    cfg.BaseOutput = outRng.Value
    Dim inRng As Range
    On Error Resume Next
    Set inRng = Application.InputBox( _
        Prompt:="Select all INPUT cells to vary (hold Ctrl to select multiple cells):" & vbCrLf & _
                "Up to " & MAX_INPUTS & " inputs supported.", _
        Title:=ADDIN_NAME & " - Step 2 of 3", _
        Type:=8)
    On Error GoTo 0
    If inRng Is Nothing Then Exit Function
    If inRng.Count > MAX_INPUTS Then
        MsgBox "Please select no more than " & MAX_INPUTS & " input cells.", vbExclamation, ADDIN_NAME
        Exit Function
    End If
    Dim pctStr As String
    pctStr = InputBox( _
        "Enter the percentage variation to apply to each input." & vbCrLf & _
        "Example: enter 10 for +/-10%", _
        ADDIN_NAME & " - Step 3 of 3", "10")
    If pctStr = "" Then Exit Function
    Dim defaultPct As Double
    defaultPct = Abs(CDbl(pctStr)) / 100
    cfg.NumInputs = inRng.Count
    cfg.NumPoints = 9
    Dim i As Integer
    Dim cell As Range
    i = 1
    For Each cell In inRng
        Set cfg.Inputs(i).cell = cell
        cfg.Inputs(i).BaseValue = cell.Value
        cfg.Inputs(i).Label = GetCellLabel(cell)
        cfg.Inputs(i).LowPct = -defaultPct
        cfg.Inputs(i).HighPct = defaultPct
        i = i + 1
    Next cell
    GetUserInputs = True
End Function

'==============================================================================
' CORE SENSITIVITY CALCULATION
'==============================================================================
Private Sub RunSensitivity(cfg As SensitivityConfig, results() As SensitivityResult)
    Dim i As Integer
    Dim j As Integer
    For i = 1 To cfg.NumInputs
        results(i).Label = cfg.Inputs(i).Label
        results(i).LowPct = cfg.Inputs(i).LowPct
        results(i).HighPct = cfg.Inputs(i).HighPct
        ' Low
        cfg.Inputs(i).cell.Value = cfg.Inputs(i).BaseValue * (1 + cfg.Inputs(i).LowPct)
        Application.Calculate
        results(i).LowOutput = cfg.OutputCell.Value
        ' High
        cfg.Inputs(i).cell.Value = cfg.Inputs(i).BaseValue * (1 + cfg.Inputs(i).HighPct)
        Application.Calculate
        results(i).HighOutput = cfg.OutputCell.Value
        results(i).Swing = Abs(results(i).HighOutput - results(i).LowOutput)
        ' Restore before spider
        cfg.Inputs(i).cell.Value = cfg.Inputs(i).BaseValue
        Application.Calculate
        ' Spider points: -40% to +40% in 10% steps
        Dim stepSize As Double
        stepSize = 0.8 / (cfg.NumPoints - 1)
        For j = 0 To cfg.NumPoints - 1
            results(i).SpiderPcts(j) = -0.4 + j * stepSize
            If Abs(results(i).SpiderPcts(j)) < 0.0001 Then results(i).SpiderPcts(j) = 0
            cfg.Inputs(i).cell.Value = cfg.Inputs(i).BaseValue * (1 + results(i).SpiderPcts(j))
            Application.Calculate
            results(i).SpiderOutputs(j) = cfg.OutputCell.Value
        Next j
        ' Restore
        cfg.Inputs(i).cell.Value = cfg.Inputs(i).BaseValue
        Application.Calculate
    Next i
    Call SortResultsBySwing(results, cfg.NumInputs)
End Sub

Private Sub SortResultsBySwing(results() As SensitivityResult, n As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim tmp As SensitivityResult
    For i = 1 To n - 1
        For j = 1 To n - i
            If results(j).Swing < results(j + 1).Swing Then
                tmp = results(j)
                results(j) = results(j + 1)
                results(j + 1) = tmp
            End If
        Next j
    Next i
End Sub

'==============================================================================
' TORNADO CHART
'==============================================================================
Private Sub BuildTornadoChart(results() As SensitivityResult, cfg As SensitivityConfig)
    Call DeleteSheet(TORNADO_SHEET)
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    ws.Name = TORNADO_SHEET
    Dim n As Integer
    n = cfg.NumInputs
    ' Header
    ws.Cells(1, 1).Value = "Input"
    ws.Cells(1, 2).Value = "Low Output"
    ws.Cells(1, 3).Value = "High Output"
    ws.Cells(1, 4).Value = "Base Output"
    ws.Cells(1, 5).Value = "Swing"
    ' Data - written bottom to top so biggest swing appears at chart top
    Dim i As Integer
    Dim r As Integer
    For i = 1 To n
        r = n - i + 2
        ws.Cells(r, 1).Value = results(i).Label
        ws.Cells(r, 2).Value = results(i).LowOutput
        ws.Cells(r, 3).Value = results(i).HighOutput
        ws.Cells(r, 4).Value = cfg.BaseOutput
        ws.Cells(r, 5).Value = results(i).Swing
    Next i
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, 5))
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
    End With
    ' Chart source columns G-J
    ws.Cells(1, 7).Value = "Label"
    ws.Cells(1, 8).Value = "Spacer"
    ws.Cells(1, 9).Value = "Low Impact"
    ws.Cells(1, 10).Value = "High Impact"
    Dim baseVal As Double
    baseVal = cfg.BaseOutput
    Dim lo As Double
    Dim hi As Double
    For i = 1 To n
        r = n - i + 2
        lo = results(i).LowOutput
        hi = results(i).HighOutput
        ws.Cells(r, 7).Value = results(i).Label
        ws.Cells(r, 8).Value = WorksheetFunction.Min(lo, hi)
        ws.Cells(r, 9).Value = Abs(baseVal - lo)
        ws.Cells(r, 10).Value = Abs(hi - baseVal)
    Next i
    With ws.Range(ws.Cells(1, 7), ws.Cells(1, 10))
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Columns("A:J").AutoFit
    ' Build chart
    Dim chObj As ChartObject
    Set chObj = ws.ChartObjects.Add( _
        Left:=ws.Columns(7).Left, _
        Top:=ws.Rows(n + 4).Top, _
        Width:=560, _
        Height:=WorksheetFunction.Max(250, n * 28 + 80))
    Dim ch As Chart
    Set ch = chObj.Chart
    ch.ChartType = xlBarStacked
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop
    ' Spacer (invisible)
    Dim sSpacer As Series
    Set sSpacer = ch.SeriesCollection.NewSeries
    sSpacer.Name = "Spacer"
    sSpacer.Values = ws.Range(ws.Cells(2, 8), ws.Cells(n + 1, 8))
    sSpacer.XValues = ws.Range(ws.Cells(2, 7), ws.Cells(n + 1, 7))
    sSpacer.Format.Fill.Visible = msoFalse
    sSpacer.Format.Line.Visible = msoFalse
    ' Low impact (red) - no labels to avoid duplicate values
    Dim sLow As Series
    Set sLow = ch.SeriesCollection.NewSeries
    sLow.Name = "Low Impact"
    sLow.Values = ws.Range(ws.Cells(2, 9), ws.Cells(n + 1, 9))
    sLow.XValues = ws.Range(ws.Cells(2, 7), ws.Cells(n + 1, 7))
    sLow.Format.Fill.ForeColor.RGB = RGB(192, 0, 0)
    sLow.Format.Line.Visible = msoFalse
    ' High impact (blue)
    Dim sHigh As Series
    Set sHigh = ch.SeriesCollection.NewSeries
    sHigh.Name = "High Impact"
    sHigh.Values = ws.Range(ws.Cells(2, 10), ws.Cells(n + 1, 10))
    sHigh.XValues = ws.Range(ws.Cells(2, 7), ws.Cells(n + 1, 7))
    sHigh.Format.Fill.ForeColor.RGB = RGB(31, 73, 125)
    sHigh.Format.Line.Visible = msoFalse
    ' Base line dashed
    Dim sBase As Series
    Set sBase = ch.SeriesCollection.NewSeries
    sBase.ChartType = xlLineMarkers
    sBase.Name = "Base Value"
    Dim baseSlice() As Double
    ReDim baseSlice(1 To n)
    For i = 1 To n
        baseSlice(i) = baseVal
    Next i
    sBase.Values = baseSlice
    sBase.Format.Line.ForeColor.RGB = RGB(0, 0, 0)
    sBase.Format.Line.DashStyle = msoLineDash
    sBase.Format.Line.Weight = 1.5
    sBase.MarkerStyle = xlMarkerStyleNone
    ch.HasTitle = True
    ch.ChartTitle.Text = "Tornado Chart - Sensitivity Analysis"
    ch.ChartTitle.Font.Size = 14
    ch.ChartTitle.Font.Bold = True
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    ' Remove spacer from legend (it is series 1)
    ch.Legend.LegendEntries(1).Delete
    ch.Axes(xlValue).HasTitle = True
    ch.Axes(xlValue).AxisTitle.Text = "Output Value"
    ch.PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ws.Cells(n + 3, 7).Value = "How to read: Wider bars = more sensitive. Bars show output range when each input varies low/high."
    ws.Cells(n + 3, 7).Font.Italic = True
    ws.Cells(n + 3, 7).Font.Color = RGB(89, 89, 89)
End Sub

'==============================================================================
' SPIDER CHART
'==============================================================================
Private Sub BuildSpiderChart(results() As SensitivityResult, cfg As SensitivityConfig)
    Call DeleteSheet(SPIDER_SHEET)
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    ws.Name = SPIDER_SHEET
    Dim n As Integer
    n = cfg.NumInputs
    Dim nPts As Integer
    nPts = cfg.NumPoints
    ' Header row
    ws.Cells(1, 1).Value = "Input \ % Change"
    Dim j As Integer
    For j = 0 To nPts - 1
        ws.Cells(1, j + 2).Value = results(1).SpiderPcts(j)
        ws.Cells(1, j + 2).NumberFormat = "0%"
    Next j
    ' Data rows
    Dim i As Integer
    For i = 1 To n
        ws.Cells(i + 1, 1).Value = results(i).Label
        For j = 0 To nPts - 1
            ws.Cells(i + 1, j + 2).Value = results(i).SpiderOutputs(j)
        Next j
    Next i
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, nPts + 1))
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    ws.Columns("A:" & ColLetter(nPts + 1)).AutoFit
    ' Build chart
    Dim chObj As ChartObject
    Set chObj = ws.ChartObjects.Add( _
        Left:=10, _
        Top:=ws.Rows(n + 4).Top, _
        Width:=580, _
        Height:=380)
    Dim ch As Chart
    Set ch = chObj.Chart
    ch.ChartType = xlLine
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop
    Dim colors(0 To 9) As Long
    colors(0) = RGB(31, 73, 125)
    colors(1) = RGB(192, 0, 0)
    colors(2) = RGB(0, 128, 0)
    colors(3) = RGB(255, 102, 0)
    colors(4) = RGB(112, 48, 160)
    colors(5) = RGB(0, 176, 240)
    colors(6) = RGB(255, 192, 0)
    colors(7) = RGB(146, 208, 80)
    colors(8) = RGB(255, 0, 255)
    colors(9) = RGB(0, 112, 192)
    Dim s As Series
    For i = 1 To n
        Set s = ch.SeriesCollection.NewSeries
        s.Name = results(i).Label
        s.Values = ws.Range(ws.Cells(i + 1, 2), ws.Cells(i + 1, nPts + 1))
        s.XValues = ws.Range(ws.Cells(1, 2), ws.Cells(1, nPts + 1))
        s.Format.Line.ForeColor.RGB = colors((i - 1) Mod 10)
        s.Format.Line.Weight = 2
        s.MarkerStyle = xlMarkerStyleCircle
        s.MarkerSize = 6
        s.MarkerForegroundColor = colors((i - 1) Mod 10)
        s.MarkerBackgroundColor = colors((i - 1) Mod 10)
    Next i
    ch.HasTitle = True
    ch.ChartTitle.Text = "Spider Chart - Sensitivity Analysis"
    ch.ChartTitle.Font.Size = 14
    ch.ChartTitle.Font.Bold = True
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    ch.Axes(xlCategory).HasTitle = True
    ch.Axes(xlCategory).AxisTitle.Text = "% Change in Input"
    ch.Axes(xlValue).HasTitle = True
    ch.Axes(xlValue).AxisTitle.Text = "Output Value"
    ch.Axes(xlValue).HasMajorGridlines = True
    ch.Axes(xlCategory).HasMajorGridlines = True
    ch.PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ws.Cells(n + 3, 1).Value = "How to read: Each line shows how the output changes as one input varies from -40% to +40%. Steeper = more sensitive."
    ws.Cells(n + 3, 1).Font.Italic = True
    ws.Cells(n + 3, 1).Font.Color = RGB(89, 89, 89)
End Sub

'==============================================================================
' HELPERS
'==============================================================================
Private Sub RestoreInputs(cfg As SensitivityConfig)
    Dim i As Integer
    For i = 1 To cfg.NumInputs
        cfg.Inputs(i).cell.Value = cfg.Inputs(i).BaseValue
    Next i
    Application.Calculate
End Sub

Private Sub DeleteSheet(sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
End Sub

Private Function GetCellLabel(cell As Range) As String
    Dim lbl As String
    lbl = ""
    If cell.Column > 1 Then
        If Len(Trim(CStr(cell.Offset(0, -1).Value))) > 0 Then
            If Not IsNumeric(cell.Offset(0, -1).Value) Then
                lbl = Trim(CStr(cell.Offset(0, -1).Value))
            End If
        End If
    End If
    If lbl = "" And cell.Row > 1 Then
        If Len(Trim(CStr(cell.Offset(-1, 0).Value))) > 0 Then
            If Not IsNumeric(cell.Offset(-1, 0).Value) Then
                lbl = Trim(CStr(cell.Offset(-1, 0).Value))
            End If
        End If
    End If
    If lbl = "" Then lbl = cell.Address(False, False)
    GetCellLabel = lbl
End Function

Private Function ColLetter(colNum As Integer) As String
    ColLetter = Split(Cells(1, colNum).Address, "$")(1)
End Function
