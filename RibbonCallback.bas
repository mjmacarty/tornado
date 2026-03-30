Attribute VB_Name = "RibbonCallback"
'==============================================================================
' RIBBON CALLBACK
' Adds "Sensitivity Analysis" button to Excel ribbon
'==============================================================================
Option Explicit

' Called by the ribbon button
Public Sub RunFromRibbon(control As IRibbonControl)
    RunCore
End Sub

' Can also be called directly from VBA editor for testing
Public Sub RunCore()
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
           "- Spider Chart tab created", vbInformation, "Sensitivity Analysis"
    Exit Sub
ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Sensitivity Analysis"
End Sub

' Ribbon XML - paste this into the CustomUI editor
' (see INSTALL.md for full instructions)
'
' <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
'   <ribbon>
'     <tabs>
'       <tab id="tabSensitivity" label="Sensitivity">
'         <group id="grpSensitivity" label="Analysis">
'           <button id="btnRunSensitivity"
'                   label="Run Analysis"
'                   imageMso="ChartInsert"
'                   size="large"
'                   onAction="RunFromRibbon"
'                   screentip="Tornado &amp; Spider Chart Sensitivity Analysis"
'                   supertip="Select an output cell and input cells to generate Tornado and Spider charts showing model sensitivity." />
'         </group>
'       </tab>
'     </tabs>
'   </ribbon>
' </customUI>
