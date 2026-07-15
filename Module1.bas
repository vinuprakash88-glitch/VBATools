Attribute VB_Name = "Module1"
Option Explicit

Sub PREPARE_DATA()
'-------------------------------SWITCHING OFF DEFAULTS
With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
    .CutCopyMode = False
    .Calculation = xlCalculationManual
End With
'-------------------------------VARIABLE NOTATIONS
'CW         - Current Workbook
'RDWS       - 'REQUIRED DATA' Worksheet
'CWS        - 'CHART' Worksheet
'OWFP       - Opened Workbook Folder Path
'OWFN       - Opened Workbook File Name
'OW         - Opened Workbook
'OWFWS      - Opened Workbook First Worksheet
'OWFWS_SR   - Opened Workbook First Worksheet Source Range
'OWFWS_CR   - Opened Workbook First Worksheet Criteria Range
'RDWS_TR    - Required Data Worksheet Target Range

'-------------------------------VARIABLE DECLARATIONS
Dim CW As Workbook
Dim RDWS As Worksheet
Dim CWS As Worksheet
Dim OWFP As String
Dim OWFN As String
Dim OW As Workbook
Dim OWFWS As Worksheet
Dim OWFWS_SR As Range
Dim OWFWS_CR As Range
Dim RDWS_TR As Range
'-------------------------------VARIABLE ASSIGMENTS
Set CW = ThisWorkbook
Set RDWS = CW.Worksheets("REQUIRED DATA")
Set CWS = CW.Worksheets("CHART")
OWFP = CW.Path & Application.PathSeparator & "NSE EXPORT" & Application.PathSeparator
OWFN = Dir(OWFP)
Set OW = Workbooks.Open(OWFP & OWFN, ReadOnly:=True)
Set OWFWS = OW.Worksheets(OW.Sheets(1).Name)
Set OWFWS_SR = OWFWS.Range("A1").CurrentRegion
Set OWFWS_CR = RDWS.Range("E1:F1")
Set RDWS_TR = RDWS.Range("E3:F3")

'-------------------------------CODE
CWS.Range("A1").CurrentRegion.Columns.EntireColumn.Delete
With RDWS
    .Range("E1").CurrentRegion.Columns.EntireColumn.Clear
    .Range("B8").ClearContents
    .Range("B10").ClearContents
    .Range("E1").Value = "DATE"
    .Range("F1").Value = "CLOSE"
    .Range("E3:F3").Value = .Range("E1:F1").Value
End With
OWFWS_SR.AdvancedFilter xlFilterCopy, _
    CriteriaRange:=OWFWS_CR, _
    CopyToRange:=RDWS_TR
OW.Close SaveChanges:=False
With RDWS
    .Range("E1:F2").Delete Shift:=xlUp
    .Range("E102:E" & .Cells(Rows.Count, 5).End(xlUp).Row).Rows.EntireRow.Delete Shift:=xlUp
    .Range("E1").CurrentRegion.Sort Key1:=.Range("E:F"), Order1:=xlAscending, Header:=xlYes
    .Range("G1").Value = "FROM"
    .Range("G3:G" & .Cells(Rows.Count, 5).End(xlUp).Row).Formula = "=$F2"
    .Range("H1").Value = "TO"
    .Range("H3:H" & .Cells(Rows.Count, 5).End(xlUp).Row).Formula = "=$F3"
    .Range("I1").Value = "FROM(RENKO VALUE)"
    .Range("I3:I" & .Cells(Rows.Count, 5).End(xlUp).Row).Formula = "=QUOTIENT($G3,$B$6)*$B$6"
    .Range("J1").Value = "TO(RENKO VALUE)"
    .Range("J3:J" & .Cells(Rows.Count, 5).End(xlUp).Row).Formula = "=QUOTIENT($H3,$B$6)*$B$6"
    .Range("K1").Value = "NO. OF BRICKS"
    .Range("K3:K" & .Cells(Rows.Count, 5).End(xlUp).Row).Formula = "=ABS((J3-I3)/$B$6)"
    .Range("L1").Value = "TREND DIRECTION"
    .Range("L3:L" & .Cells(Rows.Count, 5).End(xlUp).Row).Formula = "=IF(J3>I3,""UPTICK"",IF(J3<I3,""DOWNTICK"",""-""))"
    .Range("G2:L2").Value = "-"
    .Range("F2:J101").NumberFormat = "$ #,##0.00"
    .Range("E1:M1").Copy
    .Range("E105").PasteSpecial xlPasteAll
    .Range("L106").Value = "UPTICK"
    .Range("L107").Value = "DOWNTICK"
    .Range("E1").CurrentRegion.AdvancedFilter xlFilterCopy, _
        CriteriaRange:=.Range("E105:L107"), _
        CopyToRange:=.Range("E110")
    .Range("E1:L109").Delete Shift:=xlUp
    .Range("B8").Value = Application.WorksheetFunction.Max(.Range("I:I")) + (2 * .Range("B6").Value)
    .Range("B10").Value = Application.WorksheetFunction.Min(.Range("I:I")) - (2 * .Range("B6").Value)
End With

Call CREATE_CHART_TICKS

RDWS.Range("G2:P" & RDWS.Cells(Rows.Count, 5).End(xlUp).Row).Value = RDWS.Range("G2:P" & RDWS.Cells(Rows.Count, 5).End(xlUp).Row).Value

With RDWS.Range("E1:P1")
    .Font.Bold = True
    .Interior.Color = RGB(149, 198, 196)
End With
With RDWS.Range("E1").CurrentRegion
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Borders.LineStyle = xlContinuous
    .Columns.EntireColumn.AutoFit
End With

'-------------------------------TURNING THE DEFAULTS BACK ON
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .CutCopyMode = True
    .Calculation = xlCalculationAutomatic
End With

End Sub

Sub CREATE_CHART_TICKS()
'-------------------------------SWITCHING OFF DEFAULTS
With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
    .CutCopyMode = False
    .Calculation = xlCalculationManual
End With

'-------------------------------VARIABLE DECLARATIONS
Dim CW As Workbook
Dim RDWS As Worksheet
Dim CWS As Worksheet
Dim i As Long
'-------------------------------VARIABLE NOTATIONS
'CW     - Current Workbook
'RDWS   - Required Data Worksheet
'i      - Looping Variable
'-------------------------------VARIABLE ASSIGMENTS
Set CW = ThisWorkbook
Set RDWS = CW.Worksheets("REQUIRED DATA")
Set CWS = CW.Worksheets("CHART")
'-------------------------------CODE
With CWS
    .Range("A1:A" & .Cells(Rows.Count, 1).End(xlUp).Row).ClearContents
    .Range("A1").Value = RDWS.Range("B8").Value
    i = 2
    Do Until .Range("A" & i - 1).Value <= RDWS.Range("B10").Value
        .Range("A" & i).Value = .Range("A" & i - 1).Value - RDWS.Range("B6").Value
    i = i + 1
    Loop
End With
With RDWS
    .Range("M1").Value = "STARTING ROW"
    .Range("N1").Value = "STARTING COLUMN"
    .Range("O1").Value = "TARGET ROW"
    .Range("P1").Value = "TARGET COLUMN"
    .Range("M2").Formula = "=MATCH($I2,CHART!$A:$A,0)"
    .Range("N2").Value = 2
    .Range("O2").Formula = "=IF($L2=""UPTICK"",$M2-$K2,IF($L2=""DOWNTICK"",$M2+$K2,""-""))"
    .Range("P2").Formula = "=$N2+$K2"
    .Range("M3:M" & .Cells(Rows.Count, 5).End(xlUp).Row).Formula = "=$O2"
    .Range("N3:N" & .Cells(Rows.Count, 5).End(xlUp).Row).Formula = "=$P2"
    .Range("O3:O" & .Cells(Rows.Count, 5).End(xlUp).Row).Formula = "=IF($L3=""UPTICK"",$M3-$K3,IF($L3=""DOWNTICK"",$M3+$K3,""-""))"
    .Range("P3:P" & .Cells(Rows.Count, 5).End(xlUp).Row).Formula = "=$N3+$K3"
End With

'-------------------------------TURNING THE DEFAULTS BACK ON
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .CutCopyMode = True
    .Calculation = xlCalculationAutomatic
End With

End Sub

Sub RENKO_CHART()
'-------------------------------SWITCHING OFF DEFAULTS
With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
    .CutCopyMode = False
    .Calculation = xlCalculationManual
End With

'-------------------------------VARIABLE NOTATIONS
'CW         - Current Workbook
'RDWS       - Required Data Worksheet
'CWS        - 'CHART' Worksheet
'i          - Looping Variable
'j          - Looping Variable

'-------------------------------VARIABLE DECLARATIONS
Dim CW As Workbook
Dim RDWS As Worksheet
Dim CWS As Worksheet
Dim i As Long
Dim j As Long

'-------------------------------VARIABLE ASSIGMENTS
Set CW = ThisWorkbook
Set RDWS = CW.Worksheets("REQUIRED DATA")
Set CWS = CW.Worksheets("CHART")

'-------------------------------CODE
CWS.Range("A1").CurrentRegion.EntireColumn.Delete Shift:=xlToLeft
Call PREPARE_DATA
i = 2
Do Until RDWS.Cells(i, 12).Value = ""
    If RDWS.Cells(i, 12).Value = "UPTICK" Then
        For j = 0 To RDWS.Range("K" & i).Value
            With CWS.Cells(RDWS.Range("M" & i).Value - j, RDWS.Range("N" & i).Value + j)
                .Value = "+"
                .Borders.LineStyle = xlContinuous
                .Interior.Color = RGB(50, 255, 50)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        Next
        ElseIf RDWS.Cells(i, 12).Value = "DOWNTICK" Then
        For j = 0 To RDWS.Range("K" & i).Value
            With CWS.Cells(RDWS.Range("M" & i).Value + j, RDWS.Range("N" & i).Value + j)
                .Value = "-"
                .Borders.LineStyle = xlContinuous
                .Interior.Color = RGB(255, 50, 50)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        Next
    
    End If
i = i + 1
Loop
CWS.Range("A1").CurrentRegion.EntireColumn.AutoFit

'-------------------------------TURNING THE DEFAULTS BACK ON
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .CutCopyMode = True
    .Calculation = xlCalculationAutomatic
End With

End Sub

Sub CREATE_RENKO_CHART()

With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
    .CutCopyMode = False
    .Calculation = xlCalculationManual
End With
    Dim CW As Workbook
    Set CW = ThisWorkbook
    Dim RDWS As Worksheet
    Set RDWS = CW.Worksheets("REQUIRED DATA")
    Dim CWS As Worksheet
    Set CWS = CW.Worksheets("CHART")
    RDWS.Unprotect Password:=""
    CWS.Unprotect Password:=""
    Call PREPARE_DATA
    Call CREATE_CHART_TICKS
    Call RENKO_CHART
    RDWS.Protect Password:=""
    CWS.Protect Password:=""
    Application.Goto Reference:=CWS.Range("A1"), Scroll:=True
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .CutCopyMode = True
    .Calculation = xlCalculationAutomatic
End With

End Sub

