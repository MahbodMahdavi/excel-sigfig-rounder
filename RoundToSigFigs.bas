Option Explicit

Sub RoundToSigFigs(Optional control As IRibbonControl)
    ' Rounds every numeric cell in the current selection to 3 significant figures,
    ' overwriting the cell's value in place (same behavior as the original XLSTART macro).

    Dim cell As Range
    Dim sigFigs As Integer
    Dim val As Double
    Dim magnitude As Integer
    Dim decimals As Integer
    Dim processedCount As Long

    sigFigs = 3
    processedCount = 0

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells first.", vbExclamation, "Round to Sig Figs"
        Exit Sub
    End If

    ' Speed up processing on large ranges and avoid screen flicker
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    For Each cell In Selection
        If Not cell.HasFormula Then
            If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
                val = cell.Value
                If val <> 0 Then
                    magnitude = Int(Application.WorksheetFunction.Log10(Abs(val)))
                    decimals = sigFigs - 1 - magnitude
                    cell.Value = Application.WorksheetFunction.Round(val, decimals)
                    processedCount = processedCount + 1
                End If
            End If
        End If
    Next cell

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox processedCount & " cell(s) rounded to " & sigFigs & " significant figures.", vbInformation, "Round to Sig Figs"
    Exit Sub

CleanFail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Round to Sig Figs"
End Sub
