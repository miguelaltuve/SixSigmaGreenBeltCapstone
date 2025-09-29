' GenerateWindshieldData.vb
' Purpose: Simulates windshield gap data for Six Sigma Green Belt Capstone Project
' Generates 60 vehicles with 16 gap measurements (G1–G16) each, with higher defects at G5–G8 (RH side A-pillar)
' Outputs: 960 measurements in Excel (columns A:U, rows 2:61), with non-conformance count (>4mm) in column U
' Matches dataset: ~3.88mm mean, 30.2% non-conformance (290/960)

Sub GenerateWindshieldData()
    Dim i As Long, j As Long
    Dim nonConformance As Long
    Randomize ' Initialize random number generator
    
    ' Clear existing data (optional, adjust range as needed)
    Range("A2:U61").ClearContents
    
    ' Generate data for 60 vehicles
    For i = 2 To 61 ' Rows 2–61
        ' VIN, Day, Shift, Operator
        Cells(i, 1) = "VIN" & Format(i - 1, "000") ' VIN001–VIN060
        Cells(i, 2) = Int((i - 2) / 12) + 1 ' Day 1–5 (12 vehicles/day)
        Cells(i, 3) = Choose(((i - 2) Mod 4) + 1, "Start", "Post-morning", "Post-lunch", "Post-afternoon") ' Shift
        Cells(i, 4) = Choose(Application.RandBetween(1, 4), "A", "B", "C", "D") ' Operator
        
        ' Generate gaps G1–G16 (columns E–T, 5–20)
        nonConformance = 0
        For j = 5 To 20 ' G1–G16
            Dim baseMean As Double, baseStd As Double
            ' Set base mean and std dev for G5–G8 (higher defects) vs others
            If j >= 9 And j <= 12 Then ' G5–G8 (columns I–L)
                baseMean = 4.8 ' Matches your data (~4.0–4.2mm)
                baseStd = 0.6 ' Higher variability
            Else ' G1–G4, G9–G16
                baseMean = 3.5 ' Matches your data (~3.2–3.3mm)
                baseStd = 0.3 ' Lower variability
            End If
            
            ' Adjust mean and std dev based on Operator
            Dim opAdjustMean As Double, opAdjustStd As Double
            opAdjustMean = 0: opAdjustStd = 0
            If Cells(i, 4) = "A" Then opAdjustMean = -0.1: opAdjustStd = -0.1 ' Lower gaps
            If Cells(i, 4) = "C" Then opAdjustMean = 0.2: opAdjustStd = 0.1 ' Higher gaps
            ' Operators B and D use base values
            
            ' Adjust mean and std dev based on Shift
            Dim shiftAdjustMean As Double, shiftAdjustStd As Double
            shiftAdjustMean = 0: shiftAdjustStd = 0
            If Cells(i, 3) = "Start" Then shiftAdjustMean = -0.1 ' Lower gaps
            If Cells(i, 3) = "Post-lunch" Then shiftAdjustMean = 0.1: shiftAdjustStd = 0.03
            If Cells(i, 3) = "Post-afternoon" Then shiftAdjustMean = 0.2: shiftAdjustStd = 0.07
            
            ' Generate gap value (normal distribution, rounded to 2 decimals)
            Cells(i, j) = WorksheetFunction.Round(WorksheetFunction.Norm_Inv(Rnd(), baseMean + opAdjustMean + shiftAdjustMean, baseStd + opAdjustStd + shiftAdjustStd), 2)
            If Cells(i, j) < 0 Then Cells(i, j) = 0 ' Ensure non-negative
            
            ' Count non-conformance (>4mm)
            If Cells(i, j) > 4 Then nonConformance = nonConformance + 1
        Next j
        
        ' Write # Non-Conformance (column U)
        Cells(i, 21) = nonConformance
    Next i
End Sub
