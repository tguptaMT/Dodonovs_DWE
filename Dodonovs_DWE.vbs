Sub DWE()
 
' ----------------------------------
' This visual basic subroutine allows you to choose a column of numerical values in an opened excel workbook, estimates 
' Dodonov's distance-weighted mean, SD and 95% confidence intervals for selected data and spits out those values
' at specified location within the workbook. The distance weighted estimates are based on following reference:
' Reference: Robust measures of central tendency: weighting as a possible alternative to trimming in response-time data analysis
' Yury S. Dodonov, Yulia A. Dodonova (Psikhologicheskie Issledovaniya. ISSN 2075-7999)
' Tarun Gupta, UVM (2016)
 'Implementation: As a macro in MS Excel.
' -----------------------------------
 
 
    Dim rRange As Variant
    Dim sub_down As Variant
    Dim sub_up As Variant
    Dim i As Variant
    Dim Cdown_val As Variant
    Dim Cup_val As Variant
    Dim sub_Total As Variant
    Dim Weight As Variant
    Dim NumberOfCells As Variant
    Dim Out_file As Variant
    Dim Sum_weighted_vals As Variant
    Dim Weighted_val() As Variant
    Dim Weights_All() As Variant
    Dim Weighted_Mean As Variant
    Dim Weighted_Stdev As Variant
    Dim val_Sub_wtMean() As Double
    Dim Cl_count As Variant
 
 
 
    On Error Resume Next
 
 
        Application.DisplayAlerts = False
 
            'Specificy Range of cells on Excel sheet
            Set rRange = Application.InputBox(Prompt:= _
                "Please select a range with your Mouse for calulating DWE.", _
                    Title:="SPECIFY RANGE for distance-weighted estimator", Type:=8)
 
 
    On Error GoTo 0
 
        Application.DisplayAlerts = True
 
 
        If rRange Is Nothing Then
           Exit Sub
        Else
            NumberOfCells = rRange.Cells.Count                'Counts # of cells in the specified Range
            MsgBox "Number of Cells is " & NumberOfCells
            Cell_count = 0
            Sum_weighted_vals = 0
 
 
            For Each Cl In rRange.Cells                        ' Loop through each element in the specified range
                C_val = Cl.Value
                sub_down = 0
                sub_up = 0
                Weight = 0
                Cdown_val = 0
                Cup_val = 0
                Cell_count = Cell_count + 1
                Weighted_Mean = 0
 
 
                For i = 1 To NumberOfCells - 1              'Subloop within the main loop to estimate distance of given value in relation to all other values
 
 
                    If rRange.Cells(i + 1) <> "" Then                                 '-- check for non-empty cells below current cell
                        If IsNumeric(rRange.Cells(i + 1)) = True Then                 '-- If the cell contains a numeric value
                            If i >= Cell_count Then                                   '-- If it's positioned below the current cell
                                Cdown_val = rRange.Cells(i + 1)
                                sub_down = sub_down + Abs(C_val - Cdown_val)
                            Else
                                sub_down = sub_down + 0
                            End If
                        End If
                    End If
              
                    If i < Cell_count Then
                        If rRange.Cells(i) <> "" Then
                            If IsNumeric(rRange.Cells(i)) = True Then
                                Cup_val = rRange.Cells(i)
                                sub_up = sub_up + Abs(C_val - Cup_val)
                            End If
                        End If
                    Else
                        sub_up = sub_up + 0
 
 
                    End If
                    sub_Total = sub_down + sub_up
                                                  
 
 
                Next i
 
 
                If sub_Total <> 0 Then
                    Weight_Cl = Weight_Cl + (((NumberOfCells - 1) / (sub_Total)))
                Else
                    Weight_Cl = Weight_Cl
                End If
 
 
                ReDim Preserve Weights_All(0 To NumberOfCells - 1)
                    Weights_All(Cell_count - 1) = Weight_Cl        'Array of all Weights
 
 
                ReDim Preserve Weighted_val(0 To NumberOfCells - 1)
                    Weighted_val(Cell_count - 1) = Cl.Value * Weight_Cl        'Array of all weighted values (value * weight)
 
            Next Cl
             
        Weighted_Mean = Application.WorksheetFunction.Sum(Weighted_val) / Application.WorksheetFunction.Sum(Weights_All)
 
        Cl_count = 0
 
            For Each Cl In rRange.Cells
                
                ReDim Preserve val_Sub_wtMean(0 To NumberOfCells - 1)
                    val_Sub_wtMean(Cl_count) = Weights_All(Cl_count) * ((Cl.Value - Weighted_Mean) * (Cl.Value - Weighted_Mean))
 
                Cl_count = Cl_count + 1
 
            Next Cl
 
        Weighted_Stdev = Sqr(Application.WorksheetFunction.Sum(val_Sub_wtMean) / Application.WorksheetFunction.Sum(Weights_All))
 
        MsgBox "Dodonov's Distance Weighted Mean for specified range is: " & Weighted_Mean & " +/- S.D. " & Weighted_Stdev
 
        ' Specify cells to output Distance weighted mean and SD
        Set OutRange = Application.InputBox(Prompt:= _
                "Please select two cells for writing DWE Mean and Sdev.", _
                    Title:="SPECIFY 2 Cells for writing DWE weighted Mean and Stdev", Type:=8)
 
 
        OutRange.Cells(1) = Weighted_Mean
        OutRange.Cells(2) = Weighted_Stdev
                             
         
        End If
End Sub
