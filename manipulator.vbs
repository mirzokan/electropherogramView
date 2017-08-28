Option Explicit

Sub import_folder_beckman()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False 

    Call import_folder("b")

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True 
End Sub

Sub import_folder_mce()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False 

    Call import_folder("m")

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True 
End Sub

Sub import_folder(imode As String)
    Dim i As Long
    Dim an_sheet As Worksheet, tem_sheet As Worksheet, imp_sheet As Worksheet
    Dim file_list() As Variant
    Dim file_path As String, file_path_location As Range

    Set an_sheet = ThisWorkbook.Sheets("Analysis")
    Set tem_sheet = ThisWorkbook.Sheets("Template")
    Set imp_sheet = ThisWorkbook.Sheets("Import")

    tem_sheet.Visible = xlSheetVisible
    Set file_path_location = an_sheet.Range("O7")

    Call validate_filepath(file_path_location)
    file_path = file_path_location.Value
    file_list = list_files(file_path)

    For i = 0 To UBound(file_list)
        Call import_file(file_path, file_list(i), i + 1, imode)
    Next i

    tem_sheet.Visible = xlSheetHidden

    Call import_sheet

    'Cleanup
    Erase file_list
End Sub

Sub import_file(file_path As String, filename As Variant, column_n As Long, imode As String)
 ' This macro will import a file into this workbook
    Dim an_sheet As Worksheet, tem_sheet As Worksheet, imp_sheet As Worksheet
    Dim points As Long, slash_pos As Long, full_char_count As Long, short_char_count As Long
    Dim source As Variant
    Dim full_filename As String, short_filename As String
    Dim errmess As String

    Set an_sheet = ThisWorkbook.Sheets("Analysis")
    Set tem_sheet = ThisWorkbook.Sheets("Template")
    Set imp_sheet = ThisWorkbook.Sheets("Import")

    On Error GoTo ErrHandler
    errmess = "Problem opening files. Please verify the file path." 
    Set source = Application.Workbooks.Open(file_path & filename, ReadOnly:=True)
    
      
    If imode = "b" Then
        points = source.Sheets(1).Range("B9").Value
    ElseIf imode = "m" Then
        points = source.Sheets(1).Range("A2:A" & source.Sheets(1).Range("A2").End(xlDown).Row).Count
    End If

    imp_sheet.Range("A1").Offset(0, column_n - 1).Value = filename
       
    If imode = "b" Then
    source.Sheets(1).Range("A14", source.Sheets(1).Range("A14").Offset(points, 0)).Copy Destination:=imp_sheet.Range("A2").Offset(0, column_n - 1)
    ElseIf imode = "m" Then
    source.Sheets(1).Range("A2", source.Sheets(1).Range("A2").Offset(points, 0)).Copy Destination:=imp_sheet.Range("A2").Offset(0, column_n - 1)
    End If
            
    Windows(filename).Activate
    ActiveWorkbook.Close SaveChanges:=False
    ThisWorkbook.Sheets("Analysis").Activate
        
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

        
    Exit Sub
    ' In case of error
        ErrHandler:
        MsgBox errmess
End Sub

Sub import_sheet()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False 

    'Error handling
    On Error GoTo Errorcatch

    Dim i As Long, existing_data As Long, import_data As Long, current_column As Long, point_num As Long
    Dim an_sheet As Worksheet, tem_sheet As Worksheet, imp_sheet As Worksheet
    Dim x_col As String, y_col As String, name_col As String

    Set an_sheet = ThisWorkbook.Sheets("Analysis")
    Set tem_sheet = ThisWorkbook.Sheets("Template")
    Set imp_sheet = ThisWorkbook.Sheets("Import")

    tem_sheet.Visible = xlSheetVisible

    If imp_sheet.Range("A1").Value = "Paste file names in the first row." Then
        MsgBox ("No data to import!")
        Exit Sub
    End If    
    
    import_data = col_count(imp_sheet.Range("A1"))

    For i = 1 To import_data
        existing_data = an_sheet.Range("O22").Value
        current_column = 1 + (existing_data * 4)

        'Manipulations
        tem_sheet.Range("A29:D38").Copy Destination:=an_sheet.Cells(29, current_column)
        imp_sheet.Range("A1").Copy
        an_sheet.Cells(31, current_column + 1).PasteSpecial Paste:=xlPasteValues

        point_num = imp_sheet.Range("A2:A" & imp_sheet.Range("A2").End(xlDown).Row).Count

        'Copy Data
        imp_sheet.Range("A2:A" & imp_sheet.Range("A2").End(xlDown).Row).Copy
        an_sheet.Cells(36, current_column + 1).PasteSpecial Paste:=xlPasteValues

        'Fill up fields
        With an_sheet
           .Cells(29, current_column + 1).Value = "=Average(Indirect(" & Chr(34) & column_letter(current_column + 3) & Chr(34) & "&bg_start&" & Chr(34) & ":" & column_letter(current_column + 3) & Chr(34) & "&bg_end))"
           .Cells(29, current_column + 3).Value = "=STDEV(Indirect(" & Chr(34) & column_letter(current_column + 3) & Chr(34) & "&bg_start&" & Chr(34) & ":" & column_letter(current_column + 3) & Chr(34) & "&bg_end))"
           .Cells(30, current_column + 1).Value = "=MAX(" & column_letter(current_column + 3) & "36:" & column_letter(current_column + 3) & (point_num + 35) & ")-" & column_letter(current_column + 1) & "29"
           .Cells(30, current_column + 3).Value = "=" & column_letter(current_column + 1) & "30/" & column_letter(current_column + 3) & "29"
           
           .Cells(37, current_column).Copy Destination:=Range(.Cells(38, current_column), .Cells(point_num + 35, current_column))
           .Cells(36, current_column + 2).Copy Destination:=Range(.Cells(37, current_column + 2), .Cells(point_num + 35, current_column + 2))
            

            'Filter
            If Range("med_filter").Value = "3-point Median" Then
                .Cells(37, current_column + 3).Copy Destination:=Range(.Cells(38, current_column + 3), .Cells(point_num + 35, current_column + 3))
            ElseIf Range("med_filter").Value = "5-point Median" Then
                .Cells(38, current_column + 3).Copy Destination:=Range(.Cells(39, current_column + 3), .Cells(point_num + 35, current_column + 3))
            ElseIf Range("med_filter").Value = "SovGol" Then
                Range(.Cells(36, current_column + 3), .Cells(point_num + 35, current_column + 3)).FormulaArray="=(sovgol(" & Range(.Cells(36, current_column + 1), .Cells(point_num + 35, current_column + 1)).Address & ", $V$10,$V$11)*" & .Cells(33, current_column + 3).Address & ")+" & .Cells(33, current_column + 1).Address
            ElseIf Range("med_filter").Value = "None" Then
                .Cells(36, current_column + 3).Copy Destination:=Range(.Cells(37, current_column + 3), .Cells(point_num + 35, current_column + 3))
            End If
        End With

        imp_sheet.Columns(1).Delete xlShiftToLeft
        an_sheet.Range("O22").Value = existing_data + 1
    Next i

    Call restack

    'Chart
    an_sheet.ChartObjects("Chart 1").Activate

    existing_data = an_sheet.Range("O22").Value

    For i = 0 To existing_data - 1
        current_column = 3 + (i * 4)

        With an_sheet
            point_num = .Cells(36, current_column).End(xlDown).Row
            name_col = .Cells(31, current_column - 1).Address
            x_col = .Range(.Cells(36, current_column), .Cells(point_num, current_column)).Address
            y_col = .Range(.Cells(36, current_column + 1), .Cells(point_num, current_column + 1)).Address
        End With

        ActiveChart.SeriesCollection.NewSeries
        ActiveChart.SeriesCollection(i + 1).MarkerStyle = xlMarkerStyleNone
        ActiveChart.SeriesCollection(i + 1).Name = an_sheet.Range(name_col)
        ActiveChart.SeriesCollection(i + 1).XValues = an_sheet.Range(x_col)
        ActiveChart.SeriesCollection(i + 1).Values = an_sheet.Range(y_col)
    Next i

    ThisWorkbook.Sheets("Template").Visible = xlSheetHidden

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True 

    MsgBox ("Import Complete!")
    
    Exit Sub
    Errorcatch:
        MsgBox Err.Description
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayAlerts = True 
End Sub

Sub export_folder()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = False 

    ThisWorkbook.Sheets("Template").Visible = xlSheetVisible
    ThisWorkbook.Sheets("Export").Visible = xlSheetVisible

    Dim i As Long, existing_data As Long, current_column As Long, sampling_rate As Double, point_num As Variant, fin_count As Long
    Dim an_sheet As Worksheet, tem_sheet As Worksheet, imp_sheet As Worksheet
    Dim export_path As String, file_path_location As Range, file_name As String, ext As String
    Dim NewBook As Workbook

    Dim exp_sheet As Worksheet
    Dim downfill As Integer
    Set exp_sheet = ThisWorkbook.Sheets("Export")
    exp_sheet.Cells.Clear

    Set an_sheet = ThisWorkbook.Sheets("Analysis")
    Set tem_sheet = ThisWorkbook.Sheets("Template")
    Set imp_sheet = ThisWorkbook.Sheets("Import")

    If Range("exp_scale") = "N" Then
        'Stacking removal
        an_sheet.Range("Y_stack").Value = 0
        an_sheet.Range("X_stack").Value = 0
        Call restack
    End If

    Set file_path_location = an_sheet.Range("O16")
    Call validate_filepath(file_path_location)
    export_path = an_sheet.Range("O16").Value

    existing_data = an_sheet.Range("O22").Value
    sampling_rate = an_sheet.Range("Sampling_rate").Value

    Application.DisplayAlerts = False

    For i = 0 To existing_data - 1
        current_column = (i * 4)
        file_name = an_sheet.Cells(31, current_column + 2).Value
        point_num = an_sheet.Range("A36").Offset(0, current_column + 3).End(xlDown).Address
            
        Set NewBook = Workbooks.Add
            With NewBook
                .Title = file_name
                .Sheets(1).Range("A1").Value = "File exported using Electropherogram Manipulator"
                .Sheets(1).Range("A4").Value = "Data File:"
                .Sheets(1).Range("B4").Value = file_name
                .Sheets(1).Range("A8").Value = "Sampling Rate:"
                .Sheets(1).Range("B8").Value = sampling_rate
                .Sheets(1).Range("B11").Value = "RFU"
                If an_sheet.Range("exp_scale") = "Y" Then
                    .Sheets(1).Range("C11").Value = "µA"
                End If
                .Sheets(1).Range("A9").Value = "Total Data Points:"
                .Sheets(1).Range("A13").Value = "Signal (AU)"
            End With
            
            an_sheet.Calculate
            NewBook.Sheets(1).Calculate
            
            
            If an_sheet.Range("exp_scale") = "Y" Then

                an_sheet.Range(an_sheet.Range("A36").Offset(0, current_column), point_num).Copy
                exp_sheet.Range("A1").PasteSpecial Paste:=xlPasteValues
                downfill = exp_sheet.Range("A1").End(xlDown).Row
                 
                exp_sheet.Range("E1").Formula = "=IFERROR(LOOKUP(A1,$C$1:$C$" & downfill & ",$D$1:$D$" & downfill & "), 0)"
                exp_sheet.Range(exp_sheet.Cells(1, 5), exp_sheet.Cells(downfill, 5)).FillDown
                  
                exp_sheet.Range(exp_sheet.Cells(1, 5), exp_sheet.Cells(downfill, 5)).Copy
                exp_sheet.Range("F1").PasteSpecial Paste:=xlPasteValues
                
                exp_sheet.Range(exp_sheet.Cells(1, 1), exp_sheet.Cells(downfill, 5)).Delete Shift:=xlToLeft
                
                exp_sheet.Range(exp_sheet.Cells(1, 1), exp_sheet.Cells(exp_sheet.Cells(1, 1).End(xlDown).Row, 1)).Copy
                NewBook.Sheets(1).Range("A14").PasteSpecial Paste:=xlPasteValues
                
                exp_sheet.Cells.Clear
            Else
                an_sheet.Range(an_sheet.Range("A36").Offset(0, current_column + 3), point_num).Copy
                NewBook.Sheets(1).Range("A14").PasteSpecial Paste:=xlPasteValues
            End If
                    
            fin_count = NewBook.Sheets(1).Range(NewBook.Sheets(1).Cells(14, 1), NewBook.Sheets(1).Cells(14, 1).End(xlDown).Address).Count
            NewBook.Sheets(1).Range("B9").Value = fin_count
            ' Current
            If an_sheet.Range("exp_cur") = "Y" Then
                NewBook.Sheets(1).Range(NewBook.Sheets(1).Cells(14+fin_count, 1), NewBook.Sheets(1).Cells(14+(fin_count*2), 1)).Value = an_sheet.Range("cur_val").Value * 100000
            End If

            ext = Right(file_name, 4)
            If ext = ".csv" Then
                NewBook.SaveAs filename:=export_path & file_name, FileFormat:=xlCSV
            Else
                NewBook.SaveAs filename:=export_path & file_name & ".csv", FileFormat:=xlCSV
            End If
            NewBook.Close SaveChanges:=True

    Next i

    Application.DisplayAlerts = True

    ThisWorkbook.Sheets("Template").Visible = xlSheetHidden
    ThisWorkbook.Sheets("Export").Visible = xlSheetHidden
    

    ThisWorkbook.Sheets("Analysis").Activate
    ThisWorkbook.Sheets("Analysis").Range("A1").Select

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True 

    MsgBox ("Export Complete!")
End Sub


Sub restack()
    Dim an_sheet As Worksheet
    Dim i As Long, existing_data As Long, current_column As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False 

    Set an_sheet = ThisWorkbook.Sheets("Analysis")

    existing_data = an_sheet.Range("O22").Value
    an_sheet.Range("A33:B33").Value = 0

    If Range("Y_stack").Value <> 0 Then
        For i = 1 To existing_data - 1
            current_column = 1 + (i * 4)
            With an_sheet
                .Cells(33, current_column + 1).Formula = "=(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-4))+Y_stack"
            End With
        Next i
    Else
        For i = 1 To existing_data - 1
            current_column = 1 + (i * 4)
            With an_sheet
                .Cells(33, current_column + 1).Value = 0
            End With
        Next i
    End If

    If Range("Y_scale").Value <> 1 Then
        For i = 0 To existing_data - 1
            current_column = 1 + (i * 4)
            With an_sheet
                .Cells(33, current_column + 3).Formula = "=Y_scale"
            End With
        Next i
    Else
        For i = 0 To existing_data - 1
            current_column = 1 + (i * 4)
            With an_sheet
               .Cells(33, current_column + 3).Value = 1
            End With
        Next i
    End If


    If Range("X_stack").Value <> 0 Then
        For i = 1 To existing_data - 1
            current_column = 1 + (i * 4)
            With an_sheet
               .Cells(33, current_column).Formula = "=(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-4))+X_stack"
            End With
        Next i
    Else
        For i = 1 To existing_data - 1
            current_column = 1 + (i * 4)
            With an_sheet
                .Cells(33, current_column).Value = 0
            End With
        Next i
    End If


    If Range("X_scale").Value <> 1 Then
        For i = 0 To existing_data - 1
            current_column = 1 + (i * 4)
            With an_sheet
                .Cells(33, current_column + 2).Formula = "=X_scale"
            End With
        Next i
    Else
        For i = 0 To existing_data - 1
            current_column = 1 + (i * 4)
            With an_sheet
               .Cells(33, current_column + 2).Value = 1
            End With
        Next i
    End If



    If Range("bars").Value <> "N" Then
        For i = 0 To existing_data - 1
            current_column = 1 + (i * 4)
            With an_sheet
                Range(.Cells(36, current_column + 1), .Cells(.Cells(36, current_column + 1).End(xlDown).Row, current_column + 1)).FormatConditions.AddDatabar
                Range(.Cells(36, current_column + 1), .Cells(.Cells(36, current_column + 1).End(xlDown).Row, current_column + 1)).FormatConditions.AddTop10
                Range(.Cells(36, current_column + 1), .Cells(.Cells(36, current_column + 1).End(xlDown).Row, current_column + 1)).FormatConditions(2).Rank = 1
                Range(.Cells(36, current_column + 1), .Cells(.Cells(36, current_column + 1).End(xlDown).Row, current_column + 1)).FormatConditions(2).Interior.ColorIndex = 3
           End With
        Next i
    Else
        For i = 0 To existing_data - 1
            current_column = 1 + (i * 4)
            With an_sheet
                Range(.Cells(36, current_column + 1), .Cells(.Cells(36, current_column + 1).End(xlDown).Row, current_column + 1)).FormatConditions.Delete
           End With
        Next i
    End If
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True 

End Sub

Sub reset_all()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False 

    Dim an_sheet As Worksheet, imp_sheet As Worksheet
    Dim i As Long

    Set imp_sheet = ThisWorkbook.Sheets("Import")
    Set an_sheet = ThisWorkbook.Sheets("Analysis")
    an_sheet.Range("O22").Value = 0
    an_sheet.Range("A31").CurrentRegion.Clear

    an_sheet.ChartObjects("Chart 1").Activate
    With ActiveChart
        For i = .SeriesCollection.Count To 1 Step -1
             .SeriesCollection(i).Delete
        Next i
    End With

    imp_sheet.Cells.Clear
    imp_sheet.Range("A1").Value = "Paste file names in the first row."
    imp_sheet.Range("A2").Value = "Signal values start in the second row."

    Dim exp_sheet As Worksheet
    Set exp_sheet = ThisWorkbook.Sheets("Export")
    exp_sheet.Cells.Clear


    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True  

    MsgBox ("Reset complete!")
End Sub

Function column_letter(iCol As Integer) As String
    column_letter = Split(Cells(1, iCol).Address, "$")(1)
End Function

Function col_count(base_cell As Range) As Integer
    Dim row_num As Integer
    Dim ws As Worksheet

    Set ws = base_cell.Parent
    row_num = base_cell.row
    col_count = Application.WorksheetFunction.CountA(ws.Range(row_num & ":" & row_num)) 
End Function

Function sovgol (rrange As Range, points As Integer, Optional derOrder As Integer = 0) As Variant
 ' An array function that smoothes a range of equally spaced data using Savitzky-Golay method
 'Description of Arguments:
 'rrange = Contiguous range of data to be smoothed
 'points = Number of points for a running filter (must be an odd number larger than 3)
 'derOrder = (Optional) Order of derivation 0-2


    Dim m As Integer, i As Integer, j As Integer
    Dim c As Variant
    Dim rsize As Long
    Dim sumprod As Double
    Dim p() As Double
    Dim original() As Variant, smoothed() As Variant

    'Check if $points is odd, and convert to m
    If (points Mod 2 <> 1) Or (points < 3) Then
        sovgol = CVErr(xlErrValue)
        Exit Function
    Else 
        m = (points-1)/2
    End If
    'Create p array
    ReDim p(-m To m)

    ' Pick the right equation to define the coeficients
    Select Case derOrder
    Case 0
        For i=-m To m
            p(i)= (3*(3*(m^2)+3*m-1-5*(i^2)))/((2*m+3)*(2*m+1)*(2*m-1))
        Next i
    Case 1
        For i=-m To m
            p(i)= (3*i)/((2*m+1)*(m+1)*m)
        Next i
    Case 2
        For i=-m To m
            p(i)= (30*(3*i^2-m*(m+1)))/((2*m+3)*(2*m+1)*(2*m-1)*(m+1)*m)
        Next i
    Case Else
        sovgol = CVErr(xlErrValue)
        Exit Function
    End Select


    'Count number of points in $rrange
    rsize = Application.WorksheetFunction.CountA(rrange)
    ReDim original(1 To rsize)
    ReDim smoothed(1 To rsize, 1 to 1)
    ' Import range
    For i=1 To rsize
        original(i) = rrange.Cells(i,1).Value
    Next i

    'Start with m'th point
    For i=m+1 to rsize-m
        sumprod = 0
        For j=-m To m
            sumprod = sumprod + (original(i+j)*p(j))
        Next j
        smoothed(i,1)=sumprod
    Next i

    'Fill in missing points at start
    For i=1 to m
        smoothed(i,1) = smoothed(m+1,1)
    Next i
    'Fill in missing points at end
    For i=(rsize-(m-1)) to rsize
        smoothed(i,1) = smoothed(rsize-m,1)
    Next i
    
    sovgol = smoothed
End Function


