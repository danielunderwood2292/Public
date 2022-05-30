Attribute VB_Name = "Data_Cleaning_VBA"
Sub Data_Formatting()
'This VBA macro is designed to take an irregularly formatted text file from a bespoke analysis package
'that cannot be easily imported directly into excel
'and reformat it into easily usable excel content

'**************************************************************************************************
'Check number of open workbooks is 2
'**************************************************************************************************
'Variables for workbook and sheets names
Dim this_wb, file_wb, file_sheet As String
this_wb = ThisWorkbook.Name

'Number of open workbooks variable
Dim no_open_wbs As Integer
no_open_wbs = 0

'Variable for open workbook details
Dim wb As Workbook

'Loop around each open workbook
For Each wb In Workbooks

    If wb.FullName = ThisWorkbook.FullName Then
    Else
        file_wb = wb.Name
    End If
      
    no_open_workbooks = no_open_workbooks + 1
    
Next wb

'Check that only 2 workbooks are open - this and another
If no_open_workbooks > 2 Then
    
    MsgBox "There is more than one other workbook open. Cannot determine which contains the Polls data. In order for this macro to work correctly, only one other workbook may be open. This macro will now exit."
    GoTo Macro_Exit
      
ElseIf no_open_workbooks = 1 Then

    MsgBox "No raw file workbook is open. This macro will now exit."
    GoTo Macro_Exit
    
Else
    
End If

'**************************************************************************************************
'Find raw file worksheet name
'**************************************************************************************************
Workbooks(file_wb).Activate

Dim no_file_sheets As Integer
no_file_sheets = Worksheets.Count

If no_file_sheets > 1 Then
    MsgBox "The raw file workbook has more than one worksheet. Cannot determine which contains the data. This macro will now exit."
    GoTo Macro_Exit
Else

End If

file_sheet = ActiveSheet.Name

'**************************************************************************************************
'Add new worksheet for the output
'**************************************************************************************************
Workbooks(this_wb).Activate

Dim output As String
output = "Output"

'Check if an output sheet exists already
On Error GoTo Add_Output_Sheet
Sheets(output).Select
On Error GoTo 0

GoTo Output_Sheet_Variables

'Output sheet variables
Output_Sheet_Variables:
Dim output_heading_row, output_current_row, output_const_column, output_party_column, output_region_column, output_rawshare_column, output_2019share_column, output_totaloutput_column, output_adjshare_column, output_output_column As Integer
output_heading_row = 1
output_current_row = output_heading_row + 1
output_category_column = 1
output_device_column = 2
output_specified_column = 3
output_specunits_column = 4
output_calculated_column = 5
output_calcunits_column = 6

'Add headings to output sheet
Sheets(output).Cells(output_heading_row, output_category_column).Value = "Category"
Sheets(output).Cells(output_heading_row, output_device_column).Value = "Device"
Sheets(output).Cells(output_heading_row, output_specified_column).Value = "SpecifiedValue"
Sheets(output).Cells(output_heading_row, output_calculated_column).Value = "CalculatedValue"
Sheets(output).Cells(output_heading_row, output_specunits_column).Value = "SpecifiedUnits"
Sheets(output).Cells(output_heading_row, output_calcunits_column).Value = "CalcultatedUnits"

'**************************************************************************************************
'Find the number of rows in the file sheet
'**************************************************************************************************
Dim file_start_row, file_current_row, file_final_row, file_emptyrow_count, file_check_column As Integer
'The start row is set to 3 to avoid the first 2 rows of metadata in the file
file_start_row = 3
file_check_column = 2
file_final_row = file_start_row
file_current_row = file_start_row
file_emptyrow_count = 0

Do Until file_emptyrow_count = 3
    If Workbooks(file_wb).Sheets(file_sheet).Cells(file_current_row, file_check_column).Value = "" Then
        file_emptyrow_count = file_emptyrow_count + 1
    Else
        file_emptyrow_count = 0
    End If
    
    file_current_row = file_current_row + 1
    file_final_row = file_final_row + 1
Loop

'**************************************************************************************************
'Main loop to find info
'**************************************************************************************************
Dim no_cat_columns, file_current_column As Integer
Dim category As String
Dim content As String

For file_current_row = file_start_row To file_final_row
    
    'If the first 2 columns are empty then there is considered to be no relevant data in the file
    If IsEmpty(Workbooks(file_wb).Sheets(file_sheet).Cells(file_current_row, 1)) = True And IsEmpty(Workbooks(file_wb).Sheets(file_sheet).Cells(file_current_row, 2)) = True Then
    
    'If the first column has content then this is considered a 'category'
    ElseIf IsEmpty(Workbooks(file_wb).Sheets(file_sheet).Cells(file_current_row, 1)) = False Then
        'reset category name to blank
        category = ""
        
        'count the number of columns with content to compile the category name
        Workbooks(file_wb).Activate
        Sheets(file_sheet).Select
        no_cat_columns = WorksheetFunction.CountA(Rows(file_current_row))
        
        'Loop through these columns to compile the category name
        For file_current_column = 1 To no_cat_columns
            category = category & Workbooks(file_wb).Sheets(file_sheet).Cells(file_current_row, file_current_column).Value
        Next file_current_column
        
        category = Replace(category, ":", "")
        
        'Now loop through each of the devices under each category to find their information
        'Assume an empty row is the end of devices for this category
        Do While IsEmpty(Workbooks(file_wb).Sheets(file_sheet).Cells(file_current_row, 2)) = False
            file_current_row = file_current_row + 1
            content = Workbooks(file_wb).Sheets(file_sheet).Cells(file_current_row, 2).Value
            
            If content = "Specified" Or content = "Calculated" Or content = "Object" Or InStr(1, content, "-", vbTextCompare) <> 0 Or content = "" Then
                'Do nothing
            Else
                Workbooks(this_wb).Sheets(output).Cells(output_current_row, output_category_column).Value = category
                Workbooks(this_wb).Sheets(output).Cells(output_current_row, output_device_column).Value = content
                
                'count how many columns have content to see what should be done with the rest of the row info
                Workbooks(file_wb).Activate
                Sheets(file_sheet).Select
                no_cat_columns = WorksheetFunction.CountA(Rows(file_current_row))
                
                If no_cat_columns = 5 Then
                    Workbooks(this_wb).Sheets(output).Cells(output_current_row, output_specified_column).Value = Workbooks(file_wb).Sheets(file_sheet).Cells(file_current_row, 3).Value
                    Workbooks(this_wb).Sheets(output).Cells(output_current_row, output_specunits_column).Value = Workbooks(file_wb).Sheets(file_sheet).Cells(file_current_row, 4).Value
                    Workbooks(this_wb).Sheets(output).Cells(output_current_row, output_calculated_column).Value = Workbooks(file_wb).Sheets(file_sheet).Cells(file_current_row, 5).Value
                    Workbooks(this_wb).Sheets(output).Cells(output_current_row, output_calcunits_column).Value = Workbooks(file_wb).Sheets(file_sheet).Cells(file_current_row, 6).Value
                ElseIf no_cat_columns = 4 And Workbooks(file_wb).Sheets(file_sheet).Cells(file_current_row, 3).Value = "n/a" Then
                    Workbooks(this_wb).Sheets(output).Cells(output_current_row, output_calculated_column).Value = Workbooks(file_wb).Sheets(file_sheet).Cells(file_current_row, 4).Value
                    Workbooks(this_wb).Sheets(output).Cells(output_current_row, output_calcunits_column).Value = Workbooks(file_wb).Sheets(file_sheet).Cells(file_current_row, 5).Value
                ElseIf no_cat_columns = 4 Then
                    Workbooks(this_wb).Sheets(output).Cells(output_current_row, output_device_column).Value = content & "_" & Workbooks(file_wb).Sheets(file_sheet).Cells(file_current_row, 3).Value
                    Workbooks(this_wb).Sheets(output).Cells(output_current_row, output_calculated_column).Value = Workbooks(file_wb).Sheets(file_sheet).Cells(file_current_row, 4).Value
                    Workbooks(this_wb).Sheets(output).Cells(output_current_row, output_calcunits_column).Value = Workbooks(file_wb).Sheets(file_sheet).Cells(file_current_row, 5).Value
                ElseIf no_cat_columns = 3 Then
                    Workbooks(this_wb).Sheets(output).Cells(output_current_row, output_calculated_column).Value = Workbooks(file_wb).Sheets(file_sheet).Cells(file_current_row, 3).Value
                    Workbooks(this_wb).Sheets(output).Cells(output_current_row, output_calcunits_column).Value = Workbooks(file_wb).Sheets(file_sheet).Cells(file_current_row, 4).Value
                ElseIf no_cat_columns = 6 Then
                    For i = 3 To 5
                        content = content & "_" & Workbooks(file_wb).Sheets(file_sheet).Cells(file_current_row, i).Value
                    Next i
                    Workbooks(this_wb).Sheets(output).Cells(output_current_row, output_device_column).Value = content
                    Workbooks(this_wb).Sheets(output).Cells(output_current_row, output_calculated_column).Value = Workbooks(file_wb).Sheets(file_sheet).Cells(file_current_row, 6).Value
                    Workbooks(this_wb).Sheets(output).Cells(output_current_row, output_calcunits_column).Value = Workbooks(file_wb).Sheets(file_sheet).Cells(file_current_row, 7).Value
                Else
                    Workbooks(this_wb).Sheets(output).Cells(output_current_row, output_specified_column).Value = "ERROR CANNOT COMPUTE FILE CONTENT FOR THIS ROW"
                End If
                
                output_current_row = output_current_row + 1
            End If
        Loop
    
    Else
    End If
Next file_current_row

GoTo Macro_Exit

Add_Output_Sheet:
Sheets.Add(After:=Sheets("Instructions")).Name = output
GoTo Output_Sheet_Variables:

Macro_Exit:
Workbooks(this_wb).Activate
Sheets(output).Select
MsgBox "Data reformatting complete!'"

End Sub
