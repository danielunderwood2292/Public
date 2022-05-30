Attribute VB_Name = "Module1"
Sub baseline_compare()
'Version 0.4
'Date: 03/12/18
'By: Dan Underwood
'Macro to compare two different versions of exported requirements
'For an overall description of what this macro does, see the 'Instructions' worksheet

'**********************************************************************************************
'Timer
'**********************************************************************************************
'Set up a timer to record how long the macro takes to run
Dim StartTime As Double
Dim SecondsElapsed As Double

'Remember time when macro starts
StartTime = Timer

'**********************************************************************************************
'Speed Up Macro
'**********************************************************************************************
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

'**********************************************************************************************
'Check Workbooks Are Open
'**********************************************************************************************
Dim previous_wb, current_wb, this_wb, current_reqheading, previous_reqheading, current_sheet, previous_sheet As String
current_wb = Range("Source_Name").Value
previous_wb = Range("Target_Name").Value
current_reqheading = Range("current_reqheading").Value
previous_reqheading = Range("previous_reqheading").Value
current_sheet = Range("Current_Sheet").Value
previous_sheet = Range("Previous_Sheet").Value
this_wb = ActiveWorkbook.Name

On Error GoTo Err1:
Workbooks(previous_wb).Activate

On Error GoTo Err11:
Sheets(current_sheet).Activate

On Error GoTo Err2:
Workbooks(current_wb).Activate

On Error GoTo Err21:
Sheets(previous_sheet).Activate

On Error GoTo 0

'Test msg
'MsgBox "This workbook is: " & this_wb

'**********************************************************************************************
'Main Variables Declaration
'**********************************************************************************************
'Results Variables
Workbooks(this_wb).Activate
Dim results_sheet As String
results_sheet = "Results"

Sheets(results_sheet).Select

Dim results_title_row As Integer
results_title_row = 1

Dim results_first_row As Integer
results_first_row = 2

Dim results_id_column As Integer
search_term = "Reqt ID"
Call Search(search_term, search_row, search_column)
results_id_column = search_column

Dim results_final_row As Integer
results_final_row = WorksheetFunction.CountA(Columns(results_id_column)) + 1

'Dim results_satisfies_column As Integer
'search_term = "Satisfies [LINKED]"
'Call Search(search_term, search_row, search_column)
'results_satisfies_column = search_column

Dim results_source_column As Integer
search_term = "Source"
Call Search(search_term, search_row, search_column)
results_source_column = search_column

Dim results_type_column As Integer
search_term = "Object Type"
Call Search(search_term, search_row, search_column)
results_type_column = search_column

'Dim results_section_column As Integer
'search_term = "Section"
'Call Search(search_term, search_row, search_column)
'results_section_column = search_column

Dim results_title_column As Integer
search_term = "Requirement Title"
Call Search(search_term, search_row, search_column)
results_title_column = search_column

Dim results_rationale_column As Integer
search_term = "Rationale"
Call Search(search_term, search_row, search_column)
results_rationale_column = search_column

Dim results_maturity_column As Integer
search_term = "Requirement Maturity"
Call Search(search_term, search_row, search_column)
results_maturity_column = search_column

Dim results_comments_column As Integer
search_term = "Comments"
Call Search(search_term, search_row, search_column)
results_comments_column = search_column

Dim results_acceptance_column As Integer
search_term = "Acceptance Criteria"
Call Search(search_term, search_row, search_column)
results_acceptance_column = search_column

'Dim results_satarg_column As Integer
'search_term = "Satisfaction Statement"
'Call Search(search_term, search_row, search_column)
'results_satarg_column = search_column

'Dim results_satisfied_column As Integer
'search_term = "Satisfied By [LINKED]"
'Call Search(search_term, search_row, search_column)
'results_satisfied_column = search_column

Dim results_req_column As Integer
search_term = "Requirement Text"
Call Search(search_term, search_row, search_column)
results_req_column = search_column

Dim deleted_colour As String
deleted_colour = RGB(0, 128, 0)

Dim new_colour As String
new_colour = RGB(255, 0, 0)

Dim change_colour As String
change_colour = RGB(0, 112, 192)

Dim current_text As String


'Current Variables
Workbooks(current_wb).Activate

Dim current_title_row As Integer
current_title_row = 1

Dim current_first_row As Integer
current_first_row = 2

Dim current_id_column As Integer
search_term = "ID"
Call Search(search_term, search_row, search_column)
current_id_column = search_column

Dim current_final_row As Integer
current_final_row = WorksheetFunction.CountA(Columns(current_id_column)) + 1

'Dim current_satisfies_column As Integer
'search_term = "Satisfies [LINKED]"
'Call Search(search_term, search_row, search_column)
'current_satisfies_column = search_column

Dim current_source_column As Integer
search_term = "Requirement Source"
Call Search(search_term, search_row, search_column)
current_source_column = search_column

Dim current_type_column As Integer
search_term = "Object Type"
Call Search(search_term, search_row, search_column)
current_type_column = search_column

'Dim current_section_column As Integer
'search_term = "Section"
'Call Search(search_term, search_row, search_column)
'current_section_column = search_column

Dim current_title_column As Integer
search_term = "Title"
Call Search(search_term, search_row, search_column)
current_title_column = search_column

Dim current_rationale_column As Integer
search_term = "Rationale"
Call Search(search_term, search_row, search_column)
current_rationale_column = search_column

Dim current_maturity_column As Integer
search_term = "Requirement Maturity"
Call Search(search_term, search_row, search_column)
current_maturity_column = search_column

Dim current_comments_column As Integer
search_term = "Comments"
Call Search(search_term, search_row, search_column)
current_comments_column = search_column

Dim current_acceptance_column As Integer
search_term = "Acceptance Criterion"
Call Search(search_term, search_row, search_column)
current_acceptance_column = search_column

'Dim current_satarg_column As Integer
'search_term = "Satisfaction Statement"
'Call Search(search_term, search_row, search_column)
'current_satarg_column = search_column

'Dim current_satisfied_column As Integer
'search_term = "Satisfied By [LINKED]"
'Call Search(search_term, search_row, search_column)
'current_satisfied_column = search_column

Dim current_req_column As Integer
search_term = current_reqheading
Call Search(search_term, search_row, search_column)
current_req_column = search_column

'Previous Variables
Workbooks(previous_wb).Activate

Dim previous_title_row As Integer
previous_title_row = 1

Dim previous_first_row As Integer
previous_first_row = 2

Dim previous_id_column As Integer
search_term = "ID"
Call Search(search_term, search_row, search_column)
previous_id_column = search_column

Dim previous_final_row As Integer
previous_final_row = WorksheetFunction.CountA(Columns(previous_id_column)) + 1

'Dim previous_satisfies_column As Integer
'search_term = "Satisfies [LINKED]"
'Call Search(search_term, search_row, search_column)
'previous_satisfies_column = search_column

Dim previous_source_column As Integer
search_term = "Requirement Source"
Call Search(search_term, search_row, search_column)
previous_source_column = search_column

Dim previous_type_column As Integer
search_term = "Object Type"
Call Search(search_term, search_row, search_column)
previous_type_column = search_column

'Dim previous_section_column As Integer
'search_term = "Section"
'Call Search(search_term, search_row, search_column)
'previous_section_column = search_column

Dim previous_title_column As Integer
search_term = "Title"
Call Search(search_term, search_row, search_column)
previous_title_column = search_column

Dim previous_rationale_column As Integer
search_term = "Rationale"
Call Search(search_term, search_row, search_column)
previous_rationale_column = search_column

Dim previous_maturity_column As Integer
search_term = "Requirement Maturity"
Call Search(search_term, search_row, search_column)
previous_maturity_column = search_column

Dim previous_comments_column As Integer
search_term = "Comments"
Call Search(search_term, search_row, search_column)
previous_comments_column = search_column

Dim previous_acceptance_column As Integer
search_term = "Acceptance Criterion"
Call Search(search_term, search_row, search_column)
previous_acceptance_column = search_column

'Dim previous_satarg_column As Integer
'search_term = "Satisfaction Statement"
'Call Search(search_term, search_row, search_column)
'previous_satarg_column = search_column

'Dim previous_satisfied_column As Integer
'search_term = "Satisfied By [LINKED]"
'Call Search(search_term, search_row, search_column)
'previous_satisfied_column = search_column

Dim previous_req_column As Integer
search_term = previous_reqheading
Call Search(search_term, search_row, search_column)
previous_req_column = search_column


'**********************************************************************************************
'Delete old data in results workbook
'**********************************************************************************************
Workbooks(this_wb).Activate
Sheets(results_sheet).Select

'Find how many rows are currently filled
Dim no_filled_rows As Integer
no_filled_rows = WorksheetFunction.CountA(Columns(results_id_column)) + 1

'Test message
'MsgBox "No filled rows: " & no_filled_rows

If no_filled_rows > 2 Then
    Range(Cells(results_first_row, 1), Cells(no_filled_rows, 9)).ClearContents
    Range(Cells(results_first_row, 1), Cells(no_filled_rows, 9)).Interior.Color = RGB(256, 256, 256)
Else


End If

'End


'**********************************************************************************************
'Loop around all objects in current workbook
'**********************************************************************************************
Dim current_current_row As Integer
Dim previous_current_row As Integer
Dim results_current_row As Integer
results_current_row = results_first_row

Dim current_id As String

Dim req_found As Boolean
req_found = False

For current_current_row = current_first_row To current_final_row
    current_id = Workbooks(current_wb).Sheets(current_sheet).Cells(current_current_row, current_id_column).Value
    
    'paste in the current details into the results worksheet
    'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satisfies_column).Value = Workbooks(current_wb).Sheets(current_sheet).Cells(current_current_row, current_satisfies_column).Value
    Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_id_column).Value = Workbooks(current_wb).Sheets(current_sheet).Cells(current_current_row, current_id_column).Value
    Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_source_column).Value = Workbooks(current_wb).Sheets(current_sheet).Cells(current_current_row, current_source_column).Value
    Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_type_column).Value = Workbooks(current_wb).Sheets(current_sheet).Cells(current_current_row, current_type_column).Value
    'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_section_column).Value = Workbooks(current_wb).Sheets(current_sheet).Cells(current_current_row, current_section_column).Value
    Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_title_column).Value = Workbooks(current_wb).Sheets(current_sheet).Cells(current_current_row, current_title_column).Value
    Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_req_column).Value = Workbooks(current_wb).Sheets(current_sheet).Cells(current_current_row, current_req_column).Value
    Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_rationale_column).Value = Workbooks(current_wb).Sheets(current_sheet).Cells(current_current_row, current_rationale_column).Value
    Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_maturity_column).Value = Workbooks(current_wb).Sheets(current_sheet).Cells(current_current_row, current_maturity_column).Value
    Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_comments_column).Value = Workbooks(current_wb).Sheets(current_sheet).Cells(current_current_row, current_comments_column).Value
    Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_acceptance_column).Value = Workbooks(current_wb).Sheets(current_sheet).Cells(current_current_row, current_acceptance_column).Value
    'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satarg_column).Value = Workbooks(current_wb).Sheets(current_sheet).Cells(current_current_row, current_satarg_column).Value
    'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satisfied_column).Value = Workbooks(current_wb).Sheets(current_sheet).Cells(current_current_row, current_satisfied_column).Value
    
    'look for this req id in the previous wb
    req_found = False
    
    For previous_current_row = previous_first_row To previous_final_row
        If Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_id_column).Value = current_id Then
            req_found = True
            Exit For
        Else
        End If
        
    Next previous_current_row
    
    'now format the cells in the results workbook based on the search for the id in the previous wb
    If req_found = True Then 'need to compare all cells to see if they match
    
        'If Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satisfies_column).Value <> Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_satisfies_column).Value Then
            'current_text = Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satisfies_column).Value
            'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satisfies_column).Value = "Latest version:" & vbNewLine & current_text & vbNewLine & vbNewLine & "Previous version:" & vbNewLine & Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_satisfies_column).Value
            'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satisfies_column).Interior.Color = change_colour
            'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_id_column).Interior.Color = change_colour
        'Else 'Else do nothing
        'End If
        
        If Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_source_column).Value <> Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_source_column).Value Then
            current_text = Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_source_column).Value
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_source_column).Value = "Latest version:" & vbNewLine & current_text & vbNewLine & vbNewLine & "Previous version:" & vbNewLine & Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_source_column).Value
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_source_column).Interior.Color = change_colour
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_id_column).Interior.Color = change_colour
        Else 'Else do nothing
        End If
        
        If Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_type_column).Value <> Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_type_column).Value Then
            current_text = Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_type_column).Value
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_type_column).Value = "Latest version:" & vbNewLine & current_text & vbNewLine & vbNewLine & "Previous version:" & vbNewLine & Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_type_column).Value
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_type_column).Interior.Color = change_colour
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_id_column).Interior.Color = change_colour
        Else 'Else do nothing
        End If
        
        'If Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_section_column).Value <> Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_section_column).Value Then
            'current_text = Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_section_column).Value
            'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_section_column).Value = "Latest version:" & vbNewLine & current_text & vbNewLine & vbNewLine & "Previous version:" & vbNewLine & Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_section_column).Value
            'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_section_column).Interior.Color = change_colour
            'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_id_column).Interior.Color = change_colour
        'Else 'Else do nothing
        'End If
        
        If Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_title_column).Value <> Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_title_column).Value Then
            current_text = Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_title_column).Value
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_title_column).Value = "Latest version:" & vbNewLine & current_text & vbNewLine & vbNewLine & "Previous version:" & vbNewLine & Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_title_column).Value
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_title_column).Interior.Color = change_colour
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_id_column).Interior.Color = change_colour
        Else 'Else do nothing
        End If
        
        If Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_req_column).Value <> Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_req_column).Value Then
            current_text = Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_req_column).Value
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_req_column).Value = "Latest version:" & vbNewLine & current_text & vbNewLine & vbNewLine & "Previous version:" & vbNewLine & Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_req_column).Value
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_req_column).Interior.Color = change_colour
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_id_column).Interior.Color = change_colour
        Else 'Else do nothing
        End If
        
        If Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_rationale_column).Value <> Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_rationale_column).Value Then
            current_text = Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_rationale_column).Value
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_rationale_column).Value = "Latest version:" & vbNewLine & current_text & vbNewLine & vbNewLine & "Previous version:" & vbNewLine & Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_rationale_column).Value
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_rationale_column).Interior.Color = change_colour
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_id_column).Interior.Color = change_colour
        Else 'Else do nothing
        End If
        
        If Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_maturity_column).Value <> Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_maturity_column).Value Then
            current_text = Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_maturity_column).Value
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_maturity_column).Value = "Latest version:" & vbNewLine & current_text & vbNewLine & vbNewLine & "Previous version:" & vbNewLine & Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_maturity_column).Value
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_maturity_column).Interior.Color = change_colour
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_id_column).Interior.Color = change_colour
        Else 'Else do nothing
        End If
        
        If Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_comments_column).Value <> Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_comments_column).Value Then
            current_text = Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_comments_column).Value
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_comments_column).Value = "Latest version:" & vbNewLine & current_text & vbNewLine & vbNewLine & "Previous version:" & vbNewLine & Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_comments_column).Value
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_comments_column).Interior.Color = change_colour
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_id_column).Interior.Color = change_colour
        Else 'Else do nothing
        End If
        
        If Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_acceptance_column).Value <> Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_acceptance_column).Value Then
            current_text = Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_acceptance_column).Value
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_acceptance_column).Value = "Latest version:" & vbNewLine & current_text & vbNewLine & vbNewLine & "Previous version:" & vbNewLine & Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_acceptance_column).Value
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_acceptance_column).Interior.Color = change_colour
            Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_id_column).Interior.Color = change_colour
        Else 'Else do nothing
        End If
        
        'If Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satarg_column).Value <> Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_satarg_column).Value Then
            'current_text = Workbooks(this_wb).Sheets(results_sheet).Cells(results_satarg_row, results_satisfies_column).Value
            'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satarg_column).Value = "Latest version:" & vbNewLine & current_text & vbNewLine & vbNewLine & "Previous version:" & vbNewLine & Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_satarg_column).Value
            'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satarg_column).Interior.Color = change_colour
            'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_id_column).Interior.Color = change_colour
        'Else 'Else do nothing
        'End If
        
        'If Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satisfied_column).Value <> Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_satisfied_column).Value Then
            'current_text = Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satisfied_column).Value
            'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satisfied_column).Value = "Latest version:" & vbNewLine & current_text & vbNewLine & vbNewLine & "Previous version:" & vbNewLine & Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_satisfied_column).Value
            'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satisfied_column).Interior.Color = change_colour
            'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_id_column).Interior.Color = change_colour
        'Else 'Else do nothing
        'End If
    
    Else 'requirement was not found so is new
        'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satisfies_column).Interior.Color = new_colour
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_id_column).Interior.Color = new_colour
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_source_column).Interior.Color = new_colour
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_type_column).Interior.Color = new_colour
        'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_section_column).Interior.Color = new_colour
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_title_column).Interior.Color = new_colour
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_req_column).Interior.Color = new_colour
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_rationale_column).Interior.Color = new_colour
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_maturity_column).Interior.Color = new_colour
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_comments_column).Interior.Color = new_colour
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_acceptance_column).Interior.Color = new_colour
        'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satarg_column).Interior.Color = new_colour
        'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satisfied_column).Interior.Color = new_colour
    
    End If

    results_current_row = results_current_row + 1
    
Next current_current_row


'**********************************************************************************************
'Loop around all objects in previous workbook - to find deletions
'**********************************************************************************************

Dim previous_id As String

For previous_current_row = previous_first_row To previous_final_row
    
    current_id = Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_id_column).Value
    
    'look for this req id in the previous wb
    req_found = False
    results_final_row = WorksheetFunction.CountA(Columns(results_id_column)) + 1
    
    For results_current_row = results_first_row To results_final_row
        If Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_id_column).Value = current_id Then
            req_found = True
            Exit For
        Else
        End If
        
    Next results_current_row
    
    If req_found = False Then
        'find the row to insert the deleted row under
        previous_id = Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row - 1, previous_id_column).Value
        
        For results_current_row = results_first_row To results_final_row
            If Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_id_column).Value = previous_id Then
                Exit For
            Else
        End If
               
        Next results_current_row
        
        'insert new blank row
        Workbooks(this_wb).Activate
        results_current_row = results_current_row + 1
        Workbooks(this_wb).Sheets(results_sheet).Rows(results_current_row).Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
       
        'paste in old requirement details
        'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satisfies_column).Value = Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_satisfies_column).Value
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_id_column).Value = Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_id_column).Value
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_source_column).Value = Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_source_column).Value
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_type_column).Value = Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_type_column).Value
        'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_section_column).Value = Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_section_column).Value
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_title_column).Value = Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_title_column).Value
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_req_column).Value = Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_req_column).Value
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_rationale_column).Value = Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_rationale_column).Value
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_maturity_column).Value = Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_maturity_column).Value
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_comments_column).Value = Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_comments_column).Value
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_acceptance_column).Value = Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_acceptance_column).Value
        'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satarg_column).Value = Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_satarg_column).Value
        'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satisfied_column).Value = Workbooks(previous_wb).Sheets(previous_sheet).Cells(previous_current_row, previous_satisfied_column).Value
        
    
        'format cells as the deleted colour
        'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satisfies_column).Interior.Color = deleted_colour
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_id_column).Interior.Color = deleted_colour
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_source_column).Interior.Color = deleted_colour
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_type_column).Interior.Color = deleted_colour
        'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_section_column).Interior.Color = deleted_colour
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_title_column).Interior.Color = deleted_colour
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_req_column).Interior.Color = deleted_colour
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_rationale_column).Interior.Color = deleted_colour
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_maturity_column).Interior.Color = deleted_colour
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_comments_column).Interior.Color = deleted_colour
        Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_acceptance_column).Interior.Color = deleted_colour
        'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satarg_column).Interior.Color = deleted_colour
        'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_satisfied_column).Interior.Color = deleted_colour
    
    Else
    End If
        
Next previous_current_row


'**********************************************************************************************
'End Timer
'**********************************************************************************************
'Determine how many seconds code took to run
SecondsElapsed = Round(Timer - StartTime, 2)


'**********************************************************************************************
'End Message
'**********************************************************************************************
Workbooks(this_wb).Activate
Sheets(results_sheet).Select
MsgBox "Baseline compare completed in " & SecondsElapsed & " seconds."

GoTo Reset:


'**********************************************************************************************
'Error Messages
'**********************************************************************************************
Err1:
MsgBox "The 'current' workbook '" & current_wb & "' is not open. This macro will now exit."
GoTo Reset:

Err2:
MsgBox "The 'previous' workbook '" & previous_wb & "' is not open. This macro will now exit."
GoTo Reset:


Err11:
MsgBox "Worksheet '" & current_sheet & "' not found in workbook '" & current_wb & "'. This macro will now exit."
GoTo Reset:

Err21:
MsgBox "Worksheet '" & previous_sheet & "' not found in workbook '" & previous_wb & "'. This macro will now exit."
GoTo Reset:

'**********************************************************************************************
'Reset Excel updating
'**********************************************************************************************
Reset:
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True


End Sub

Sub Search(ByRef search_term, ByRef search_row, ByRef search_column)
'Standard search function to find the location of certain terms within worksheets

'Test message box
'MsgBox "Search term: " & Search_Term

'reset the search row and column to 1 before starting the search
search_row = 1
search_column = 1

Dim search_end_row As Integer
search_end_row = 1000 'sets the number of rows the search will cycle through before giving up

Dim search_end_column As Integer
search_end_column = 1000 'sets the number of columns the search will look in before giving up

Do While search_row <= search_end_row
    Do While search_column <= search_end_column

        If Cells(search_row, search_column).Value = search_term Then

            'Test message box
            'MsgBox "'" & Search_Term & "' found in " & ActiveSheet.Name & " worksheet at row: " & search_row & ", column: " & search_column
            Exit Sub

        Else
        
        End If

        search_column = search_column + 1
        
    Loop 'column loop

    search_column = 1 'reset search column
    search_row = search_row + 1
    
Loop 'row loop

MsgBox "'" & search_term & "' not found in '" & ActiveSheet.Name & "' worksheet. This macro will now exit."

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True

End

End Sub
