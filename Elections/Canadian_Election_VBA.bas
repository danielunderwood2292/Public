Attribute VB_Name = "UK_Election_Main"
Sub Add_New_Poll()
Attribute Add_New_Poll.VB_ProcData.VB_Invoke_Func = "A\n14"

'**************************************************************************************************
'Start Timing**************************************************************************************
'**************************************************************************************************
'Set up a timer to record how long the macro takes to run
Dim StartTime As Double
Dim SecondsElapsed As Double

'Remember time when macro starts
StartTime = Timer

'**************************************************************************************************
'Switch off functionality
'**************************************************************************************************
Call Switch_Off_Functionality

'**************************************************************************************************
'Check number of open workbooks is 2
'**************************************************************************************************
'Variable for this workbook name
Dim this_wb As String
this_wb = ThisWorkbook.Name

'Polls workbook name
Dim poll_wb As String

'Number of open workbooks variable
Dim no_open_wbs As Integer
no_open_wbs = 0

'Variable for open workbook details
Dim wb As Workbook

'Loop around each open workbook
For Each wb In Workbooks

    If wb.FullName = ThisWorkbook.FullName Then
    Else
        poll_wb = wb.Name
    End If
      
    no_open_workbooks = no_open_workbooks + 1
    
Next wb

'Check that only 2 workbooks are open - this and another
If no_open_workbooks > 2 Then
    
    MsgBox "There is more than one other workbook open. Cannot determine which contains the Polls data. In order for this macro to work correctly, only one other workbook may be open. This macro will now exit."
    GoTo Macro_Exit
      
ElseIf no_open_workbooks = 1 Then

    MsgBox "No poll workbook is open. This macro will now exit."
    GoTo Macro_Exit
    
Else
    
End If

'**************************************************************************************************
'Setup variables
'**************************************************************************************************
'Poll inut sheet variables
Workbooks(poll_wb).Activate

Dim total_no_poll_sheets As Integer
total_no_poll_sheets = Worksheets.Count

If total_no_poll_sheets > 1 Then
    MsgBox "The poll input workbook has more than one worksheet. Cannot determine which contains the poll data. This macro will now exit."
    GoTo Macro_Exit
Else

End If

Dim sPoll As String
sPoll = ActiveSheet.Name

'Sheet names
Dim sOverall, sConst, sCand, sMap, sRegions, sShare, sPollsters, sPolls, sParties As String
Call VA_Worksheet_Names(sOverall, sConst, sCand, sMap, sRegions, sShare, sPollsters, sPolls, sParties)

'Poll sheet variables
Workbooks(poll_wb).Activate
Sheets(sPoll).Select
Dim poll_metadata_column, poll_metadata_row, poll_pollster_row, poll_date_row, poll_type_row, poll_scope_row, poll_party_column, poll_party_row, poll_region_row, poll_region_column, poll_swing_row, poll_swing_column, poll_final_row As Integer
Call VB_Poll_Variables(poll_metadata_column, poll_metadata_row, poll_pollster_row, poll_date_row, poll_type_row, poll_scope_row, poll_party_column, poll_party_row, poll_region_row, poll_region_column, poll_swing_row, poll_swing_column, poll_final_row)

'Pollster sheet variables
Workbooks(this_wb).Activate
Sheets(sPollsters).Select
Dim pollsters_heading_row, pollsters_pollster_column, pollsters_type_column, pollsters_parties_column, pollsters_final_row As Integer
Call VC_Pollsters_Variables(pollsters_heading_row, pollsters_pollster_column, pollsters_type_column, pollsters_parties_column, pollsters_final_row)

'Regions sheet variables
Workbooks(this_wb).Activate
Sheets(sRegions).Select
Dim regions_heading_row, regions_type_column, regions_region_column, regions_votes_column, regions_constituencies_column, regions_final_row As Integer
Call VD_Regions_Variables(regions_heading_row, regions_type_column, regions_region_column, regions_votes_column, regions_constituencies_column, regions_final_row)

'Share sheet variables
Workbooks(this_wb).Activate
Sheets(sShare).Select
Dim share_heading_row, share_final_row, share_type_column, share_region_column, share_party_column, share_adjusted_column As Integer
Call VE_Share_Variables(share_heading_row, share_final_row, share_type_column, share_region_column, share_party_column, share_adjusted_column)

'Candidates variables
Workbooks(this_wb).Activate
Sheets(sCand).Select
Dim cand_heading_row, cand_final_row, cand_const_column, cand_party_column, cand_2019share_column, cand_final_column, cand_totalvotes_column, cand_2019winner_column, cand_predvotes_column, cand_predshare_column, cand_polls_column, cand_dates_column, cand_standing_column As Integer
Call VF_Cand_Variables(cand_heading_row, cand_final_row, cand_const_column, cand_party_column, cand_2019share_column, cand_final_column, cand_totalvotes_column, cand_2019winner_column, cand_predvotes_column, cand_predshare_column, cand_polls_column, cand_dates_column, cand_standing_column)

'Polls sheet variables
Workbooks(this_wb).Activate
Sheets(sPolls).Select
Dim polls_heading_row, polls_final_row, polls_pollster_column, polls_date_column, polls_type_column, polls_scope_column, polls_file_column, polls_candcolumn_column, polls_applicable_column, polls_final_column, polls_const_column, polls_region_column, polls_nation_column As Integer
Call VG_Polls_Variables(polls_heading_row, polls_final_row, polls_pollster_column, polls_date_column, polls_type_column, polls_scope_column, polls_file_column, polls_candcolumn_column, polls_applicable_column, polls_final_column, polls_const_column, polls_region_column, polls_nation_column)

'**************************************************************************************************
'Check poll sheet has all of the fields completed
'**************************************************************************************************
Workbooks(poll_wb).Activate
Sheets(sPoll).Select

If IsEmpty(Cells(poll_pollster_row, poll_metadata_column).Value) = True Or IsEmpty(Cells(poll_date_row, poll_metadata_column).Value) = True Or IsEmpty(Cells(poll_type_row, poll_metadata_column).Value) = True Or IsEmpty(Cells(poll_scope_row, poll_metadata_column).Value) = True Then
    MsgBox "The poll metadata is incomplete. This macro will now exit."
    GoTo Macro_Exit
Else
End If

'**************************************************************************************************
'Check poll sheet metadata is in the expected format
'**************************************************************************************************
Dim pollster, poll_type, poll_scope As String
Dim poll_date As Date

pollster = Workbooks(poll_wb).Sheets(sPoll).Cells(poll_pollster_row, poll_metadata_column).Value
On Error GoTo Date_Error
poll_date = Workbooks(poll_wb).Sheets(sPoll).Cells(poll_date_row, poll_metadata_column).Value
On Error GoTo 0
poll_type = Workbooks(poll_wb).Sheets(sPoll).Cells(poll_type_row, poll_metadata_column).Value
poll_scope = Workbooks(poll_wb).Sheets(sPoll).Cells(poll_scope_row, poll_metadata_column).Value

Dim pollsters_current_row As Integer
Dim pollster_found, type_found As Boolean
pollster_found = False
type_found = False

For pollsters_current_row = pollsters_heading_row To pollsters_final_row

    If Workbooks(this_wb).Sheets(sPollsters).Cells(pollsters_current_row, pollsters_pollster_column).Value = pollster Then
        pollster_found = True
        Exit For
    Else
    End If

Next pollsters_current_row

If pollster_found = False Then
    MsgBox "This pollster has not been found in the database. This macro will now exit."
    GoTo Macro_Exit
Else
End If

For pollsters_current_row = pollsters_heading_row To pollsters_final_row

    If Workbooks(this_wb).Sheets(sPollsters).Cells(pollsters_current_row, pollsters_type_column).Value = poll_type Then
        type_found = True
        Exit For
    Else
    End If

Next pollsters_current_row

If type_found = False Then
    MsgBox "The poll type has not been found in the database. This macro will now exit."
    GoTo Macro_Exit
Else
End If

'Find the region column in the candidate worksheet
Dim cand_region_column, cand_current_column, cand_current_row As Integer
Dim region_found As Boolean
region_found = False

If poll_type = "constituency" Then

    cand_region_column = cand_const_column
    region_found = True
Else

    For cand_current_column = 1 To cand_final_column
        If Workbooks(this_wb).Sheets(sCand).Cells(cand_heading_row, cand_current_column).Value = poll_type Then
            cand_region_column = cand_current_column
            region_found = True
            Exit For
        Else
        End If
    Next cand_current_column
    
    If region_found = False Then
        MsgBox "The regions applicable to this poll have not been found in the VP worksheet. This macro will now exit."
        GoTo Macro_Exit
    Else
    End If
    
End If

'**************************************************************************************************
'Check poll sheet has all of the regions expected for the poll type
'**************************************************************************************************
'Determine number of regions for the poll type
Dim no_regions, regions_current_row, poll_current_column, poll_current_row As Integer
poll_current_column = poll_region_column

If poll_type = "constituency" Or poll_scope <> "all" Then

    no_regions = 1

Else

    For regions_current_row = regions_heading_row To regions_final_row
    
        If Workbooks(this_wb).Sheets(sRegions).Cells(regions_current_row, regions_type_column).Value = poll_type Then
            Do While Workbooks(this_wb).Sheets(sRegions).Cells(regions_current_row, regions_type_column).Value = poll_type
            
                If Workbooks(this_wb).Sheets(sRegions).Cells(regions_current_row, regions_region_column).Value = Workbooks(poll_wb).Sheets(sPoll).Cells(poll_region_row, poll_current_column).Value Then
                
                Else
                    MsgBox Workbooks(this_wb).Sheets(sRegions).Cells(regions_current_row, regions_region_column).Value & " region not found in poll workbook. This macro will not exit."
                    GoTo Macro_Exit
                End If
                
                no_regions = no_regions + 1
                poll_current_column = poll_current_column + 1
                regions_current_row = regions_current_row + 1
            Loop
            
            Exit For
        Else
        End If
    
    Next regions_current_row
    
End If

'**************************************************************************************************
'Check poll sheet has all of the correct regions in the swing section
'**************************************************************************************************
For poll_current_column = poll_region_column To poll_region_column + no_regions - 1
    If Workbooks(poll_wb).Sheets(sPoll).Cells(poll_region_row, poll_current_column).Value = Workbooks(poll_wb).Sheets(sPoll).Cells(poll_region_row, poll_current_column + no_regions).Value Then
    Else
        MsgBox "Region '" & Workbooks(poll_wb).Sheets(sPoll).Cells(poll_region_row, poll_current_column).Value & "' not found in corresponding 'Swing' column (" & poll_current_column + no_regions & "). This macro will now exit."
        GoTo Macro_Exit
    End If
Next poll_current_column

'**************************************************************************************************
'Check a poll with the same metadata has not been added to the model
'**************************************************************************************************
Dim polls_current_row As Integer
Dim current_pollster, current_poll_type, current_poll_scope As String
Dim current_poll_date As Date

For polls_current_row = polls_heading_row + 1 To polls_final_row

    current_pollster = Workbooks(this_wb).Sheets(sPolls).Cells(polls_current_row, polls_pollster_column).Value
    current_poll_type = Workbooks(this_wb).Sheets(sPolls).Cells(polls_current_row, polls_type_column).Value
    current_poll_scope = Workbooks(this_wb).Sheets(sPolls).Cells(polls_current_row, polls_scope_column).Value
    current_poll_date = Workbooks(this_wb).Sheets(sPolls).Cells(polls_current_row, polls_date_column).Value

    If current_pollster = pollster And current_poll_type = poll_type And current_poll_scope = poll_scope And current_poll_date = poll_date Then
        MsgBox "A poll with this same metadata has already been added to the model. This macro will now exit."
        GoTo Macro_Exit
    Else
    End If

Next polls_current_row

'**************************************************************************************************
'Determine percentage swings for each region - in poll spreadsheet
'**************************************************************************************************
'Find each matching type, region and party in the share worksheet
'Need to loop around each column then each row in the poll input worksheet
Dim current_party, current_region As String
Dim swing, poll_share, previous_share, share_2019, raw_share As Double
Dim share_current_row As Integer
Dim share_found As Boolean

For poll_current_column = poll_region_column To poll_region_column + no_regions - 1

    For poll_current_row = poll_region_row + 1 To poll_final_row
    
        current_party = Workbooks(poll_wb).Sheets(sPoll).Cells(poll_current_row, poll_party_column).Value
        If current_party = "" Then
            MsgBox "No party name found at cell (" & poll_current_row & "," & poll_party_column & ") in the poll worksheet. This macro will not exit."
            GoTo Macro_Exit
        Else
        End If
        
        current_region = Workbooks(poll_wb).Sheets(sPoll).Cells(poll_region_row, poll_current_column).Value
        If current_party = "" Then
            MsgBox "No party name found at cell (" & poll_current_row & "," & poll_party_column & ") in the poll worksheet. This macro will not exit."
            GoTo Macro_Exit
        Else
        End If
        
        poll_share = Workbooks(poll_wb).Sheets(sPoll).Cells(poll_current_row, poll_current_column).Value
        If poll_share = "" Then
            MsgBox "No poll numbers added at cell (" & poll_current_row & "," & poll_current_column & ") in the poll worksheet. This macro will not exit."
            GoTo Macro_Exit
        Else
        End If
        
        share_found = False
        
        If poll_type = "constituency" Then
        
            For cand_current_row = cand_heading_row + 1 To cand_final_row
                If Workbooks(this_wb).Sheets(sCand).Cells(cand_current_row, cand_party_column).Value = current_party And Workbooks(this_wb).Sheets(sCand).Cells(cand_current_row, cand_const_column).Value = poll_scope Then
                    share_found = True
                    previous_share = Workbooks(this_wb).Sheets(sCand).Cells(cand_current_row, cand_2019share_column).Value
                    swing = poll_share - previous_share
                    Workbooks(poll_wb).Sheets(sPoll).Cells(poll_current_row, poll_current_column + no_regions).Value = swing
                    Exit For
                Else
                End If
            Next cand_current_row
        
        Else
        
            For share_current_row = share_heading_row To share_final_row
            
                If Workbooks(this_wb).Sheets(sShare).Cells(share_current_row, share_type_column).Value = poll_type And Workbooks(this_wb).Sheets(sShare).Cells(share_current_row, share_party_column).Value = current_party And Workbooks(this_wb).Sheets(sShare).Cells(share_current_row, share_region_column).Value = current_region Then
                    share_found = True
                    previous_share = Workbooks(this_wb).Sheets(sShare).Cells(share_current_row, share_adjusted_column).Value
                    swing = poll_share - previous_share
                    Workbooks(poll_wb).Sheets(sPoll).Cells(poll_current_row, poll_current_column + no_regions).Value = swing
                    Exit For
                Else
                End If
            
            Next share_current_row
            
        End If
        
        If share_found = False Then
            MsgBox current_party & ", " & current_region & ", " & poll_type & " not found in model. This macro will not exit."
            GoTo Macro_Exit
        Else
        End If
    
    Next poll_current_row

Next poll_current_column

'**************************************************************************************************
'Add new worksheet in poll workbook to contain the calculation results
'**************************************************************************************************
Dim sVotes As String
sVotes = "Votes"
Sheets.Add(After:=Sheets(sPoll)).Name = sVotes

'Votes sheet variables
Dim votes_heading_row, votes_const_column, votes_party_column, votes_region_column, votes_rawshare_column, votes_2019share_column, votes_totalvotes_column, votes_adjshare_column, votes_votes_column As Integer
votes_heading_row = 1
votes_const_column = 1
votes_party_column = 2
votes_region_column = 3
votes_2019share_column = 4
votes_totalvotes_column = 5
votes_standing_column = 6
votes_rawshare_column = 7
votes_adjshare_column = 8
votes_votes_column = 9

'Add headings to votes sheet
Workbooks(poll_wb).Sheets(sVotes).Cells(votes_heading_row, votes_const_column).Value = "constituency"
Workbooks(poll_wb).Sheets(sVotes).Cells(votes_heading_row, votes_party_column).Value = "party"
Workbooks(poll_wb).Sheets(sVotes).Cells(votes_heading_row, votes_region_column).Value = "region"
Workbooks(poll_wb).Sheets(sVotes).Cells(votes_heading_row, votes_2019share_column).Value = "share_2019"
Workbooks(poll_wb).Sheets(sVotes).Cells(votes_heading_row, votes_totalvotes_column).Value = "total_votes"
Workbooks(poll_wb).Sheets(sVotes).Cells(votes_heading_row, votes_standing_column).Value = "standing"
Workbooks(poll_wb).Sheets(sVotes).Cells(votes_heading_row, votes_rawshare_column).Value = "raw_share"
Workbooks(poll_wb).Sheets(sVotes).Cells(votes_heading_row, votes_adjshare_column).Value = "adjusted_share"
Workbooks(poll_wb).Sheets(sVotes).Cells(votes_heading_row, votes_votes_column).Value = "votes"

'**************************************************************************************************
'Populate the votes sheet with the applicable candidates and 2019 vote share including when 2019 share = 0
'**************************************************************************************************
Dim votes_current_row, scope_column As Integer
Dim current_const, standing As String

votes_current_row = votes_heading_row + 1

'check applicability
'If poll_type = "constituency" Then
    
        'scope_column = cand_const_column
        
'Else

    For cand_current_row = cand_heading_row + 1 To cand_final_row

        current_const = Workbooks(this_wb).Sheets(sCand).Cells(cand_current_row, cand_const_column).Value
        current_party = Workbooks(this_wb).Sheets(sCand).Cells(cand_current_row, cand_party_column).Value
        current_region = Workbooks(this_wb).Sheets(sCand).Cells(cand_current_row, cand_region_column).Value
        share_2019 = Workbooks(this_wb).Sheets(sCand).Cells(cand_current_row, cand_2019share_column).Value
        standing = Workbooks(this_wb).Sheets(sCand).Cells(cand_current_row, cand_standing_column).Value
        
        If Workbooks(this_wb).Sheets(sCand).Cells(cand_current_row, cand_region_column).Value = "N/A" Then
        
        ElseIf poll_scope <> "all" And Workbooks(this_wb).Sheets(sCand).Cells(cand_current_row, cand_region_column).Value <> poll_scope Then
        
        Else
            
            Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_const_column).Value = current_const
            Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_party_column).Value = current_party
            Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_region_column).Value = current_region
            Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_standing_column).Value = standing
            Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_totalvotes_column).Value = Workbooks(this_wb).Sheets(sCand).Cells(cand_current_row, cand_totalvotes_column).Value
                                    
            'If 2019 share = 0 then vote_share = adjusted_vote_share_2024 + swing
            If share_2019 > 0 Or poll_type = "constituency" Then
            
                Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_2019share_column).Value = share_2019
          
            Else
                'find adjusted share in share seats
                For share_current_row = share_heading_row + 1 To share_final_row
                    If Workbooks(this_wb).Sheets(sShare).Cells(share_current_row, share_type_column) = poll_type And Workbooks(this_wb).Sheets(sShare).Cells(share_current_row, share_party_column).Value = current_party And Workbooks(this_wb).Sheets(sShare).Cells(share_current_row, share_region_column).Value = current_region Then
                    
                        Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_2019share_column).Value = Workbooks(this_wb).Sheets(sShare).Cells(share_current_row, share_adjusted_column).Value
                        Exit For
                    
                    Else
                    End If
                Next share_current_row
                
            End If
           
            votes_current_row = votes_current_row + 1
        
        End If

    Next cand_current_row

'End If

'**************************************************************************************************
'Add swing to each candidate in each applicable constituency
'**************************************************************************************************
Dim votes_final_row As Integer
Dim is_standing As String

votes_final_row = WorksheetFunction.CountA(Columns(votes_const_column))

For votes_current_row = votes_heading_row + 1 To votes_final_row

        current_const = Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_const_column).Value
        current_party = Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_party_column).Value
        current_region = Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_region_column).Value
        share_2019 = Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_2019share_column).Value
        is_standing = Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_standing_column).Value
        
        'Find the correct swing to apply from the poll sheet
        For poll_current_row = poll_region_row + 1 To poll_final_row
        
            If Workbooks(poll_wb).Sheets(sPoll).Cells(poll_current_row, poll_party_column).Value = current_party Then
                For poll_current_column = poll_swing_column To poll_swing_column + no_regions - 1
                    If Workbooks(poll_wb).Sheets(sPoll).Cells(poll_region_row, poll_current_column).Value = current_region Then
                        swing = Workbooks(poll_wb).Sheets(sPoll).Cells(poll_current_row, poll_current_column).Value
                        raw_share = swing + share_2019
                        
                        If raw_share < 0 Or is_standing <> "Y" Then
                            Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_rawshare_column).Value = 0
                        Else
                            Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_rawshare_column).Value = raw_share
                        End If
                        
                        Exit For
                    Else
                    End If
                Next poll_current_column
            Else
            End If
            
        Next poll_current_row

Next votes_current_row

'**************************************************************************************************
'Adjust modified vote shares to 100% and Determine number of votes for each candidate
'**************************************************************************************************
Dim total_share, share_raw As Double
Dim new_votes As Long

For votes_current_row = votes_heading_row + 1 To votes_final_row

    current_const = Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_const_column).Value
    share_raw = Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_rawshare_column).Value
    
    Workbooks(poll_wb).Activate
    Sheets(sVotes).Select
    total_share = WorksheetFunction.SumIfs(Columns(votes_rawshare_column), Columns(votes_const_column), current_const)
    Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_adjshare_column).Value = share_raw / total_share
    
    new_votes = Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_totalvotes_column).Value * Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_adjshare_column).Value
    Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_votes_column).Value = new_votes
   
Next votes_current_row

'**************************************************************************************************
'Add poll metadata to overall spreadsheet
'**************************************************************************************************
poll_current_row = polls_final_row + 1

If poll_type = "constituency" Then

    Workbooks(this_wb).Sheets(sPolls).Cells(polls_current_row, polls_const_column).Value = "Y"
    Workbooks(this_wb).Sheets(sPolls).Cells(polls_current_row, polls_region_column).Value = "N"
    Workbooks(this_wb).Sheets(sPolls).Cells(polls_current_row, polls_nation_column).Value = "N"

ElseIf poll_type = "gta" Or poll_type = "gta_mainstreet" Or poll_type = "gta_all" Or poll_type = "insights_bc" Or poll_type = "campaign_ontario" Or poll_type = "leger_manitoba" Or poll_scope <> "all" Then

    Workbooks(this_wb).Sheets(sPolls).Cells(polls_current_row, polls_const_column).Value = "N"
    Workbooks(this_wb).Sheets(sPolls).Cells(polls_current_row, polls_region_column).Value = "Y"
    Workbooks(this_wb).Sheets(sPolls).Cells(polls_current_row, polls_nation_column).Value = "N"
    
ElseIf poll_type = "Canada" Or poll_type = "region" Or poll_type = "province" Or poll_type = "ekos" Or poll_type = "counsel" Then

    Workbooks(this_wb).Sheets(sPolls).Cells(polls_current_row, polls_const_column).Value = "N"
    Workbooks(this_wb).Sheets(sPolls).Cells(polls_current_row, polls_region_column).Value = "N"
    Workbooks(this_wb).Sheets(sPolls).Cells(polls_current_row, polls_nation_column).Value = "Y"

Else
    MsgBox "Poll type " & poll_type & " cannot be computed by the model and cannot be added to the polls worksheet. This macro will not exit."
    GoTo Macro_Exit
End If

Workbooks(this_wb).Sheets(sPolls).Cells(polls_current_row, polls_pollster_column).Value = pollster
Workbooks(this_wb).Sheets(sPolls).Cells(polls_current_row, polls_type_column).Value = poll_type
Workbooks(this_wb).Sheets(sPolls).Cells(polls_current_row, polls_scope_column).Value = poll_scope
Workbooks(this_wb).Sheets(sPolls).Cells(polls_current_row, polls_date_column).Value = poll_date
Workbooks(this_wb).Sheets(sPolls).Cells(polls_current_row, polls_file_column).Value = poll_wb

'**************************************************************************************************
'Add votes to overall spreadsheet
'**************************************************************************************************
Dim cand_poll_column As Integer
cand_poll_column = cand_polls_column + 1

Dim cand_const, cand_party As String

Workbooks(this_wb).Activate
Sheets(sCand).Select
Columns(cand_poll_column).Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow

Workbooks(this_wb).Sheets(sCand).Cells(cand_heading_row, cand_poll_column).Value = poll_wb
votes_current_row = votes_heading_row + 1

For cand_current_row = cand_heading_row + 1 To cand_final_row

        current_const = Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_const_column).Value
        current_party = Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_party_column).Value
        new_votes = Workbooks(poll_wb).Sheets(sVotes).Cells(votes_current_row, votes_votes_column).Value
        
        cand_const = Workbooks(this_wb).Sheets(sCand).Cells(cand_current_row, cand_const_column).Value
        cand_party = Workbooks(this_wb).Sheets(sCand).Cells(cand_current_row, cand_party_column).Value
        
        If cand_const = current_const And cand_party = current_party Then
                Workbooks(this_wb).Sheets(sCand).Cells(cand_current_row, cand_poll_column).Value = new_votes
                votes_current_row = votes_current_row + 1
        Else
        End If
        
Next cand_current_row

'**************************************************************************************************
'Switch on functionality
'**************************************************************************************************
Call Switch_On_Functionality

'**************************************************************************************************
'Results Summary Message***************************************************************************
'**************************************************************************************************
'Determine how many seconds code took to run
 SecondsElapsed = Round(Timer - StartTime, 2)

'String to contain the message contents for the results message
Dim results_summary As String
results_summary = "Poll incorporated in: " & SecondsElapsed & " seconds"
MsgBox results_summary
Workbooks(poll_wb).Activate
End

'**************************************************************************************************
'Error Messages
'**************************************************************************************************
Date_Error:
MsgBox "Poll date is not in a date format. This macro will now exit."
GoTo Macro_Exit

'**************************************************************************************************
'Exit Macro
'**************************************************************************************************
Macro_Exit:
Call Switch_On_Functionality
Workbooks(poll_wb).Activate

End Sub

Sub Determine_Date_Votes()
Attribute Determine_Date_Votes.VB_ProcData.VB_Invoke_Func = "D\n14"
'**************************************************************************************************
'Start Timing
'**************************************************************************************************
'Set up a timer to record how long the macro takes to run
Dim StartTime As Double
Dim SecondsElapsed As Double

'Remember time when macro starts
StartTime = Timer

'**************************************************************************************************
'Switch Off Functionality
'**************************************************************************************************
Call Switch_Off_Functionality

'**************************************************************************************************
'Setup Variables
'**************************************************************************************************
Dim election_called As Boolean
election_called = True
Dim national_poll_limit, regional_poll_limit As Integer

If election_called = True Then
    national_poll_limit = 7
    regional_poll_limit = 30
Else
    national_poll_limit = 30
    regional_poll_limit = 90
End If

'Sheet names
Dim sOverall, sConst, sCand, sMap, sRegions, sShare, sPollsters, sPolls, sParties As String
Call VA_Worksheet_Names(sOverall, sConst, sCand, sMap, sRegions, sShare, sPollsters, sPolls, sParties)

'Candidates variables
Sheets(sCand).Select
Dim cand_heading_row, cand_final_row, cand_const_column, cand_party_column, cand_2019share_column, cand_final_column, cand_totalvotes_column, cand_2019winner_column, cand_predvotes_column, cand_predshare_column, cand_polls_column, cand_dates_column, cand_standing_column As Integer
Call VF_Cand_Variables(cand_heading_row, cand_final_row, cand_const_column, cand_party_column, cand_2019share_column, cand_final_column, cand_totalvotes_column, cand_2019winner_column, cand_predvotes_column, cand_predshare_column, cand_polls_column, cand_dates_column, cand_standing_column)
'Polls sheet variables
Sheets(sPolls).Select
Dim polls_heading_row, polls_final_row, polls_pollster_column, polls_date_column, polls_type_column, polls_scope_column, polls_file_column, polls_candcolumn_column, polls_applicable_column, polls_final_column, polls_const_column, polls_region_column, polls_nation_column As Integer
Call VG_Polls_Variables(polls_heading_row, polls_final_row, polls_pollster_column, polls_date_column, polls_type_column, polls_scope_column, polls_file_column, polls_candcolumn_column, polls_applicable_column, polls_final_column, polls_const_column, polls_region_column, polls_nation_column)

'Constituency sheet variables
Sheets(sConst).Select
Dim const_heading_row, const_final_row, const_final_column, const_region_column, const_const_column, const_2019votes_column, const_2019winner_column, const_2019maj_column, const_2019majpc_column, const_predwinner_column, const_predmaj_column, const_predmajpc_column, const_gain_column, const_loss_column As Integer
Call VH_Const_Variables(const_heading_row, const_final_row, const_final_column, const_region_column, const_const_column, const_2019votes_column, const_2019winner_column, const_2019maj_column, const_2019majpc_column, const_predwinner_column, const_predmaj_column, const_predmajpc_column, const_gain_column, const_loss_column)

'Overall Sheet Variables
Sheets(sOverall).Select
Dim overall_heading_row, overall_heading_column, overall_final_row, overall_no_parties, overall_votes_row, overall_share_row, overall_seatdelta_row As Integer
Call VI_Overall_Variables(overall_heading_row, overall_heading_column, overall_final_row, overall_no_parties, overall_votes_row, overall_share_row, overall_seatdelta_row)

'**************************************************************************************************
'Get assess date from user
'**************************************************************************************************
Dim assess_date, min_date, cand_date As Date
Dim date_string As String
Dim is_date As Boolean
is_date = False

Do While is_date = False
    date_string = InputBox("Enter date to be assessed:")
    If IsDate(date_string) Then
        assess_date = DateValue(date_string)
        is_date = True
    Else
        MsgBox "Invalid date, please try again."
    End If
Loop

Dim date_delta, cand_this_column As Integer
cand_this_column = cand_dates_column + 1

min_date = "01/08/2021"
cand_date = Sheets("Cand").Cells(cand_heading_row, cand_this_column).Value

If assess_date <= cand_date Then
    MsgBox "The entered date has already been assessed in the model. This macro will now exit."
    GoTo Macro_Exit
Else
End If

'**************************************************************************************************
'Add assess date column in the cand worksheet
'**************************************************************************************************
Sheets(sCand).Select
Columns(cand_this_column).Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
Cells(cand_heading_row, cand_this_column).Value = assess_date

'**************************************************************************************************
'Add assess date column in the const worksheet
'**************************************************************************************************
Dim const_date_column, const_current_column As Integer
const_date_column = const_loss_column + 1

Sheets(sConst).Select
Columns(const_date_column).Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
Cells(const_heading_row, const_date_column).Value = assess_date

'**************************************************************************************************
'Add assess date column in the overall worksheet
'**************************************************************************************************
Sheets(sOverall).Select
Dim overall_results_column As Integer
overall_results_column = WorksheetFunction.CountA(Rows(overall_heading_row)) + 1

Columns(overall_results_column).Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Cells(overall_heading_row, overall_results_column).Value = assess_date

'**************************************************************************************************
'Sort polls sheet
'**************************************************************************************************
Sheets(sPolls).Select
ActiveWorkbook.Worksheets(sPolls).sort.SortFields.Clear
With ActiveSheet.sort
     .SortFields.Add Key:=Range(Cells(polls_heading_row, polls_date_column).Address(RowAbsolute:=False, ColumnAbsolute:=False)), Order:=xlDescending
     .SetRange Range(Cells(polls_heading_row, 1), Cells(polls_final_row, polls_final_column))
     .Header = xlYes
     .Apply
End With

'**************************************************************************************************
'Determine applicability of each poll
'**************************************************************************************************
Dim polls_current_row, cand_current_column As Integer
Dim current_poll_date As Date
Dim current_poll_type, current_poll_company, current_poll_scope, current_poll_file, current_poll_nation, current_poll_region, current_poll_constituency As String
Dim cand_poll_found As Boolean

For polls_current_row = polls_heading_row + 1 To polls_final_row

    current_poll_type = Sheets(sPolls).Cells(polls_current_row, polls_type_column).Value
    current_poll_company = Sheets(sPolls).Cells(polls_current_row, polls_pollster_column).Value
    current_poll_date = Sheets(sPolls).Cells(polls_current_row, polls_date_column).Value
    current_poll_scope = Sheets(sPolls).Cells(polls_current_row, polls_scope_column).Value
    current_poll_file = Sheets(sPolls).Cells(polls_current_row, polls_file_column).Value
    current_poll_nation = Sheets(sPolls).Cells(polls_current_row, polls_nation_column).Value
    current_poll_region = Sheets(sPolls).Cells(polls_current_row, polls_region_column).Value
    current_poll_constituency = Sheets(sPolls).Cells(polls_current_row, polls_const_column).Value
    date_delta = assess_date - current_poll_date

    'Find column position of the poll in Cand worksheet
    cand_poll_found = False
    For cand_current_column = cand_polls_column + 1 To cand_final_column
        If Sheets(sCand).Cells(cand_heading_row, cand_current_column).Value = current_poll_file Then
            cand_poll_found = True
            Sheets(sPolls).Cells(polls_current_row, polls_candcolumn_column).Value = cand_current_column
            Exit For
        Else
        End If
    Next cand_current_column
    
    If cand_poll_found = False Then
        MsgBox "'" & current_poll_file & "' not found in the 'Cand' worksheet. This macro will now exit."
        GoTo Macro_Exit
    Else
    End If
    
    If date_delta < 0 Then
        Sheets(sPolls).Cells(polls_current_row, polls_applicable_column).Value = "N"
    ElseIf date_delta > regional_poll_limit Then
        Sheets(sPolls).Cells(polls_current_row, polls_applicable_column).Value = "N"
    ElseIf date_delta > national_poll_limit Then
        If current_poll_constituency = "Y" Or current_poll_region = "Y" Then
            Sheets(sPolls).Cells(polls_current_row, polls_applicable_column).Value = "Y"
        Else
            Sheets(sPolls).Cells(polls_current_row, polls_applicable_column).Value = "N"
        End If
    Else
        Sheets(sPolls).Cells(polls_current_row, polls_applicable_column).Value = "Y"
    End If

Next polls_current_row

'**************************************************************************************************
'Recheck applicable polls to determine if any are the same poll company
'**************************************************************************************************
Dim polls_recheck_row As Integer
Dim recheck_poll_type, recheck_poll_company, recheck_poll_scope As String

For polls_current_row = polls_heading_row + 1 To polls_final_row

    current_poll_type = Sheets(sPolls).Cells(polls_current_row, polls_type_column).Value
    current_poll_company = Sheets(sPolls).Cells(polls_current_row, polls_pollster_column).Value
    current_poll_scope = Sheets(sPolls).Cells(polls_current_row, polls_scope_column).Value
    
    For polls_recheck_row = polls_current_row + 1 To polls_final_row
        If Sheets(sPolls).Cells(polls_current_row, polls_applicable_column).Value = "Y" Then
            recheck_poll_type = Sheets(sPolls).Cells(polls_recheck_row, polls_type_column).Value
            recheck_poll_company = Sheets(sPolls).Cells(polls_recheck_row, polls_pollster_column).Value
            recheck_poll_scope = Sheets(sPolls).Cells(polls_recheck_row, polls_scope_column).Value
            
            If recheck_poll_type = current_poll_type And recheck_poll_company = current_poll_company And recheck_poll_scope = current_poll_scope Then
                Sheets(sPolls).Cells(polls_recheck_row, polls_applicable_column).Value = "N"
            Else
            End If
        Else
        End If
    Next polls_recheck_row
Next polls_current_row

'**************************************************************************************************
'Sort polls sheet so only applicables are at the top
'**************************************************************************************************
Sheets(sPolls).Select
ActiveWorkbook.Worksheets(sPolls).sort.SortFields.Clear
With ActiveSheet.sort
     .SortFields.Add Key:=Range(Cells(polls_heading_row, polls_applicable_column).Address(RowAbsolute:=False, ColumnAbsolute:=False)), Order:=xlDescending
     .SortFields.Add Key:=Range(Cells(polls_heading_row, polls_nation_column).Address(RowAbsolute:=False, ColumnAbsolute:=False)), Order:=xlDescending
     .SortFields.Add Key:=Range(Cells(polls_heading_row, polls_region_column).Address(RowAbsolute:=False, ColumnAbsolute:=False)), Order:=xlDescending
     .SortFields.Add Key:=Range(Cells(polls_heading_row, polls_const_column).Address(RowAbsolute:=False, ColumnAbsolute:=False)), Order:=xlDescending
     .SortFields.Add Key:=Range(Cells(polls_heading_row, polls_date_column).Address(RowAbsolute:=False, ColumnAbsolute:=False)), Order:=xlDescending
     .SetRange Range(Cells(polls_heading_row, 1), Cells(polls_final_row, polls_final_column))
     .Header = xlYes
     .Apply
End With

'**************************************************************************************************
'Determine number of applicable polls of each type
'**************************************************************************************************
Sheets(sPolls).Select
Dim no_nation_polls, no_region_polls, no_const_polls As Integer
no_nation_polls = WorksheetFunction.CountIfs(Columns(polls_applicable_column), "Y", Columns(polls_nation_column), "Y")
no_region_polls = WorksheetFunction.CountIfs(Columns(polls_applicable_column), "Y", Columns(polls_region_column), "Y")
no_const_polls = WorksheetFunction.CountIfs(Columns(polls_applicable_column), "Y", Columns(polls_const_column), "Y")

'**************************************************************************************************
'Calculate average votes for each candidate in each constituency
'**************************************************************************************************
Dim average_votes, nation_votes, region_votes, const_votes As Double
Dim final_votes As Long
Dim applicable_const_polls, applicable_region_polls, applicable_nation_polls As Integer

For cand_current_row = cand_heading_row + 1 To cand_final_row
    average_votes = 0
    nation_votes = 0
    region_votes = 0
    const_votes = 0
    final_votes = 0
    applicable_const_polls = 0
    applicable_region_polls = 0
    applicable_nation_polls = 0
    
    For polls_current_row = polls_heading_row + 1 To polls_heading_row + no_nation_polls
        cand_current_column = Sheets(sPolls).Cells(polls_current_row, polls_candcolumn_column).Value
        
        If Sheets(sCand).Cells(cand_current_row, cand_current_column).Value <> "" Then
            applicable_nation_polls = applicable_nation_polls + 1
            nation_votes = nation_votes + Sheets(sCand).Cells(cand_current_row, cand_current_column).Value
        Else
        End If
    Next polls_current_row
        
    If applicable_nation_polls = 0 Then
        nation_votes = 0
    Else
        nation_votes = nation_votes / applicable_nation_polls
    End If
    
    For polls_current_row = polls_heading_row + no_nation_polls + 1 To polls_heading_row + no_nation_polls + no_region_polls
        cand_current_column = Sheets(sPolls).Cells(polls_current_row, polls_candcolumn_column).Value
        
        If Sheets(sCand).Cells(cand_current_row, cand_current_column).Value <> "" Then
            applicable_region_polls = applicable_region_polls + 1
            region_votes = region_votes + Sheets(sCand).Cells(cand_current_row, cand_current_column).Value
        Else
        End If
    
    Next polls_current_row
    
    If applicable_region_polls = 0 Then
        region_votes = 0
    Else
        region_votes = region_votes / applicable_region_polls
    End If
    
    For polls_current_row = polls_heading_row + no_nation_polls + no_region_polls + 1 To polls_heading_row + no_nation_polls + no_region_polls + no_const_polls
        cand_current_column = Sheets(sPolls).Cells(polls_current_row, polls_candcolumn_column).Value
        
        If Sheets(sCand).Cells(cand_current_row, cand_current_column).Value <> "" Then
            applicable_const_polls = applicable_const_polls + 1
            const_votes = const_votes + Sheets(sCand).Cells(cand_current_row, cand_current_column).Value
        Else
        End If

    Next polls_current_row
    
    If applicable_const_polls = 0 Then
        const_votes = 0
    Else
        const_votes = const_votes / applicable_const_polls
    End If
    
    'calculate overall votes
    If applicable_const_polls = 0 And applicable_region_polls = 0 And applicable_nation_polls = 0 Then
        average_votes = 0
    ElseIf applicable_nation_polls > 0 Then
        If applicable_region_polls = 0 Then
            If applicable_const_polls = 0 Then
                average_votes = nation_votes
            ElseIf applicable_const_polls > 0 Then
                average_votes = (nation_votes + const_votes) / 2
            Else
                GoTo average_votes_error
            End If
        ElseIf applicable_region_polls > 0 Then
            If applicable_const_polls = 0 Then
                average_votes = (nation_votes + region_votes) / 2
            ElseIf applicable_const_polls > 0 Then
                average_votes = (const_votes + ((nation_votes + region_votes) / 2)) / 2
            Else
                GoTo average_votes_error
            End If
        Else
            GoTo average_votes_error
        End If
    ElseIf applicable_nation_polls = 0 Then
        If applicable_region_polls = 0 Then
            average_votes = const_votes
        ElseIf applicable_region_polls > 0 Then
            If applicable_const_polls = 0 Then
                average_votes = region_votes
            ElseIf applicable_const_polls > 0 Then
                average_votes = (region_votes + const_votes) / 2
            Else
                GoTo average_votes_error
            End If
        Else
            GoTo average_votes_error
        End If
        
    Else
        GoTo average_votes_error
    End If
    
    final_votes = Round(average_votes, 0)
    Sheets(sCand).Cells(cand_current_row, cand_this_column).Value = final_votes
    Sheets(sCand).Cells(cand_current_row, cand_predvotes_column).Value = final_votes
    Sheets(sCand).Cells(cand_current_row, cand_predshare_column).Value = final_votes / Sheets(sCand).Cells(cand_current_row, cand_totalvotes_column).Value

Next cand_current_row

'**************************************************************************************************
'Determine seats winners
'**************************************************************************************************
Dim const_current_row As Integer
Dim const_current_const As String

'Consolidated votes variables
Dim console_lead_votes As Long
Dim console_second_votes As Long
Dim console_margin_votes As Long

Dim console_lead_share As Double
Dim console_second_share As Double
Dim console_margin_share As Double

cand_current_row = cand_heading_row + 1

For const_current_row = const_heading_row + 1 To const_final_row

    const_current_const = Sheets(sConst).Cells(const_current_row, const_const_column).Value
    
    console_lead_votes = 0
    console_second_votes = 0
    console_margin_votes = 0
    
    console_lead_share = 0
    console_second_share = 0
    console_margin_share = 0
    
    console_leader = ""

    Do While Sheets("Cand").Cells(cand_current_row, cand_const_column).Value = const_current_const
                     
        If Sheets("Cand").Cells(cand_current_row, cand_predvotes_column).Value > console_lead_votes Then
        
            console_second_votes = console_lead_votes
            console_second_share = console_lead_share
            console_leader = Sheets("Cand").Cells(cand_current_row, cand_party_column).Value
            console_lead_votes = Sheets("Cand").Cells(cand_current_row, cand_predvotes_column).Value
        
        ElseIf Sheets("Cand").Cells(cand_current_row, cand_predvotes_column).Value > console_second_votes Then
             
            console_second_votes = Sheets("Cand").Cells(cand_current_row, cand_predvotes_column).Value
        
        End If
        
        cand_current_row = cand_current_row + 1
    
    Loop
    
    console_margin_votes = console_lead_votes - console_second_votes
    console_margin_share = console_margin_votes / Sheets(sConst).Cells(const_current_row, const_2019votes_column).Value
    const_winner = console_leader
    
    Sheets(sConst).Cells(const_current_row, const_predwinner_column).Value = const_winner
    Sheets(sConst).Cells(const_current_row, const_date_column).Value = const_winner
    Sheets(sConst).Cells(const_current_row, const_predmaj_column).Value = console_margin_votes
    Sheets(sConst).Cells(const_current_row, const_predmajpc_column).Value = console_margin_share
    
Next const_current_row

'**************************************************************************************************
'GAIN/LOSS Algorithm*******************************************************************************
'**************************************************************************************************
Dim no_seats_changed As Integer
Dim const_gain, const_loss As String

For const_current_row = const_heading_row + 1 To const_final_row

    If Sheets(sConst).Cells(const_current_row, const_predwinner_column).Value = Sheets(sConst).Cells(const_current_row, const_2019winner_column).Value Then
    
        const_gain = Sheets(sConst).Cells(const_current_row, const_predwinner_column).Value & " HOLD"
        Sheets(sConst).Cells(const_current_row, const_gain_column).Value = const_gain
        Sheets(sConst).Cells(const_current_row, const_loss_column).Value = const_gain
    
    Else
    
        const_gain = Sheets(sConst).Cells(const_current_row, const_predwinner_column).Value & " GAIN"
        Sheets(sConst).Cells(const_current_row, const_gain_column).Value = const_gain
        
        const_loss = Sheets(sConst).Cells(const_current_row, const_2019winner_column).Value & " LOSS"
        Sheets(sConst).Cells(const_current_row, const_loss_column).Value = const_loss
        
        no_seats_changed = no_seats_changed + 1
    
    End If

Next const_current_row

'**************************************************************************************************
'Populate Overall Results
'**************************************************************************************************
Dim overall_current_party As String

Dim overall_party_share As Double

Dim overall_current_row, overall_total_seats, no_overall_party_seats, overall_seat_delta As Integer
overall_total_seats = 338

Dim overall_total_votes, overall_party_votes As Long
Sheets(sCand).Select
overall_total_votes = WorksheetFunction.Sum(Columns(cand_predvotes_column))

For overall_current_row = overall_heading_row + 1 To overall_heading_row + overall_no_parties

    overall_current_party = Sheets(sOverall).Cells(overall_current_row, overall_heading_column).Value
    
    'Calculate number of seats each party has
    Sheets(sConst).Select
    no_overall_party_seats = WorksheetFunction.CountIf(Columns(const_predwinner_column), overall_current_party)
    Sheets(sOverall).Cells(overall_current_row, overall_results_column).Value = no_overall_party_seats
    
    'Calculate number of votes
    Sheets(sCand).Select
    overall_party_votes = WorksheetFunction.SumIfs(Columns(cand_predvotes_column), Columns(cand_party_column), overall_current_party)
    Sheets(sOverall).Cells(overall_current_row + overall_no_parties + 2, overall_results_column).Value = overall_party_votes
    
    'Calculate % share
    overall_party_share = overall_party_votes / overall_total_votes
    Sheets(sOverall).Cells(overall_current_row + (2 * (overall_no_parties + 2)), overall_results_column).Value = overall_party_share
    
    'Calculate Seat deltas
    overall_seat_delta = no_overall_party_seats - Sheets(sOverall).Cells(overall_current_row, overall_heading_column + 1).Value
    Sheets(sOverall).Cells(overall_current_row + (3 * (overall_no_parties + 2)), overall_results_column).Value = overall_seat_delta
    
Next overall_current_row

'calculate totals
Dim seats_check, seatsdelta_check As Integer
Dim votes_check As Long
Dim share_check As Double

Sheets(sOverall).Select
seats_check = WorksheetFunction.Sum(Range(Cells(overall_heading_row + 1, overall_results_column), Cells(overall_heading_row + overall_no_parties, overall_results_column)))

If seats_check = overall_total_seats Then
    Sheets(sOverall).Cells(overall_heading_row + overall_no_parties + 1, overall_results_column).Value = overall_total_seats
Else
    MsgBox "The overall number of seats does not match the expected number."
End If

votes_check = WorksheetFunction.Sum(Range(Cells(overall_heading_row + overall_no_parties + 3, overall_results_column), Cells(overall_heading_row + (2 * overall_no_parties) + 2, overall_results_column)))

If votes_check = overall_total_votes Then
    Sheets(sOverall).Cells(overall_heading_row + (2 * overall_no_parties) + 3, overall_results_column).Value = overall_total_votes
Else
    MsgBox "The total overall number of votes does not match the expected number."
End If

share_check = WorksheetFunction.Sum(Range(Cells(overall_heading_row + (2 * overall_no_parties) + 5, overall_results_column), Cells(overall_heading_row + (3 * overall_no_parties) + 4, overall_results_column)))

If share_check < 1.001 Or share_check > 0.999 Then
    Sheets(sOverall).Cells(overall_heading_row + (3 * overall_no_parties) + 5, overall_results_column).Value = 1
Else
    MsgBox "The total overall share of votes does not add up to 100%."
End If

seats_check = WorksheetFunction.Sum(Range(Cells(overall_heading_row + (3 * overall_no_parties) + 7, overall_results_column), Cells(overall_heading_row + (4 * overall_no_parties) + 6, overall_results_column)))

If seats_check = 0 Then
    Sheets(sOverall).Cells(overall_heading_row + (4 * overall_no_parties) + 7, overall_results_column).Value = 0
Else
    MsgBox "The overall change in seats does not equal 0."
End If


'**************************************************************************************************
'Switch On Functionality
'**************************************************************************************************
Call Switch_On_Functionality

'**************************************************************************************************
'Results Summary Message***************************************************************************
'**************************************************************************************************
'Determine how many seconds code took to run
 SecondsElapsed = Round(Timer - StartTime, 2)

'String to contain the message contents for the results message
Dim results_summary As String
results_summary = "Votes evaluated for " & assess_date & " in: " & SecondsElapsed & " seconds"
MsgBox results_summary
End

'**************************************************************************************************
'Average votes error
'**************************************************************************************************
average_votes_error:
MsgBox "Cannot determine the average number of votes at row " & cand_current_row & ". This macro will now exit."
GoTo Macro_Exit

'**************************************************************************************************
'Exit Macro
'**************************************************************************************************
Macro_Exit:
Call Switch_On_Functionality
End

End Sub

Sub Map()
Attribute Map.VB_ProcData.VB_Invoke_Func = "M\n14"
'**************************************************************************************************
'Start Timing
'**************************************************************************************************
'Set up a timer to record how long the macro takes to run
Dim StartTime As Double
Dim SecondsElapsed As Double

'Remember time when macro starts
StartTime = Timer

'**************************************************************************************************
'Switch Off Functionality
'**************************************************************************************************
Call Switch_Off_Functionality

'**************************************************************************************************
'Setup Variables
'**************************************************************************************************
'Sheet names
Dim sOverall, sConst, sCand, sMap, sRegions, sShare, sPollsters, sPolls, sParties As String
Call VA_Worksheet_Names(sOverall, sConst, sCand, sMap, sRegions, sShare, sPollsters, sPolls, sParties)

'Candidates variables
Sheets(sCand).Select
Dim cand_heading_row, cand_final_row, cand_const_column, cand_party_column, cand_2019share_column, cand_final_column, cand_totalvotes_column, cand_2019winner_column, cand_predvotes_column, cand_predshare_column, cand_polls_column, cand_dates_column, cand_standing_column As Integer
Call VF_Cand_Variables(cand_heading_row, cand_final_row, cand_const_column, cand_party_column, cand_2019share_column, cand_final_column, cand_totalvotes_column, cand_2019winner_column, cand_predvotes_column, cand_predshare_column, cand_polls_column, cand_dates_column, cand_standing_column)

'Constituency sheet variables
Sheets(sConst).Select
Dim const_heading_row, const_final_row, const_final_column, const_region_column, const_const_column, const_2019votes_column, const_2019winner_column, const_2019maj_column, const_2019majpc_column, const_predwinner_column, const_predmaj_column, const_predmajpc_column, const_gain_column, const_loss_column As Integer
Call VH_Const_Variables(const_heading_row, const_final_row, const_final_column, const_region_column, const_const_column, const_2019votes_column, const_2019winner_column, const_2019maj_column, const_2019majpc_column, const_predwinner_column, const_predmaj_column, const_predmajpc_column, const_gain_column, const_loss_column)

'Overall Sheet Variables
Sheets(sOverall).Select
Dim overall_heading_row, overall_heading_column, overall_final_row, overall_no_parties, overall_votes_row, overall_share_row, overall_seatdelta_row As Integer
Call VI_Overall_Variables(overall_heading_row, overall_heading_column, overall_final_row, overall_no_parties, overall_votes_row, overall_share_row, overall_seatdelta_row)

'Map sheet variables
Sheets(sMap).Select
Dim map_start_row, map_final_row, map_start_column, map_final_column, map_party_column, map_voteshare_column, map_sharedelta_column, map_seats_column, map_seatsdelta_column, map_gain_column, map_results_start_row, map_hold_column, map_results_final_row, no_parties, extra_gap, delta_rows As Integer
Call VM_Map_Variables(map_start_row, map_final_row, map_start_column, map_final_column, map_party_column, map_voteshare_column, map_sharedelta_column, map_seats_column, map_seatsdelta_column, map_gain_column, map_results_start_row, map_hold_column, map_results_final_row, no_parties, extra_gap, delta_rows)

Dim map_current_row, map_current_column, const_current_row, overall_current_row As Integer
Dim const_current_const, const_winner As String

'Colours array
Dim colours_array(1 To 14, 1 To 2) As String
colours_array(1, 1) = "LPC HOLD"
colours_array(1, 2) = RGB(255, 124, 128)
colours_array(2, 1) = "LPC GAIN"
colours_array(2, 2) = RGB(255, 0, 0)
colours_array(3, 1) = "CPC HOLD"
colours_array(3, 2) = RGB(153, 204, 255)
colours_array(4, 1) = "CPC GAIN"
colours_array(4, 2) = RGB(0, 0, 255)
colours_array(5, 1) = "NDP HOLD"
colours_array(5, 2) = RGB(255, 204, 153)
colours_array(6, 1) = "NDP GAIN"
colours_array(6, 2) = RGB(255, 127, 0)
colours_array(7, 1) = "BQ HOLD"
colours_array(7, 2) = RGB(204, 255, 255)
colours_array(8, 1) = "BQ GAIN"
colours_array(8, 2) = RGB(0, 255, 255)
colours_array(9, 1) = "GPC HOLD"
colours_array(9, 2) = RGB(127, 255, 127)
colours_array(10, 1) = "GPC GAIN"
colours_array(10, 2) = RGB(0, 255, 0)
colours_array(11, 1) = "PPC HOLD"
colours_array(11, 2) = RGB(255, 204, 255)
colours_array(12, 1) = "PPC GAIN"
colours_array(12, 2) = RGB(255, 102, 255)
colours_array(13, 1) = "Other HOLD"
colours_array(13, 2) = RGB(217, 217, 217)
colours_array(14, 1) = "Other GAIN"
colours_array(14, 2) = RGB(128, 128, 128)

Dim const_status_found As Boolean
const_status_found = False

Dim i, j As Integer
'Workbooks(this_wb).Sheets(results_sheet).Cells(results_current_row, results_id_column).Interior.Color = change_colour

Dim constituency As String
Dim input_message As String
Dim majority_delta As Double
Dim new_majoritypc As Double
Dim prev_majoritypc As Double
Dim swing As Double

Dim overall_final_column As Integer
Sheets(sOverall).Select
overall_final_column = WorksheetFunction.CountA(Rows(overall_final_row))

'**************************************************************************************************
'Paste Results in Top Right************************************************************************
'**************************************************************************************************
overall_current_row = overall_heading_row + 1
i = 1

For map_current_row = map_results_start_row To map_results_final_row

    Sheets(sMap).Cells(map_current_row, map_voteshare_column).Value = Sheets(sOverall).Cells(overall_current_row + (2 * delta_rows), overall_final_column).Value
    Sheets(sMap).Cells(map_current_row, map_sharedelta_column).Value = Sheets(sOverall).Cells(overall_current_row + (2 * delta_rows), overall_final_column).Value - Sheets(sOverall).Cells(overall_current_row + (2 * delta_rows), overall_heading_column + 1).Value
    Sheets(sMap).Cells(map_current_row, map_seats_column).Value = Sheets(sOverall).Cells(overall_current_row, overall_final_column).Value
    Sheets(sMap).Cells(map_current_row, map_seatsdelta_column).Value = Sheets(sOverall).Cells(overall_current_row, overall_final_column).Value - Sheets(sOverall).Cells(overall_current_row, overall_heading_column + 1).Value
    Sheets(sMap).Cells(map_current_row, map_hold_column).Interior.Color = colours_array((2 * i - 1), 2)
    Sheets(sMap).Cells(map_current_row, map_gain_column).Interior.Color = colours_array((2 * i), 2)
    
    i = i + 1
    overall_current_row = overall_current_row + 1
       
Next map_current_row


'**************************************************************************************************
'Loop Around Each Constituency in Seats Sheet******************************************************
'**************************************************************************************************
'Select the map sheet so that cells can be selected more easily when it comes to colouring them
Sheets(sMap).Select

For const_current_row = const_heading_row + 1 To const_final_row

    For map_current_row = map_start_row To map_final_row
    
        For map_current_column = map_start_column To map_final_column

            If Sheets(sConst).Cells(const_current_row, const_const_column).Value = Sheets(sMap).Cells(map_current_row, map_current_column).Value Then


'**************************************************************************************************
'Main If Statement for Colouring*******************************************************************
'**************************************************************************************************
                'select current map cell so this statement does not need to be repeated against each colour block
                For i = 1 To 14
               
                    If Sheets(sConst).Cells(const_current_row, const_gain_column).Value = colours_array(i, 1) Then
                        const_status_found = True
                        Exit For
                    Else
                                   
                    End If
    
                Next i
                
                If const_status_found = True Then
                    Sheets(sMap).Cells(map_current_row, map_current_column).Interior.Color = colours_array(i, 2)
                    With Sheets(sMap).Cells(map_current_row, map_current_column).validation
                    .Delete
                    .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
                    
                    constituency = Sheets(sConst).Cells(const_current_row, const_const_column).Value
                    
                    If Len(constituency) > 32 Then
                        .InputTitle = Left(constituency, 32)
                    Else
                        .InputTitle = constituency
                    End If
                    
                    majority_delta = Round((Sheets(sConst).Cells(const_current_row, const_predmajpc_column).Value - Sheets(sConst).Cells(const_current_row, const_2019majpc_column).Value) * 100, 1)
                    new_majoritypc = Round(Sheets(sConst).Cells(const_current_row, const_predmajpc_column).Value * 100, 1)
                    prev_majoritypc = Round(Sheets(sConst).Cells(const_current_row, const_2019majpc_column).Value * 100, 1)
                      
                    If Sheets(sConst).Cells(const_current_row, const_2019winner_column).Value = Sheets(sConst).Cells(const_current_row, const_predwinner_column).Value Then
                                                  
                        If majority_delta >= 0 Then
                            input_message = Sheets(sConst).Cells(const_current_row, const_gain_column).Value & " by " & new_majoritypc & "% (+" & majority_delta & "%)"
                        Else
                            input_message = Sheets(sConst).Cells(const_current_row, const_gain_column).Value & " by " & new_majoritypc & "% (" & majority_delta & "%)"
                        End If
                    Else
                        swing = Round((new_majoritypc + prev_majoritypc) / 2, 1)
                        input_message = Sheets(sConst).Cells(const_current_row, const_gain_column).Value & " from " & Sheets(sConst).Cells(const_current_row, const_2019winner_column).Value & " by " & new_majoritypc & "% (Prev: " & prev_majoritypc & "%)"
                    End If
                    
                    .InputMessage = input_message
                    
                    End With
                Else
                    MsgBox "The status of constituency: " & Sheets(sConst).Cells(const_current_row, const_const_column).Value & " could not be found in the colours_array. This macro will now exit."
                    End
                End If
                
                const_status_found = False
                
            Else
    
            End If
            
        Next map_current_column
            
    Next map_current_row

Next const_current_row

'**************************************************************************************************
'Switch On functionality
'**************************************************************************************************
Call Switch_On_Functionality

'**************************************************************************************************
'Results Summary Message
'**************************************************************************************************
'Determine how many seconds code took to run
 SecondsElapsed = Round(Timer - StartTime, 2)

'String to contain the message contents for the results message
Dim results_summary As String
results_summary = "Map complete in: " & SecondsElapsed & " seconds"
MsgBox results_summary
Sheets(sMap).Select

End Sub
Sub Remove_Date()
Attribute Remove_Date.VB_ProcData.VB_Invoke_Func = "R\n14"
'**************************************************************************************************
'Switch Off Functionality
'**************************************************************************************************
Call Switch_Off_Functionality

'**************************************************************************************************
'Setup Variables
'**************************************************************************************************
'module specific variables
Dim current_column As Integer

'Sheet names
Dim sOverall, sConst, sCand, sMap, sRegions, sShare, sPollsters, sPolls, sParties As String
Call VA_Worksheet_Names(sOverall, sConst, sCand, sMap, sRegions, sShare, sPollsters, sPolls, sParties)

'Candidates variables
Sheets(sCand).Select
Dim cand_heading_row, cand_final_row, cand_const_column, cand_party_column, cand_2019share_column, cand_final_column, cand_totalvotes_column, cand_2019winner_column, cand_predvotes_column, cand_predshare_column, cand_polls_column, cand_dates_column, cand_standing_column As Integer
Call VF_Cand_Variables(cand_heading_row, cand_final_row, cand_const_column, cand_party_column, cand_2019share_column, cand_final_column, cand_totalvotes_column, cand_2019winner_column, cand_predvotes_column, cand_predshare_column, cand_polls_column, cand_dates_column, cand_standing_column)

'Polls sheet variables
Sheets(sPolls).Select
Dim polls_heading_row, polls_final_row, polls_pollster_column, polls_date_column, polls_type_column, polls_scope_column, polls_file_column, polls_candcolumn_column, polls_applicable_column, polls_final_column, polls_const_column, polls_region_column, polls_nation_column As Integer
Call VG_Polls_Variables(polls_heading_row, polls_final_row, polls_pollster_column, polls_date_column, polls_type_column, polls_scope_column, polls_file_column, polls_candcolumn_column, polls_applicable_column, polls_final_column, polls_const_column, polls_region_column, polls_nation_column)

'Constituency sheet variables
Sheets(sConst).Select
Dim const_heading_row, const_final_row, const_final_column, const_region_column, const_const_column, const_2019votes_column, const_2019winner_column, const_2019maj_column, const_2019majpc_column, const_predwinner_column, const_predmaj_column, const_predmajpc_column, const_gain_column, const_loss_column As Integer
Call VH_Const_Variables(const_heading_row, const_final_row, const_final_column, const_region_column, const_const_column, const_2019votes_column, const_2019winner_column, const_2019maj_column, const_2019majpc_column, const_predwinner_column, const_predmaj_column, const_predmajpc_column, const_gain_column, const_loss_column)

'Overall Sheet Variables
Sheets(sOverall).Select
Dim overall_heading_row, overall_heading_column, overall_final_row, overall_no_parties, overall_votes_row, overall_share_row, overall_seatdelta_row As Integer
Call VI_Overall_Variables(overall_heading_row, overall_heading_column, overall_final_row, overall_no_parties, overall_votes_row, overall_share_row, overall_seatdelta_row)

'**************************************************************************************************
'Get assess date from user
'**************************************************************************************************
Dim assess_date, min_date, cand_date As Date
Dim date_string As String
Dim is_date As Boolean
is_date = False

Do While is_date = False
    date_string = InputBox("Enter date to be removed from the model:")
    If IsDate(date_string) Then
        assess_date = DateValue(date_string)
        is_date = True
    Else
        MsgBox "Invalid date, please try again."
    End If
Loop

'**************************************************************************************************
'Remove assess date column in the cand worksheet
'**************************************************************************************************
Sheets(sCand).Select
For current_column = 1 To cand_final_column
    If Sheets(sCand).Cells(cand_heading_row, current_column).Value = assess_date Then
        Columns(current_column).Select
        Selection.Delete
        Exit For
    Else
    End If
Next current_column

'**************************************************************************************************
'Remove assess date column in the const worksheet
'**************************************************************************************************
Sheets(sConst).Select
For current_column = 1 To const_final_column
    If Sheets(sConst).Cells(const_heading_row, current_column).Value = assess_date Then
        Columns(current_column).Select
        Selection.Delete
        Exit For
    Else
    End If
Next current_column


'**************************************************************************************************
'Remove assess date column in the overall worksheet
'**************************************************************************************************
Sheets(sOverall).Select
Dim overall_final_column As Integer
overall_final_column = WorksheetFunction.CountA(Rows(overall_heading_row))
For current_column = 1 To overall_final_column
    If Sheets(sOverall).Cells(overall_heading_row, current_column).Value = assess_date Then
        Columns(current_column).Select
        Selection.Delete
        Exit For
    Else
    End If
Next current_column


'**************************************************************************************************
'Switch On functionality
'**************************************************************************************************
Call Switch_On_Functionality

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

Call Switch_On_Functionality

End

End Sub

Sub Switch_Off_Functionality()
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
End Sub

Sub Switch_On_Functionality()
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True
End Sub

Sub VA_Worksheet_Names(ByRef sOverall, ByRef sConst, ByRef sCand, ByRef sMap, ByRef sRegions, ByRef sShare, ByRef sPollsters, ByRef sPolls, ByRef sParties)

sOverall = "Overall"
sConst = "Const"
sCand = "Cand"
sMap = "Map"
sRegions = "Regions"
sShare = "Share"
sPollsters = "Pollsters"
sPolls = "Polls"
sParties = "Parties"

End Sub

Sub VB_Poll_Variables(ByRef poll_metadata_column, ByRef poll_metadata_row, ByRef poll_pollster_row, ByRef poll_date_row, ByRef poll_type_row, ByRef poll_scope_row, ByRef poll_party_column, ByRef poll_party_row, ByRef poll_region_row, ByRef poll_region_column, ByRef poll_swing_row, ByRef poll_swing_column, ByRef poll_final_row)
search_term = "metadata"
Call Search(search_term, search_row, search_column)
poll_metadata_column = search_column + 1
poll_metadata_row = search_row
poll_final_row = WorksheetFunction.CountA(Columns(search_column))

search_term = "pollster"
Call Search(search_term, search_row, search_column)
poll_pollster_row = search_row

search_term = "date"
Call Search(search_term, search_row, search_column)
poll_date_row = search_row

search_term = "region_type"
Call Search(search_term, search_row, search_column)
poll_type_row = search_row

search_term = "scope"
Call Search(search_term, search_row, search_column)
poll_scope_row = search_row

search_term = "Party"
Call Search(search_term, search_row, search_column)
poll_party_row = search_row
poll_party_column = search_column

search_term = "Region"
Call Search(search_term, search_row, search_column)
poll_region_row = search_row + 1
poll_region_column = search_column

search_term = "Swing"
Call Search(search_term, search_row, search_column)
poll_swing_row = search_row + 1
poll_swing_column = search_column

End Sub

Sub VC_Pollsters_Variables(ByRef pollsters_heading_row, ByRef pollsters_pollster_column, ByRef pollsters_type_column, ByRef pollsters_parties_column, ByRef pollsters_final_row)

search_term = "pollster"
Call Search(search_term, search_row, search_column)
pollsters_pollster_column = search_column
pollsters_heading_row = search_row

search_term = "region_type"
Call Search(search_term, search_row, search_column)
pollsters_type_column = search_column

pollsters_final_row = WorksheetFunction.CountA(Columns(pollsters_type_column))

search_term = "parties"
Call Search(search_term, search_row, search_column)
pollsters_parties_column = search_column

End Sub

Sub VD_Regions_Variables(ByRef regions_heading_row, ByRef regions_type_column, ByRef regions_region_column, ByRef regions_votes_column, ByRef regions_constituencies_column, ByRef regions_final_row)

search_term = "region_type"
Call Search(search_term, search_row, search_column)
regions_type_column = search_column
regions_heading_row = search_row

regions_final_row = WorksheetFunction.CountA(Columns(regions_type_column))

search_term = "region"
Call Search(search_term, search_row, search_column)
regions_region_column = search_column

search_term = "total_votes"
Call Search(search_term, search_row, search_column)
regions_votes_column = search_column

search_term = "constituencies"
Call Search(search_term, search_row, search_column)
regions_constituencies_column = search_column

End Sub

Sub VE_Share_Variables(ByRef share_heading_row, ByRef share_final_row, ByRef share_type_column, ByRef share_region_column, ByRef share_party_column, ByRef share_adjusted_column)
 
search_term = "region_type"
Call Search(search_term, search_row, search_column)
share_heading_row = search_row
share_type_column = search_column
share_final_row = WorksheetFunction.CountA(Columns(share_type_column))

search_term = "region"
Call Search(search_term, search_row, search_column)
share_region_column = search_column

search_term = "party"
Call Search(search_term, search_row, search_column)
share_party_column = search_column

search_term = "adjusted_base_share"
Call Search(search_term, search_row, search_column)
share_adjusted_column = search_column

End Sub

Sub VF_Cand_Variables(ByRef cand_heading_row, ByRef cand_final_row, ByRef cand_const_column, ByRef cand_party_column, ByRef cand_2019share_column, ByRef cand_final_column, ByRef cand_totalvotes_column, ByRef cand_2019winner_column, ByRef cand_predvotes_column, ByRef cand_predshare_column, ByRef cand_polls_column, ByRef cand_dates_column, ByRef cand_standing_column)

search_term = "constituency_name"
Call Search(search_term, search_row, search_column)
cand_heading_row = search_row
cand_const_column = search_column
cand_final_row = WorksheetFunction.CountA(Columns(cand_const_column))
cand_final_column = WorksheetFunction.CountA(Rows(cand_heading_row)) + 1

search_term = "party"
Call Search(search_term, search_row, search_column)
cand_party_column = search_column

search_term = "share_2019"
Call Search(search_term, search_row, search_column)
cand_2019share_column = search_column

search_term = "total_votes_2019"
Call Search(search_term, search_row, search_column)
cand_totalvotes_column = search_column

search_term = "winner_2019"
Call Search(search_term, search_row, search_column)
cand_2019winner_column = search_column

search_term = "votes_pred"
Call Search(search_term, search_row, search_column)
cand_predvotes_column = search_column

search_term = "share_pred"
Call Search(search_term, search_row, search_column)
cand_predshare_column = search_column

search_term = "Polls"
Call Search(search_term, search_row, search_column)
cand_polls_column = search_column

search_term = "Dates"
Call Search(search_term, search_row, search_column)
cand_dates_column = search_column

search_term = "standing"
Call Search(search_term, search_row, search_column)
cand_standing_column = search_column

End Sub

Sub VG_Polls_Variables(ByRef polls_heading_row, ByRef polls_final_row, ByRef polls_pollster_column, ByRef polls_date_column, ByRef polls_type_column, ByRef polls_scope_column, ByRef polls_file_column, ByRef polls_candcolumn_column, ByRef polls_applicable_column, ByRef polls_final_column, ByRef polls_const_column, ByRef polls_region_column, ByRef polls_nation_column)

search_term = "pollster"
Call Search(search_term, search_row, search_column)
polls_heading_row = search_row
polls_pollster_column = search_column
polls_final_row = WorksheetFunction.CountA(Columns(polls_pollster_column))
polls_final_column = WorksheetFunction.CountA(Rows(polls_heading_row))

search_term = "date"
Call Search(search_term, search_row, search_column)
polls_date_column = search_column

search_term = "region_type"
Call Search(search_term, search_row, search_column)
polls_type_column = search_column

search_term = "scope"
Call Search(search_term, search_row, search_column)
polls_scope_column = search_column

search_term = "file_name"
Call Search(search_term, search_row, search_column)
polls_file_column = search_column

search_term = "cand_column"
Call Search(search_term, search_row, search_column)
polls_candcolumn_column = search_column

search_term = "applicable"
Call Search(search_term, search_row, search_column)
polls_applicable_column = search_column

search_term = "const"
Call Search(search_term, search_row, search_column)
polls_const_column = search_column

search_term = "region"
Call Search(search_term, search_row, search_column)
polls_region_column = search_column

search_term = "nation"
Call Search(search_term, search_row, search_column)
polls_nation_column = search_column

End Sub

Sub VH_Const_Variables(ByRef const_heading_row, ByRef const_final_row, ByRef const_final_column, ByRef const_region_column, ByRef const_const_column, ByRef const_2019votes_column, ByRef const_2019winner_column, ByRef const_2019maj_column, ByRef const_2019majpc_column, ByRef const_predwinner_column, ByRef const_predmaj_column, ByRef const_predmajpc_column, ByRef const_gain_column, ByRef const_loss_column)
search_term = "region"
Call Search(search_term, search_row, search_column)
const_heading_row = search_row
const_region_column = search_column
const_final_row = WorksheetFunction.CountA(Columns(const_region_column))
const_final_column = WorksheetFunction.CountA(Rows(const_heading_row))

search_term = "constituency"
Call Search(search_term, search_row, search_column)
const_const_column = search_column

search_term = "votes_2019"
Call Search(search_term, search_row, search_column)
const_2019votes_column = search_column

search_term = "winner_2019"
Call Search(search_term, search_row, search_column)
const_2019winner_column = search_column

search_term = "maj_2019"
Call Search(search_term, search_row, search_column)
const_2019maj_column = search_column

search_term = "maj_pc_2019"
Call Search(search_term, search_row, search_column)
const_2019majpc_column = search_column

search_term = "winner_pred"
Call Search(search_term, search_row, search_column)
const_predwinner_column = search_column

search_term = "maj_pred"
Call Search(search_term, search_row, search_column)
const_predmaj_column = search_column

search_term = "maj_pc_pred"
Call Search(search_term, search_row, search_column)
const_predmajpc_column = search_column

search_term = "GAIN/HOLD"
Call Search(search_term, search_row, search_column)
const_gain_column = search_column

search_term = "LOSS/HOLD"
Call Search(search_term, search_row, search_column)
const_loss_column = search_column

End Sub

Sub VI_Overall_Variables(ByRef overall_heading_row, ByRef overall_heading_column, ByRef overall_final_row, ByRef overall_no_parties, ByRef overall_votes_row, ByRef overall_share_row, ByRef overall_seatdelta_row)

search_term = "Party"
Call Search(search_term, search_row, search_column)
overall_heading_row = search_row
overall_heading_column = search_column
overall_final_row = WorksheetFunction.CountA(Columns(overall_heading_column))

search_term = "Total"
Call Search(search_term, search_row, search_column)
overall_no_parties = search_row - overall_heading_row - 1

search_term = "Votes"
Call Search(search_term, search_row, search_column)
overall_votes_row = search_row

search_term = "Share"
Call Search(search_term, search_row, search_column)
overall_share_row = search_row

search_term = "Seat Delta"
Call Search(search_term, search_row, search_column)
overall_seatdelta_row = search_row

End Sub

Sub VM_Map_Variables(ByRef map_start_row, ByRef map_final_row, ByRef map_start_column, ByRef map_final_column, ByRef map_party_column, ByRef map_voteshare_column, ByRef map_sharedelta_column, ByRef map_seats_column, ByRef map_seatsdelta_column, ByRef map_gain_column, ByRef map_results_start_row, ByRef map_hold_column, ByRef map_results_final_row, ByRef no_parties, ByRef extra_gap, ByRef delta_rows)

map_start_row = 1
map_final_row = 25
map_start_column = 1
map_final_column = 60

search_term = "Party"
Call Search(search_term, search_row, search_column)
map_party_column = search_column

search_term = "Share"
Call Search(search_term, search_row, search_column)
map_voteshare_column = search_column
map_sharedelta_column = search_column + 2

search_term = "Seats"
Call Search(search_term, search_row, search_column)
map_seats_column = search_column
map_seatsdelta_column = search_column + 2

search_term = "G"
Call Search(search_term, search_row, search_column)
map_gain_column = search_column
map_results_start_row = search_row + 1

search_term = "H"
Call Search(search_term, search_row, search_column)
map_hold_column = search_column

search_term = "Other"
Call Search(search_term, search_row, search_column)
map_results_final_row = search_row

no_parties = 7
extra_gap = 2
delta_rows = no_parties + extra_gap

End Sub
