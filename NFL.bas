Attribute VB_Name = "NFL"
Sub A_Populate_R()
Attribute A_Populate_R.VB_ProcData.VB_Invoke_Func = "A\n14"

Call F_Switch_Off_Functionality

'Get current cell co-ordinates
Dim games_current_row As Long
games_current_row = ActiveCell.Row
current_sheet = ActiveSheet.Name

'Sheet names
Dim elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet As String
Call VA_Worksheet_Names(elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet)

If current_sheet = games_sheet Then
Else
    MsgBox games_sheet & " worksheet is not selected. This macro will now exit."
    Call O_Switch_On_Functionality
    End
End If

'Games variables
Sheets(games_sheet).Select
Dim games_winner_column, games_gameref_column, games_ra_column, games_rh_column, games_rhh_column, games_ea_column, games_eh_column, games_year_column, games_week_column, games_yw_column, games_awayteam_column, games_hometeam_column, games_neutral_column, games_prediction_column, games_awayscore_column, games_homescore_column, games_winnerstatus_column, games_predcorrect_column, games_sa_column, games_sh_column, games_rd_column, games_radash_column, games_rhdash_column, games_predtotal_column As Integer
Call VG_Games_Sheet_Variables(games_winner_column, games_gameref_column, games_ra_column, games_rh_column, games_rhh_column, games_ea_column, games_eh_column, games_year_column, games_week_column, games_yw_column, games_awayteam_column, games_hometeam_column, games_neutral_column, games_prediction_column, games_awayscore_column, games_homescore_column, games_winnerstatus_column, games_predcorrect_column, games_sa_column, games_sh_column, games_rd_column, games_radash_column, games_rhdash_column, games_predtotal_column)

'Elo sheet variables
Sheets(elo_sheet).Select
Dim elo_ref_column, elo_year_column, elo_team_column, elo_yw_column, elo_week_column, elo_elo_column As Integer
Call VE_Elo_Sheet_Variables(elo_ref_column, elo_year_column, elo_team_column, elo_yw_column, elo_week_column, elo_elo_column)

Dim elo_start_row As Integer
elo_start_row = 2

Dim elo_final_row As Long
elo_final_row = WorksheetFunction.CountA(Columns(elo_yw_column))

Dim elo_current_row As Long

Dim modified_elo, current_elo As Double

'H variables
Dim H As Double
H = Range("H").Value

Sheets(games_sheet).Select

'Error messages
If IsEmpty(Sheets(games_sheet).Cells(games_current_row, games_year_column)) = True Then
    MsgBox "The YEAR is not populated against the current game. This macro will now exit."
    Call O_Switch_On_Functionality
    End
Else
End If

If IsEmpty(Sheets(games_sheet).Cells(games_current_row, games_week_column)) = True Then
    MsgBox "The WEEK is not populated against the current game. This macro will now exit."
    Call O_Switch_On_Functionality
    End
Else
End If

If IsEmpty(Sheets(games_sheet).Cells(games_current_row, games_awayteam_column)) = True Then
    MsgBox "The AWAYTEAM is not populated against the current game. This macro will now exit."
    Call O_Switch_On_Functionality
    End
Else
End If

If IsEmpty(Sheets(games_sheet).Cells(games_current_row, games_hometeam_column)) = True Then
    MsgBox "The HOMETEAM is not populated against the current game. This macro will now exit."
    Call O_Switch_On_Functionality
    End
Else
End If

If IsEmpty(Sheets(games_sheet).Cells(games_current_row, games_neutral_column)) = True Then
    MsgBox "The NEUTRAL (stadium) column is not populated against the current game. This macro will now exit."
    Call O_Switch_On_Functionality
    End
Else
End If

If IsEmpty(Sheets(games_sheet).Cells(games_current_row, games_ra_column)) = False Then
    MsgBox "The RA column is populated against the current game. This macro will now exit."
    Call O_Switch_On_Functionality
    End
Else
End If

'determine teams to find
Dim away_team As String
Dim home_team As String
Dim team_found As Boolean
team_found = False

away_team = Sheets(games_sheet).Cells(games_current_row, games_awayteam_column).Value
home_team = Sheets(games_sheet).Cells(games_current_row, games_hometeam_column).Value

'check to see if either team has an outstanding game
'games_current_row
Dim games_start_row, games_findteam_row As Integer
games_start_row = 2
games_findteam_row = games_current_row - 1

Do Until games_findteam_row < games_start_row
    
    If IsEmpty(Sheets(games_sheet).Cells(games_findteam_row, games_winner_column)) = True Then
        If Sheets(games_sheet).Cells(games_findteam_row, games_awayteam_column).Value = away_team Or Sheets(games_sheet).Cells(games_findteam_row, games_hometeam_column).Value = away_team Then
            MsgBox "The Away Team: " & away_team & " has an outstanding game. This macro will now exit."
            MsgBox "Row: " & games_findteam_row
            Call O_Switch_On_Functionality
            End
    
        ElseIf Sheets(games_sheet).Cells(games_findteam_row, games_awayteam_column).Value = home_team Or Sheets(games_sheet).Cells(games_findteam_row, games_hometeam_column).Value = home_team Then
            MsgBox "The Home Team: " & home_team & " has an outstanding game. This macro will now exit."
            MsgBox "Row: " & games_findteam_row
            Call O_Switch_On_Functionality
            End
        Else
        End If
    Else
    End If
    
    games_findteam_row = games_findteam_row - 1
Loop

'Find and populate away team R
elo_current_row = elo_final_row
team_found = False

Do Until elo_current_row < elo_start_row

    If Sheets(elo_sheet).Cells(elo_current_row, elo_team_column).Value = away_team Then
            
        current_elo = Sheets(elo_sheet).Cells(elo_current_row, elo_elo_column).Value
    
        team_found = True
        Exit Do
        
    Else
    
    End If

    elo_current_row = elo_current_row - 1
Loop

If team_found = False Then
    MsgBox "The Away Team: " & away_team & " has not been found in the database. This macro will now exit."
    Call O_Switch_On_Functionality
    End
Else
    Sheets(games_sheet).Cells(games_current_row, games_ra_column).Value = current_elo
End If

'Find and populate home team R
team_found = False
elo_current_row = elo_final_row

Do Until elo_current_row < elo_start_row

    If Sheets(elo_sheet).Cells(elo_current_row, elo_team_column).Value = home_team Then

        current_elo = Sheets(elo_sheet).Cells(elo_current_row, elo_elo_column).Value
        team_found = True
        Exit Do
        
    Else
    
    End If

    elo_current_row = elo_current_row - 1
Loop

If team_found = False Then
    Sheets(games_sheet).Cells(games_current_row, games_ra_column).Value = ""
    MsgBox "The Home Team: " & home_team & " has not been found in the database. This macro will now exit."
    Call O_Switch_On_Functionality
    End
Else
    Sheets(games_sheet).Cells(games_current_row, games_rh_column).Value = current_elo
    If Sheets(games_sheet).Cells(games_current_row, games_neutral_column).Value = "Y" Then
        Sheets(games_sheet).Cells(games_current_row, games_rhh_column).Value = current_elo
    Else
        Sheets(games_sheet).Cells(games_current_row, games_rhh_column).Value = current_elo + H
    End If

End If

'populate game info (first two columns)
Dim gameref As String
gameref = "G" & Sheets(games_sheet).Cells(games_current_row, games_year_column).Value & Sheets(games_sheet).Cells(games_current_row, games_week_column).Value & Sheets(games_sheet).Cells(games_current_row, games_awayteam_column).Value & Sheets(games_sheet).Cells(games_current_row, games_hometeam_column).Value
Sheets(games_sheet).Cells(games_current_row, games_gameref_column).Value = gameref

gameweek = Sheets(games_sheet).Cells(games_current_row, games_week_column).Value
If gameweek < 10 Then
    Sheets(games_sheet).Cells(games_current_row, games_yw_column).Value = Sheets(games_sheet).Cells(games_current_row, games_year_column).Value & "-0" & gameweek
Else
    Sheets(games_sheet).Cells(games_current_row, games_yw_column).Value = Sheets(games_sheet).Cells(games_current_row, games_year_column).Value & "-" & gameweek
End If

'Populate Es
Sheets(games_sheet).Cells(games_current_row, games_ea_column).Value = 1 / (1 + 10 ^ ((Sheets(games_sheet).Cells(games_current_row, games_rhh_column).Value - Sheets(games_sheet).Cells(games_current_row, games_ra_column).Value) / 400))
Sheets(games_sheet).Cells(games_current_row, games_eh_column).Value = 1 / (1 + 10 ^ ((Sheets(games_sheet).Cells(games_current_row, games_ra_column).Value - Sheets(games_sheet).Cells(games_current_row, games_rhh_column).Value) / 400))

'Populate Prediction
If Sheets(games_sheet).Cells(games_current_row, games_ea_column).Value > Sheets(games_sheet).Cells(games_current_row, games_eh_column).Value Then

    Sheets(games_sheet).Cells(games_current_row, games_prediction_column).Value = away_team

Else

    Sheets(games_sheet).Cells(games_current_row, games_prediction_column).Value = home_team

End If

Sheets(games_sheet).Cells(games_current_row + 1, games_awayteam_column).Select

Call O_Switch_On_Functionality

End Sub

Sub B_Odds()
Attribute B_Odds.VB_ProcData.VB_Invoke_Func = "B\n14"
On Err GoTo Reset:

Call F_Switch_Off_Functionality

'Get current cell co-ordinates
Dim odds_current_row As Long
odds_current_row = ActiveCell.Row

Dim odds_start_row As Long
odds_start_row = odds_current_row

'Sheet names
Dim elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet As String
Call VA_Worksheet_Names(elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet)

current_sheet = ActiveSheet.Name

If current_sheet = odds_sheet Then
Else
    MsgBox odds_sheet & " worksheet is not selected. This macro will now exit."
    GoTo Reset:
End If

'Odds variables
Sheets(odds_sheet).Select
Dim odds_gameref_column, odds_datetime_column, odds_awayodds_column, odds_homeodds_column, odds_awayelodec_column, odds_homeelodec_column, odds_awayoddsdec_column, odds_homeoddsdec_column, odds_awayedge_column, odds_homeedge_column, odds_betteam_column, odds_bethan_column, odds_kelly_column, odds_betamount_column, odds_winamount_column, odds_awayreqdec_column, odds_homereqdec_column, odds_awayreqodds_column, odds_homereqodds_column, odds_year_column, odds_week_column, odds_betplaced_column, odds_betodds_column, odds_betedge_column, odds_winlose_column, odds_opposeteam_column, odds_edgeedge_column, odds_betunits_column, odds_awaytrend_column, odds_hometrend_column, odds_neutral_column, odds_awayteam_column, odds_hometeam_column, odds_profitloss_column, odds_winnerstatus_column, odds_cumulative_column As Integer
Call VO_Odds_Sheet_Variables(odds_gameref_column, odds_datetime_column, odds_awayodds_column, odds_homeodds_column, odds_awayelodec_column, odds_homeelodec_column, odds_awayoddsdec_column, odds_homeoddsdec_column, odds_awayedge_column, odds_homeedge_column, odds_betteam_column, odds_bethan_column, odds_kelly_column, odds_betamount_column, odds_winamount_column, odds_awayreqdec_column, odds_homereqdec_column, odds_awayreqodds_column, odds_homereqodds_column, odds_year_column, odds_week_column, odds_betplaced_column, odds_betodds_column, odds_betedge_column, odds_winlose_column, odds_opposeteam_column, odds_edgeedge_column, odds_betunits_column, odds_awaytrend_column, odds_hometrend_column, odds_neutral_column, odds_awayteam_column, odds_hometeam_column, odds_profitloss_column, odds_winnerstatus_column, odds_cumulative_column)

Dim odds_betfind_row As Integer

'Games variables
Sheets(games_sheet).Select
Dim games_winner_column, games_gameref_column, games_ra_column, games_rh_column, games_rhh_column, games_ea_column, games_eh_column, games_year_column, games_week_column, games_yw_column, games_awayteam_column, games_hometeam_column, games_neutral_column, games_prediction_column, games_awayscore_column, games_homescore_column, games_winnerstatus_column, games_predcorrect_column, games_sa_column, games_sh_column, games_rd_column, games_radash_column, games_rhdash_column, games_predtotal_column As Integer
Call VG_Games_Sheet_Variables(games_winner_column, games_gameref_column, games_ra_column, games_rh_column, games_rhh_column, games_ea_column, games_eh_column, games_year_column, games_week_column, games_yw_column, games_awayteam_column, games_hometeam_column, games_neutral_column, games_prediction_column, games_awayscore_column, games_homescore_column, games_winnerstatus_column, games_predcorrect_column, games_sa_column, games_sh_column, games_rd_column, games_radash_column, games_rhdash_column, games_predtotal_column)

Dim games_final_row As Integer
games_final_row = WorksheetFunction.CountA(Columns(games_year_column))

'Name variables
Dim A1WL, A1WU, A1OL1, A1OU1, A1E1, A1OL2, A1OU2, A1E2, A1OL3, A1OU3, A1E3, A1OL4, A1OU4, A1E4, A1OL5, A1OU5, A1E5, A1OL6, A1OU6, A1E6, A1OL7, A1OU7, A1E7, A1OL8, A1OU8, A1E8, A1OL9, A1OU9, A1E9, A1OL10, A1OU10, A1E10, A1OL11, A1OU11, A1E11, A1OL12, A1OU12, A1E12 As Double
Dim A2WL, A2WU, A2OL1, A2OU1, A2E1, A2OL2, A2OU2, A2E2, A2OL3, A2OU3, A2E3, A2OL4, A2OU4, A2E4, A2OL5, A2OU5, A2E5, A2OL6, A2OU6, A2E6, A2OL7, A2OU7, A2E7, A2OL8, A2OU8, A2E8, A2OL9, A2OU9, A2E9, A2OL10, A2OU10, A2E10, A2OL11, A2OU11, A2E11, A2OL12, A2OU12, A2E12 As Double
Dim A3WL, A3WU, A3OL1, A3OU1, A3E1, A3OL2, A3OU2, A3E2, A3OL3, A3OU3, A3E3, A3OL4, A3OU4, A3E4, A3OL5, A3OU5, A3E5, A3OL6, A3OU6, A3E6, A3OL7, A3OU7, A3E7, A3OL8, A3OU8, A3E8, A3OL9, A3OU9, A3E9, A3OL10, A3OU10, A3E10, A3OL11, A3OU11, A3E11, A3OL12, A3OU12, A3E12 As Double

Dim H1WL, H1WU, H1OL1, H1OU1, H1E1, H1OL2, H1OU2, H1E2, H1OL3, H1OU3, H1E3, H1OL4, H1OU4, H1E4, H1OL5, H1OU5, H1E5, H1OL6, H1OU6, H1E6, H1OL7, H1OU7, H1E7, H1OL8, H1OU8, H1E8, H1OL9, H1OU9, H1E9, H1OL10, H1OU10, H1E10, H1OL11, H1OU11, H1E11, H1OL12, H1OU12, H1E12 As Double
Dim H2WL, H2WU, H2OL1, H2OU1, H2E1, H2OL2, H2OU2, H2E2, H2OL3, H2OU3, H2E3, H2OL4, H2OU4, H2E4, H2OL5, H2OU5, H2E5, H2OL6, H2OU6, H2E6, H2OL7, H2OU7, H2E7, H2OL8, H2OU8, H2E8, H2OL9, H2OU9, H2E9, H2OL10, H2OU10, H2E10, H2OL11, H2OU11, H2E11, H2OL12, H2OU12, H2E12 As Double
Dim H3WL, H3WU, H3OL1, H3OU1, H3E1, H3OL2, H3OU2, H3E2, H3OL3, H3OU3, H3E3, H3OL4, H3OU4, H3E4, H3OL5, H3OU5, H3E5, H3OL6, H3OU6, H3E6, H3OL7, H3OU7, H3E7, H3OL8, H3OU8, H3E8, H3OL9, H3OU9, H3E9, H3OL10, H3OU10, H3E10, H3OL11, H3OU11, H3E11, H3OL12, H3OU12, H3E12 As Double

A1WL = Range("A_1_WL").Value
A1WU = Range("A_1_WU").Value
A1OL1 = Range("A_1_OL_1").Value
A1OU1 = Range("A_1_OU_1").Value
A1E1 = Range("A_1_E_1").Value
A1OL2 = Range("A_1_OL_2").Value
A1OU2 = Range("A_1_OU_2").Value
A1E2 = Range("A_1_E_2").Value
A1OL3 = Range("A_1_OL_3").Value
A1OU3 = Range("A_1_OU_3").Value
A1E3 = Range("A_1_E_3").Value
A1OL4 = Range("A_1_OL_4").Value
A1OU4 = Range("A_1_OU_4").Value
A1E4 = Range("A_1_E_4").Value
A1OL5 = Range("A_1_OL_5").Value
A1OU5 = Range("A_1_OU_5").Value
A1E5 = Range("A_1_E_5").Value
A1OL6 = Range("A_1_OL_6").Value
A1OU6 = Range("A_1_OU_6").Value
A1E6 = Range("A_1_E_6").Value
A1OL7 = Range("A_1_OL_7").Value
A1OU7 = Range("A_1_OU_7").Value
A1E7 = Range("A_1_E_7").Value
A1OL8 = Range("A_1_OL_8").Value
A1OU8 = Range("A_1_OU_8").Value
A1E8 = Range("A_1_E_8").Value
A1OL9 = Range("A_1_OL_9").Value
A1OU9 = Range("A_1_OU_9").Value
A1E9 = Range("A_1_E_9").Value
A1OL10 = Range("A_1_OL_10").Value
A1OU10 = Range("A_1_OU_10").Value
A1E10 = Range("A_1_E_10").Value
A1OL11 = Range("A_1_OL_11").Value
A1OU11 = Range("A_1_OU_11").Value
A1E11 = Range("A_1_E_11").Value
A1OL12 = Range("A_1_OL_12").Value
A1OU12 = Range("A_1_OU_12").Value
A1E12 = Range("A_1_E_12").Value

A2WL = Range("A_2_WL").Value
A2WU = Range("A_2_WU").Value
A2OL1 = Range("A_2_OL_1").Value
A2OU1 = Range("A_2_OU_1").Value
A2E1 = Range("A_2_E_1").Value
A2OL2 = Range("A_2_OL_2").Value
A2OU2 = Range("A_2_OU_2").Value
A2E2 = Range("A_2_E_2").Value
A2OL3 = Range("A_2_OL_3").Value
A2OU3 = Range("A_2_OU_3").Value
A2E3 = Range("A_2_E_3").Value
A2OL4 = Range("A_2_OL_4").Value
A2OU4 = Range("A_2_OU_4").Value
A2E4 = Range("A_2_E_4").Value
A2OL5 = Range("A_2_OL_5").Value
A2OU5 = Range("A_2_OU_5").Value
A2E5 = Range("A_2_E_5").Value
A2OL6 = Range("A_2_OL_6").Value
A2OU6 = Range("A_2_OU_6").Value
A2E6 = Range("A_2_E_6").Value
A2OL7 = Range("A_2_OL_7").Value
A2OU7 = Range("A_2_OU_7").Value
A2E7 = Range("A_2_E_7").Value
A2OL8 = Range("A_2_OL_8").Value
A2OU8 = Range("A_2_OU_8").Value
A2E8 = Range("A_2_E_8").Value
A2OL9 = Range("A_2_OL_9").Value
A2OU9 = Range("A_2_OU_9").Value
A2E9 = Range("A_2_E_9").Value
A2OL10 = Range("A_2_OL_10").Value
A2OU10 = Range("A_2_OU_10").Value
A2E10 = Range("A_2_E_10").Value
A2OL11 = Range("A_2_OL_11").Value
A2OU11 = Range("A_2_OU_11").Value
A2E11 = Range("A_2_E_11").Value
A2OL12 = Range("A_2_OL_12").Value
A2OU12 = Range("A_2_OU_12").Value
A2E12 = Range("A_2_E_12").Value

A3WL = Range("A_3_WL").Value
A3WU = Range("A_3_WU").Value
A3OL1 = Range("A_3_OL_1").Value
A3OU1 = Range("A_3_OU_1").Value
A3E1 = Range("A_3_E_1").Value
A3OL2 = Range("A_3_OL_2").Value
A3OU2 = Range("A_3_OU_2").Value
A3E2 = Range("A_3_E_2").Value
A3OL3 = Range("A_3_OL_3").Value
A3OU3 = Range("A_3_OU_3").Value
A3E3 = Range("A_3_E_3").Value
A3OL4 = Range("A_3_OL_4").Value
A3OU4 = Range("A_3_OU_4").Value
A3E4 = Range("A_3_E_4").Value
A3OL5 = Range("A_3_OL_5").Value
A3OU5 = Range("A_3_OU_5").Value
A3E5 = Range("A_3_E_5").Value
A3OL6 = Range("A_3_OL_6").Value
A3OU6 = Range("A_3_OU_6").Value
A3E6 = Range("A_3_E_6").Value
A3OL7 = Range("A_3_OL_7").Value
A3OU7 = Range("A_3_OU_7").Value
A3E7 = Range("A_3_E_7").Value
A3OL8 = Range("A_3_OL_8").Value
A3OU8 = Range("A_3_OU_8").Value
A3E8 = Range("A_3_E_8").Value
A3OL9 = Range("A_3_OL_9").Value
A3OU9 = Range("A_3_OU_9").Value
A3E9 = Range("A_3_E_9").Value
A3OL10 = Range("A_3_OL_10").Value
A3OU10 = Range("A_3_OU_10").Value
A3E10 = Range("A_3_E_10").Value
A3OL11 = Range("A_3_OL_11").Value
A3OU11 = Range("A_3_OU_11").Value
A3E11 = Range("A_3_E_11").Value
A3OL12 = Range("A_3_OL_12").Value
A3OU12 = Range("A_3_OU_12").Value
A3E12 = Range("A_3_E_12").Value

H1WL = Range("H_1_WL").Value
H1WU = Range("H_1_WU").Value
H1OL1 = Range("H_1_OL_1").Value
H1OU1 = Range("H_1_OU_1").Value
H1E1 = Range("H_1_E_1").Value
H1OL2 = Range("H_1_OL_2").Value
H1OU2 = Range("H_1_OU_2").Value
H1E2 = Range("H_1_E_2").Value
H1OL3 = Range("H_1_OL_3").Value
H1OU3 = Range("H_1_OU_3").Value
H1E3 = Range("H_1_E_3").Value
H1OL4 = Range("H_1_OL_4").Value
H1OU4 = Range("H_1_OU_4").Value
H1E4 = Range("H_1_E_4").Value
H1OL5 = Range("H_1_OL_5").Value
H1OU5 = Range("H_1_OU_5").Value
H1E5 = Range("H_1_E_5").Value
H1OL6 = Range("H_1_OL_6").Value
H1OU6 = Range("H_1_OU_6").Value
H1E6 = Range("H_1_E_6").Value
H1OL7 = Range("H_1_OL_7").Value
H1OU7 = Range("H_1_OU_7").Value
H1E7 = Range("H_1_E_7").Value
H1OL8 = Range("H_1_OL_8").Value
H1OU8 = Range("H_1_OU_8").Value
H1E8 = Range("H_1_E_8").Value
H1OL9 = Range("H_1_OL_9").Value
H1OU9 = Range("H_1_OU_9").Value
H1E9 = Range("H_1_E_9").Value
H1OL10 = Range("H_1_OL_10").Value
H1OU10 = Range("H_1_OU_10").Value
H1E10 = Range("H_1_E_10").Value
H1OL11 = Range("H_1_OL_11").Value
H1OU11 = Range("H_1_OU_11").Value
H1E11 = Range("H_1_E_11").Value
H1OL12 = Range("H_1_OL_12").Value
H1OU12 = Range("H_1_OU_12").Value
H1E12 = Range("H_1_E_12").Value

H2WL = Range("H_2_WL").Value
H2WU = Range("H_2_WU").Value
H2OL1 = Range("H_2_OL_1").Value
H2OU1 = Range("H_2_OU_1").Value
H2E1 = Range("H_2_E_1").Value
H2OL2 = Range("H_2_OL_2").Value
H2OU2 = Range("H_2_OU_2").Value
H2E2 = Range("H_2_E_2").Value
H2OL3 = Range("H_2_OL_3").Value
H2OU3 = Range("H_2_OU_3").Value
H2E3 = Range("H_2_E_3").Value
H2OL4 = Range("H_2_OL_4").Value
H2OU4 = Range("H_2_OU_4").Value
H2E4 = Range("H_2_E_4").Value
H2OL5 = Range("H_2_OL_5").Value
H2OU5 = Range("H_2_OU_5").Value
H2E5 = Range("H_2_E_5").Value
H2OL6 = Range("H_2_OL_6").Value
H2OU6 = Range("H_2_OU_6").Value
H2E6 = Range("H_2_E_6").Value
H2OL7 = Range("H_2_OL_7").Value
H2OU7 = Range("H_2_OU_7").Value
H2E7 = Range("H_2_E_7").Value
H2OL8 = Range("H_2_OL_8").Value
H2OU8 = Range("H_2_OU_8").Value
H2E8 = Range("H_2_E_8").Value
H2OL9 = Range("H_2_OL_9").Value
H2OU9 = Range("H_2_OU_9").Value
H2E9 = Range("H_2_E_9").Value
H2OL10 = Range("H_2_OL_10").Value
H2OU10 = Range("H_2_OU_10").Value
H2E10 = Range("H_2_E_10").Value
H2OL11 = Range("H_2_OL_11").Value
H2OU11 = Range("H_2_OU_11").Value
H2E11 = Range("H_2_E_11").Value
H2OL12 = Range("H_2_OL_12").Value
H2OU12 = Range("H_2_OU_12").Value
H2E12 = Range("H_2_E_12").Value

H3WL = Range("H_3_WL").Value
H3WU = Range("H_3_WU").Value
H3OL1 = Range("H_3_OL_1").Value
H3OU1 = Range("H_3_OU_1").Value
H3E1 = Range("H_3_E_1").Value
H3OL2 = Range("H_3_OL_2").Value
H3OU2 = Range("H_3_OU_2").Value
H3E2 = Range("H_3_E_2").Value
H3OL3 = Range("H_3_OL_3").Value
H3OU3 = Range("H_3_OU_3").Value
H3E3 = Range("H_3_E_3").Value
H3OL4 = Range("H_3_OL_4").Value
H3OU4 = Range("H_3_OU_4").Value
H3E4 = Range("H_3_E_4").Value
H3OL5 = Range("H_3_OL_5").Value
H3OU5 = Range("H_3_OU_5").Value
H3E5 = Range("H_3_E_5").Value
H3OL6 = Range("H_3_OL_6").Value
H3OU6 = Range("H_3_OU_6").Value
H3E6 = Range("H_3_E_6").Value
H3OL7 = Range("H_3_OL_7").Value
H3OU7 = Range("H_3_OU_7").Value
H3E7 = Range("H_3_E_7").Value
H3OL8 = Range("H_3_OL_8").Value
H3OU8 = Range("H_3_OU_8").Value
H3E8 = Range("H_3_E_8").Value
H3OL9 = Range("H_3_OL_9").Value
H3OU9 = Range("H_3_OU_9").Value
H3E9 = Range("H_3_E_9").Value
H3OL10 = Range("H_3_OL_10").Value
H3OU10 = Range("H_3_OU_10").Value
H3E10 = Range("H_3_E_10").Value
H3OL11 = Range("H_3_OL_11").Value
H3OU11 = Range("H_3_OU_11").Value
H3E11 = Range("H_3_E_11").Value
H3OL12 = Range("H_3_OL_12").Value
H3OU12 = Range("H_3_OU_12").Value
H3E12 = Range("H_3_E_12").Value

Dim KPP As Double
KPP = Range("KPP").Value

Dim bankroll As Double
bankroll = Range("B").Value

Dim E1 As Double
E1 = Range("E_1").Value

Dim neutral As String

Dim bet_amount As Double

Dim gameref As String

Dim games_start_row As Integer
games_start_row = 2

Dim games_current_row As Long
Dim gameref_found As Boolean
gameref_found = False

'Count the number of rows to be populated
Sheets(odds_sheet).Select

Dim odds_final_row As Integer
odds_final_row = WorksheetFunction.CountA(Columns(odds_gameref_column))

Dim odds_rowstopop As Integer
odds_rowstopop = 0

Dim awaysoddsdec, homeoddsdec, awayedge, homeedge As Double

Dim awayelodec, homeelodec, awayreqdec, homereqdec, awayreqodds, homereqodds As Double

Dim week, year As Integer

For odds_current_row = odds_start_row To odds_final_row
    If IsEmpty(Cells(odds_current_row, odds_awayelodec_column)) = True Then
        odds_rowstopop = odds_rowstopop + 1
    Else
        Exit For
    End If
Next odds_current_row

'Test message
'MsgBox "Number of rows to populate with calcs: " & odds_rowstopop

For odds_current_row = odds_start_row To odds_rowstopop + odds_start_row - 1

    'Error checker
    If IsEmpty(Sheets(odds_sheet).Cells(odds_current_row, odds_gameref_column)) = True Then
        MsgBox "The GAMEREF is not populated against the current odds. This macro will now exit."
        GoTo Reset:
    Else
    End If
       
    If IsEmpty(Sheets(odds_sheet).Cells(odds_current_row, odds_awayodds_column)) = True Then
        MsgBox "The AWAYODDS are not populated against the current odds. This macro will now exit."
        GoTo Reset:
    Else
    End If
    
    If IsEmpty(Sheets(odds_sheet).Cells(odds_current_row, odds_homeodds_column)) = True Then
        MsgBox "The HOMEODDS are not populated against the current odds. This macro will now exit."
        GoTo Reset:
    Else
    End If
    
    'Identify the game reference to look for in the games sheets
    gameref = Sheets(odds_sheet).Cells(odds_current_row, odds_gameref_column).Value
    
    'find which row in the games sheet the game is
    For games_current_row = games_start_row To games_final_row
    
        If Sheets(games_sheet).Cells(games_current_row, games_gameref_column).Value = gameref Then
    
            gameref_found = True
            Exit For
    
        Else
            gameref_found = False
        End If
    
    Next games_current_row

    If gameref_found = False Then
    
        MsgBox "Gameref not found in worksheet: " & games_sheet & ". This macro will now exit."
        GoTo Reset:
    
    Else
    
    End If
    
    'populate date & time
    Sheets(odds_sheet).Cells(odds_current_row, odds_datetime_column).Value = Now()
    
    'Populate week and year columns
    year = Sheets(games_sheet).Cells(games_current_row, games_year_column).Value
    Sheets(odds_sheet).Cells(odds_current_row, odds_year_column).Value = year
    week = Sheets(games_sheet).Cells(games_current_row, games_week_column).Value
    Sheets(odds_sheet).Cells(odds_current_row, odds_week_column).Value = week
    
    'Calculate AwayEloDec
    Sheets(odds_sheet).Cells(odds_current_row, odds_awayelodec_column).Value = Sheets(games_sheet).Cells(games_current_row, games_eh_column).Value / Sheets(games_sheet).Cells(games_current_row, games_ea_column).Value
    
    'Calculate HomeEloDec
    Sheets(odds_sheet).Cells(odds_current_row, odds_homeelodec_column).Value = Sheets(games_sheet).Cells(games_current_row, games_ea_column).Value / Sheets(games_sheet).Cells(games_current_row, games_eh_column).Value
    
    'Calculate AwaysOddsDec
    If Sheets(odds_sheet).Cells(odds_current_row, odds_awayodds_column).Value = 0 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_awayoddsdec_column).Value = 0
        
    ElseIf Sheets(odds_sheet).Cells(odds_current_row, odds_awayodds_column).Value < 100 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_awayoddsdec_column).Value = -100 / Sheets(odds_sheet).Cells(odds_current_row, odds_awayodds_column).Value
    
    Else
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_awayoddsdec_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_awayodds_column).Value / 100
    
    End If
    
    'Calculate HomeOddsDec
    If Sheets(odds_sheet).Cells(odds_current_row, odds_homeodds_column).Value = 0 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_homeoddsdec_column).Value = 0
        
    ElseIf Sheets(odds_sheet).Cells(odds_current_row, odds_homeodds_column).Value < 100 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_homeoddsdec_column).Value = -100 / Sheets(odds_sheet).Cells(odds_current_row, odds_homeodds_column).Value
    
    Else
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_homeoddsdec_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_homeodds_column).Value / 100
    
    End If
    
    'Calculate away edge
    Sheets(odds_sheet).Cells(odds_current_row, odds_awayedge_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_awayoddsdec_column).Value - Sheets(odds_sheet).Cells(odds_current_row, odds_awayelodec_column).Value
    
    'Calculate home edge
    Sheets(odds_sheet).Cells(odds_current_row, odds_homeedge_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_homeoddsdec_column).Value - Sheets(odds_sheet).Cells(odds_current_row, odds_homeelodec_column).Value
    
    'Calculate Neutral venue or not
    neutral = Sheets(games_sheet).Cells(games_current_row, games_neutral_column).Value
    
    'populate neutral column
    Sheets(odds_sheet).Cells(odds_current_row, odds_neutral_column).Value = neutral
    
    'Calculate bet HAN
    awayoddsdec = Sheets(odds_sheet).Cells(odds_current_row, odds_awayoddsdec_column).Value
    homeoddsdec = Sheets(odds_sheet).Cells(odds_current_row, odds_homeoddsdec_column).Value
    awayedge = Sheets(odds_sheet).Cells(odds_current_row, odds_awayedge_column).Value
    homeedge = Sheets(odds_sheet).Cells(odds_current_row, odds_homeedge_column).Value
    
    If week >= H1WL And week <= H1WU And homeoddsdec > H1OL1 And homeoddsdec <= H1OU1 And homeedge >= H1E1 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec > A1OL1 And awayoddsdec <= A1OU1 And awayedge >= A1E1 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"
        
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec > H1OL2 And homeoddsdec <= H1OU2 And homeedge >= H1E2 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec > A1OL2 And awayoddsdec <= A1OU2 And awayedge >= A1E2 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"
        
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec > H1OL3 And homeoddsdec <= H1OU3 And homeedge >= H1E3 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec > A1OL3 And awayoddsdec <= A1OU3 And awayedge >= A1E3 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"
        
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec > H1OL4 And homeoddsdec <= H1OU4 And homeedge >= H1E4 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec > A1OL4 And awayoddsdec <= A1OU4 And awayedge >= A1E4 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H1WL And week <= H1WU And homeoddsdec > H1OL5 And homeoddsdec <= H1OU5 And homeedge >= H1E5 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec > A1OL5 And awayoddsdec <= A1OU5 And awayedge >= A1E5 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H1WL And week <= H1WU And homeoddsdec > H1OL6 And homeoddsdec <= H1OU6 And homeedge >= H1E6 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec > A1OL6 And awayoddsdec <= A1OU6 And awayedge >= A1E6 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H1WL And week <= H1WU And homeoddsdec > H1OL7 And homeoddsdec <= H1OU7 And homeedge >= H1E7 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec > A1OL7 And awayoddsdec <= A1OU7 And awayedge >= A1E7 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H1WL And week <= H1WU And homeoddsdec > H1OL8 And homeoddsdec <= H1OU8 And homeedge >= H1E8 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec > A1OL8 And awayoddsdec <= A1OU8 And awayedge >= A1E8 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H1WL And week <= H1WU And homeoddsdec > H1OL9 And homeoddsdec <= H1OU9 And homeedge >= H1E9 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec > A1OL9 And awayoddsdec <= A1OU9 And awayedge >= A1E9 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H1WL And week <= H1WU And homeoddsdec > H1OL10 And homeoddsdec <= H1OU10 And homeedge >= H1E10 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec > A1OL10 And awayoddsdec <= A1OU10 And awayedge >= A1E10 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H1WL And week <= H1WU And homeoddsdec > H1OL11 And homeoddsdec <= H1OU11 And homeedge >= H1E11 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec > A1OL11 And awayoddsdec <= A1OU11 And awayedge >= A1E11 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H1WL And week <= H1WU And homeoddsdec > H1OL12 And homeoddsdec <= H1OU12 And homeedge >= H1E12 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec > A1OL12 And awayoddsdec <= A1OU12 And awayedge >= A1E12 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"
        
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec > H2OL1 And homeoddsdec <= H2OU1 And homeedge >= H2E1 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec > A2OL1 And awayoddsdec <= A2OU1 And awayedge >= A2E1 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"
        
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec > H2OL2 And homeoddsdec <= H2OU2 And homeedge >= H2E2 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec > A2OL2 And awayoddsdec <= A2OU2 And awayedge >= A2E2 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"
        
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec > H2OL3 And homeoddsdec <= H2OU3 And homeedge >= H2E3 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec > A2OL3 And awayoddsdec <= A2OU3 And awayedge >= A2E3 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"
        
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec > H2OL4 And homeoddsdec <= H2OU4 And homeedge >= H2E4 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec > A2OL4 And awayoddsdec <= A2OU4 And awayedge >= A2E4 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H2WL And week <= H2WU And homeoddsdec > H2OL5 And homeoddsdec <= H2OU5 And homeedge >= H2E5 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec > A2OL5 And awayoddsdec <= A2OU5 And awayedge >= A2E5 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H2WL And week <= H2WU And homeoddsdec > H2OL6 And homeoddsdec <= H2OU6 And homeedge >= H2E6 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec > A2OL6 And awayoddsdec <= A2OU6 And awayedge >= A2E6 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H2WL And week <= H2WU And homeoddsdec > H2OL7 And homeoddsdec <= H2OU7 And homeedge >= H2E7 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec > A2OL7 And awayoddsdec <= A2OU7 And awayedge >= A2E7 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H2WL And week <= H2WU And homeoddsdec > H2OL8 And homeoddsdec <= H2OU8 And homeedge >= H2E8 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec > A2OL8 And awayoddsdec <= A2OU8 And awayedge >= A2E8 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H2WL And week <= H2WU And homeoddsdec > H2OL9 And homeoddsdec <= H2OU9 And homeedge >= H2E9 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec > A2OL9 And awayoddsdec <= A2OU9 And awayedge >= A2E9 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H2WL And week <= H2WU And homeoddsdec > H2OL10 And homeoddsdec <= H2OU10 And homeedge >= H2E10 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec > A2OL10 And awayoddsdec <= A2OU10 And awayedge >= A2E10 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H2WL And week <= H2WU And homeoddsdec > H2OL11 And homeoddsdec <= H2OU11 And homeedge >= H2E11 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec > A2OL11 And awayoddsdec <= A2OU11 And awayedge >= A2E11 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H2WL And week <= H2WU And homeoddsdec > H2OL12 And homeoddsdec <= H2OU12 And homeedge >= H2E12 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec > A2OL12 And awayoddsdec <= A2OU12 And awayedge >= A2E12 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"
        
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec > H3OL1 And homeoddsdec <= H3OU1 And homeedge >= H3E1 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec > A3OL1 And awayoddsdec <= A3OU1 And awayedge >= A3E1 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"
        
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec > H3OL2 And homeoddsdec <= H3OU2 And homeedge >= H3E2 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec > A3OL2 And awayoddsdec <= A3OU2 And awayedge >= A3E2 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"
        
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec > H3OL3 And homeoddsdec <= H3OU3 And homeedge >= H3E3 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec > A3OL3 And awayoddsdec <= A3OU3 And awayedge >= A3E3 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"
        
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec > H3OL4 And homeoddsdec <= H3OU4 And homeedge >= H3E4 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec > A3OL4 And awayoddsdec <= A3OU4 And awayedge >= A3E4 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H3WL And week <= H3WU And homeoddsdec > H3OL5 And homeoddsdec <= H3OU5 And homeedge >= H3E5 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec > A3OL5 And awayoddsdec <= A3OU5 And awayedge >= A3E5 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H3WL And week <= H3WU And homeoddsdec > H3OL6 And homeoddsdec <= H3OU6 And homeedge >= H3E6 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec > A3OL6 And awayoddsdec <= A3OU6 And awayedge >= A3E6 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H3WL And week <= H3WU And homeoddsdec > H3OL7 And homeoddsdec <= H3OU7 And homeedge >= H3E7 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec > A3OL7 And awayoddsdec <= A3OU7 And awayedge >= A3E7 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H3WL And week <= H3WU And homeoddsdec > H3OL8 And homeoddsdec <= H3OU8 And homeedge >= H3E8 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec > A3OL8 And awayoddsdec <= A3OU8 And awayedge >= A3E8 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H3WL And week <= H3WU And homeoddsdec > H3OL9 And homeoddsdec <= H3OU9 And homeedge >= H3E9 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec > A3OL9 And awayoddsdec <= A3OU9 And awayedge >= A3E9 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H3WL And week <= H3WU And homeoddsdec > H3OL10 And homeoddsdec <= H3OU10 And homeedge >= H3E10 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec > A3OL10 And awayoddsdec <= A3OU10 And awayedge >= A3E10 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H3WL And week <= H3WU And homeoddsdec > H3OL11 And homeoddsdec <= H3OU11 And homeedge >= H3E11 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec > A3OL11 And awayoddsdec <= A3OU11 And awayedge >= A3E11 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"

    ElseIf week >= H3WL And week <= H3WU And homeoddsdec > H3OL12 And homeoddsdec <= H3OU12 And homeedge >= H3E12 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME"
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec > A3OL12 And awayoddsdec <= A3OU12 And awayedge >= A3E12 Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY"
   
    Else
     
        Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "NONE"
        
    End If
    
    'Calculate bet team and non-bet team
    If Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "NONE" Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_betteam_column).Value = "NONE"
        Sheets(odds_sheet).Cells(odds_current_row, odds_opposeteam_column).Value = "NONE"
        
    ElseIf Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY" Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_betteam_column).Value = Sheets(games_sheet).Cells(games_current_row, games_awayteam_column).Value
        Sheets(odds_sheet).Cells(odds_current_row, odds_opposeteam_column).Value = Sheets(games_sheet).Cells(games_current_row, games_hometeam_column).Value
        
    ElseIf Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME" Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_betteam_column).Value = Sheets(games_sheet).Cells(games_current_row, games_hometeam_column).Value
        Sheets(odds_sheet).Cells(odds_current_row, odds_opposeteam_column).Value = Sheets(games_sheet).Cells(games_current_row, games_awayteam_column).Value
    
    Else
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_betteam_column).Value = "NONE"
        Sheets(odds_sheet).Cells(odds_current_row, odds_opposeteam_column).Value = "NONE"
        
    End If
    
    'Populate home team nad away team
    Sheets(odds_sheet).Cells(odds_current_row, odds_awayteam_column).Value = Sheets(games_sheet).Cells(games_current_row, games_awayteam_column).Value
    Sheets(odds_sheet).Cells(odds_current_row, odds_hometeam_column).Value = Sheets(games_sheet).Cells(games_current_row, games_hometeam_column).Value
        
    'Calculate Kelly%
    If Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "NONE" Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_kelly_column).Value = 0
        
    ElseIf Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY" Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_kelly_column).Value = Sheets(games_sheet).Cells(games_current_row, games_ea_column).Value - (Sheets(games_sheet).Cells(games_current_row, games_eh_column).Value / (awayoddsdec - E1))
    
    ElseIf Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME" Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_kelly_column).Value = Sheets(games_sheet).Cells(games_current_row, games_eh_column).Value - (Sheets(games_sheet).Cells(games_current_row, games_ea_column).Value / (homeoddsdec - E1))
        
    Else
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_kelly_column).Value = 0
    
    End If
    
    'Calculate bet amount
    bet_amount = Round(KPP * Sheets(odds_sheet).Cells(odds_current_row, odds_kelly_column).Value * bankroll, 0)
    
    Sheets(odds_sheet).Cells(odds_current_row, odds_betamount_column).Value = bet_amount
       
    'Calculate required odds
    awayelodec = Sheets(odds_sheet).Cells(odds_current_row, odds_awayelodec_column).Value
    homeelodec = Sheets(odds_sheet).Cells(odds_current_row, odds_homeelodec_column).Value
    
    'Awayrequired odds
    If week >= A1WL And week <= A1WU And awayoddsdec <= A1OU1 And awayelodec + A1E1 < 0 Then
        awayreqdec = 0.01
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU1 And awayelodec < A1OU1 - A1E1 Then
        awayreqdec = awayelodec + A1E1
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU1 And awayelodec >= A1OU1 - A1E1 Then
        awayreqdec = awayelodec + A1E2
        
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU2 And awayelodec + A1E2 < 0 Then
        awayreqdec = 0.01
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU2 And awayelodec < A1OU2 - A1E2 Then
        awayreqdec = awayelodec + A1E2
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU2 And awayelodec >= A1OU2 - A1E2 Then
        awayreqdec = awayelodec + A1E3
        
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU3 And awayelodec < A1OU3 - A1E3 Then
        awayreqdec = awayelodec + A1E3
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU3 And awayelodec >= A1OU3 - A1E3 Then
        awayreqdec = awayelodec + A1E4
        
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU4 And awayelodec < A1OU4 - A1E4 Then
        awayreqdec = awayelodec + A1E4
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU4 And awayelodec >= A1OU4 - A1E4 Then
        awayreqdec = awayelodec + A1E5
        
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU5 And awayelodec < A1OU5 - A1E5 Then
        awayreqdec = awayelodec + A1E5
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU5 And awayelodec >= A1OU5 - A1E5 Then
        awayreqdec = awayelodec + A1E6
        
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU6 And awayelodec < A1OU6 - A1E6 Then
        awayreqdec = awayelodec + A1E6
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU6 And awayelodec >= A1OU6 - A1E6 Then
        awayreqdec = awayelodec + A1E7
        
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU7 And awayelodec < A1OU7 - A1E7 Then
        awayreqdec = awayelodec + A1E7
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU7 And awayelodec >= A1OU7 - A1E7 Then
        awayreqdec = awayelodec + A1E8
        
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU8 And awayelodec < A1OU8 - A1E8 Then
        awayreqdec = awayelodec + A1E8
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU8 And awayelodec >= A1OU8 - A1E8 Then
        awayreqdec = awayelodec + A1E9

    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU9 And awayelodec < A1OU9 - A1E9 Then
        awayreqdec = awayelodec + A1E9
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU9 And awayelodec >= A1OU9 - A1E9 Then
        awayreqdec = awayelodec + A1E10
        
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU10 And awayelodec < A1OU10 - A1E10 Then
        awayreqdec = awayelodec + A1E10
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU10 And awayelodec >= A1OU10 - A1E10 Then
        awayreqdec = awayelodec + A1E11
        
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU11 And awayelodec < A1OU11 - A1E11 Then
        awayreqdec = awayelodec + A1E11
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU11 And awayelodec >= A1OU11 - A1E11 Then
        awayreqdec = awayelodec + A1E12
        
    ElseIf week >= A1WL And week <= A1WU And awayoddsdec <= A1OU12 Then
        awayreqdec = awayelodec + A1E12
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU1 And awayelodec + A2E1 < 0 Then
        awayreqdec = 0.01
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU1 And awayelodec < A2OU1 - A2E1 Then
        awayreqdec = awayelodec + A2E1
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU1 And awayelodec >= A2OU1 - A2E1 Then
        awayreqdec = awayelodec + A2E2
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU2 And awayelodec + A2E2 < 0 Then
        awayreqdec = 0.01
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU2 And awayelodec < A2OU2 - A2E2 Then
        awayreqdec = awayelodec + A2E2
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU2 And awayelodec >= A2OU2 - A2E2 Then
        awayreqdec = awayelodec + A2E3
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU3 And awayelodec < A2OU3 - A2E3 Then
        awayreqdec = awayelodec + A2E3
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU3 And awayelodec >= A2OU3 - A2E3 Then
        awayreqdec = awayelodec + A2E4
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU4 And awayelodec < A2OU4 - A2E4 Then
        awayreqdec = awayelodec + A2E4
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU4 And awayelodec >= A2OU4 - A2E4 Then
        awayreqdec = awayelodec + A2E5
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU5 And awayelodec < A2OU5 - A2E5 Then
        awayreqdec = awayelodec + A2E5
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU5 And awayelodec >= A2OU5 - A2E5 Then
        awayreqdec = awayelodec + A2E6
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU6 And awayelodec < A2OU6 - A2E6 Then
        awayreqdec = awayelodec + A2E6
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU6 And awayelodec >= A2OU6 - A2E6 Then
        awayreqdec = awayelodec + A2E7
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU7 And awayelodec < A2OU7 - A2E7 Then
        awayreqdec = awayelodec + A2E7
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU7 And awayelodec >= A2OU7 - A2E7 Then
        awayreqdec = awayelodec + A2E8
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU8 And awayelodec < A2OU8 - A2E8 Then
        awayreqdec = awayelodec + A2E8
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU8 And awayelodec >= A2OU8 - A2E8 Then
        awayreqdec = awayelodec + A2E9

    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU9 And awayelodec < A2OU9 - A2E9 Then
        awayreqdec = awayelodec + A2E9
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU9 And awayelodec >= A2OU9 - A2E9 Then
        awayreqdec = awayelodec + A2E10
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU10 And awayelodec < A2OU10 - A2E10 Then
        awayreqdec = awayelodec + A2E10
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU10 And awayelodec >= A2OU10 - A2E10 Then
        awayreqdec = awayelodec + A2E11
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU11 And awayelodec < A2OU11 - A2E11 Then
        awayreqdec = awayelodec + A2E11
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU11 And awayelodec >= A2OU11 - A2E11 Then
        awayreqdec = awayelodec + A2E12
        
    ElseIf week >= A2WL And week <= A2WU And awayoddsdec <= A2OU12 Then
        awayreqdec = awayelodec + A2E12
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU1 And awayelodec + A3E1 < 0 Then
        awayreqdec = 0.01
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU1 And awayelodec < A3OU1 - A3E1 Then
        awayreqdec = awayelodec + A3E1
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU1 And awayelodec >= A3OU1 - A3E1 Then
        awayreqdec = awayelodec + A3E2
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU2 And awayelodec + A3E2 < 0 Then
        awayreqdec = 0.01
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU2 And awayelodec < A3OU2 - A3E2 Then
        awayreqdec = awayelodec + A3E2
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU2 And awayelodec >= A3OU2 - A3E2 Then
        awayreqdec = awayelodec + A3E3
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU3 And awayelodec < A3OU3 - A3E3 Then
        awayreqdec = awayelodec + A3E3
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU3 And awayelodec >= A3OU3 - A3E3 Then
        awayreqdec = awayelodec + A3E4
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU4 And awayelodec < A3OU4 - A3E4 Then
        awayreqdec = awayelodec + A3E4
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU4 And awayelodec >= A3OU4 - A3E4 Then
        awayreqdec = awayelodec + A3E5
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU5 And awayelodec < A3OU5 - A3E5 Then
        awayreqdec = awayelodec + A3E5
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU5 And awayelodec >= A3OU5 - A3E5 Then
        awayreqdec = awayelodec + A3E6
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU6 And awayelodec < A3OU6 - A3E6 Then
        awayreqdec = awayelodec + A3E6
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU6 And awayelodec >= A3OU6 - A3E6 Then
        awayreqdec = awayelodec + A3E7
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU7 And awayelodec < A3OU7 - A3E7 Then
        awayreqdec = awayelodec + A3E7
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU7 And awayelodec >= A3OU7 - A3E7 Then
        awayreqdec = awayelodec + A3E8
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU8 And awayelodec < A3OU8 - A3E8 Then
        awayreqdec = awayelodec + A3E8
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU8 And awayelodec >= A3OU8 - A3E8 Then
        awayreqdec = awayelodec + A3E9

    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU9 And awayelodec < A3OU9 - A3E9 Then
        awayreqdec = awayelodec + A3E9
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU9 And awayelodec >= A3OU9 - A3E9 Then
        awayreqdec = awayelodec + A3E10
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU10 And awayelodec < A3OU10 - A3E10 Then
        awayreqdec = awayelodec + A3E10
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU10 And awayelodec >= A3OU10 - A3E10 Then
        awayreqdec = awayelodec + A3E11
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU11 And awayelodec < A3OU11 - A3E11 Then
        awayreqdec = awayelodec + A3E11
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU11 And awayelodec >= A3OU11 - A3E11 Then
        awayreqdec = awayelodec + A3E12
        
    ElseIf week >= A3WL And week <= A3WU And awayoddsdec <= A3OU12 Then
        awayreqdec = awayelodec + A3E12
       
    End If
    
    'Homerequired odds
    If week >= H1WL And week <= H1WU And homeoddsdec <= H1OU1 And homeelodec + H1E1 < 0 Then
        homereqdec = 0.01
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU1 And homeelodec < H1OU1 - H1E1 Then
        homereqdec = homeelodec + H1E1
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU1 And homeelodec >= H1OU1 - H1E1 Then
        homereqdec = homeelodec + H1E2
        
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU2 And homeelodec + H1E2 < 0 Then
        homereqdec = 0.01
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU2 And homeelodec < H1OU2 - H1E2 Then
        homereqdec = homeelodec + H1E2
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU2 And homeelodec >= H1OU2 - H1E2 Then
        homereqdec = homeelodec + H1E3
        
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU3 And homeelodec < H1OU3 - H1E3 Then
        homereqdec = homeelodec + H1E3
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU3 And homeelodec >= H1OU3 - H1E3 Then
        homereqdec = homeelodec + H1E4
        
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU4 And homeelodec < H1OU4 - H1E4 Then
        homereqdec = homeelodec + H1E4
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU4 And homeelodec >= H1OU4 - H1E4 Then
        homereqdec = homeelodec + H1E5
        
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU5 And homeelodec < H1OU5 - H1E5 Then
        homereqdec = homeelodec + H1E5
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU5 And homeelodec >= H1OU5 - H1E5 Then
        homereqdec = homeelodec + H1E6
        
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU6 And homeelodec < H1OU6 - H1E6 Then
        homereqdec = homeelodec + H1E6
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU6 And homeelodec >= H1OU6 - H1E6 Then
        homereqdec = homeelodec + H1E7
        
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU7 And homeelodec < H1OU7 - H1E7 Then
        homereqdec = homeelodec + H1E7
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU7 And homeelodec >= H1OU7 - H1E7 Then
        homereqdec = homeelodec + H1E8
        
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU8 And homeelodec < H1OU8 - H1E8 Then
        homereqdec = homeelodec + H1E8
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU8 And homeelodec >= H1OU8 - H1E8 Then
        homereqdec = homeelodec + H1E9

    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU9 And homeelodec < H1OU9 - H1E9 Then
        homereqdec = homeelodec + H1E9
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU9 And homeelodec >= H1OU9 - H1E9 Then
        homereqdec = homeelodec + H1E10
        
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU10 And homeelodec < H1OU10 - H1E10 Then
        homereqdec = homeelodec + H1E10
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU10 And homeelodec >= H1OU10 - H1E10 Then
        homereqdec = homeelodec + H1E11
        
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU11 And homeelodec < H1OU11 - H1E11 Then
        homereqdec = homeelodec + H1E11
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU11 And homeelodec >= H1OU11 - H1E11 Then
        homereqdec = homeelodec + H1E12
        
    ElseIf week >= H1WL And week <= H1WU And homeoddsdec <= H1OU12 Then
        homereqdec = homeelodec + H1E12
    
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU1 And homeelodec + H2E1 < 0 Then
        homereqdec = 0.01
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU1 And homeelodec < H2OU1 - H2E1 Then
        homereqdec = homeelodec + H2E1
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU1 And homeelodec >= H2OU1 - H2E1 Then
        homereqdec = homeelodec + H2E2
        
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU2 And homeelodec + H2E2 < 0 Then
        homereqdec = 0.01
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU2 And homeelodec < H2OU2 - H2E2 Then
        homereqdec = homeelodec + H2E2
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU2 And homeelodec >= H2OU2 - H2E2 Then
        homereqdec = homeelodec + H2E3
        
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU3 And homeelodec < H2OU3 - H2E3 Then
        homereqdec = homeelodec + H2E3
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU3 And homeelodec >= H2OU3 - H2E3 Then
        homereqdec = homeelodec + H2E4
        
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU4 And homeelodec < H2OU4 - H2E4 Then
        homereqdec = homeelodec + H2E4
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU4 And homeelodec >= H2OU4 - H2E4 Then
        homereqdec = homeelodec + H2E5
        
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU5 And homeelodec < H2OU5 - H2E5 Then
        homereqdec = homeelodec + H2E5
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU5 And homeelodec >= H2OU5 - H2E5 Then
        homereqdec = homeelodec + H2E6
        
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU6 And homeelodec < H2OU6 - H2E6 Then
        homereqdec = homeelodec + H2E6
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU6 And homeelodec >= H2OU6 - H2E6 Then
        homereqdec = homeelodec + H2E7
        
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU7 And homeelodec < H2OU7 - H2E7 Then
        homereqdec = homeelodec + H2E7
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU7 And homeelodec >= H2OU7 - H2E7 Then
        homereqdec = homeelodec + H2E8
        
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU8 And homeelodec < H2OU8 - H2E8 Then
        homereqdec = homeelodec + H2E8
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU8 And homeelodec >= H2OU8 - H2E8 Then
        homereqdec = homeelodec + H2E9

    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU9 And homeelodec < H2OU9 - H2E9 Then
        homereqdec = homeelodec + H2E9
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU9 And homeelodec >= H2OU9 - H2E9 Then
        homereqdec = homeelodec + H2E10
        
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU10 And homeelodec < H2OU10 - H2E10 Then
        homereqdec = homeelodec + H2E10
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU10 And homeelodec >= H2OU10 - H2E10 Then
        homereqdec = homeelodec + H2E11
        
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU11 And homeelodec < H2OU11 - H2E11 Then
        homereqdec = homeelodec + H2E11
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU11 And homeelodec >= H2OU11 - H2E11 Then
        homereqdec = homeelodec + H2E12
        
    ElseIf week >= H2WL And week <= H2WU And homeoddsdec <= H2OU12 Then
        homereqdec = homeelodec + H2E12
        
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU1 And homeelodec + H3E1 < 0 Then
        homereqdec = 0.01
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU1 And homeelodec < H3OU1 - H3E1 Then
        homereqdec = homeelodec + H3E1
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU1 And homeelodec >= H3OU1 - H3E1 Then
        homereqdec = homeelodec + H3E2
        
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU2 And homeelodec + H3E2 < 0 Then
        homereqdec = 0.01
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU2 And homeelodec < H3OU2 - H3E2 Then
        homereqdec = homeelodec + H3E2
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU2 And homeelodec >= H3OU2 - H3E2 Then
        homereqdec = homeelodec + H3E3
        
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU3 And homeelodec < H3OU3 - H3E3 Then
        homereqdec = homeelodec + H3E3
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU3 And homeelodec >= H3OU3 - H3E3 Then
        homereqdec = homeelodec + H3E4
        
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU4 And homeelodec < H3OU4 - H3E4 Then
        homereqdec = homeelodec + H3E4
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU4 And homeelodec >= H3OU4 - H3E4 Then
        homereqdec = homeelodec + H3E5
        
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU5 And homeelodec < H3OU5 - H3E5 Then
        homereqdec = homeelodec + H3E5
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU5 And homeelodec >= H3OU5 - H3E5 Then
        homereqdec = homeelodec + H3E6
        
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU6 And homeelodec < H3OU6 - H3E6 Then
        homereqdec = homeelodec + H3E6
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU6 And homeelodec >= H3OU6 - H3E6 Then
        homereqdec = homeelodec + H3E7
        
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU7 And homeelodec < H3OU7 - H3E7 Then
        homereqdec = homeelodec + H3E7
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU7 And homeelodec >= H3OU7 - H3E7 Then
        homereqdec = homeelodec + H3E8
        
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU8 And homeelodec < H3OU8 - H3E8 Then
        homereqdec = homeelodec + H3E8
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU8 And homeelodec >= H3OU8 - H3E8 Then
        homereqdec = homeelodec + H3E9

    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU9 And homeelodec < H3OU9 - H3E9 Then
        homereqdec = homeelodec + H3E9
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU9 And homeelodec >= H3OU9 - H3E9 Then
        homereqdec = homeelodec + H3E10
        
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU10 And homeelodec < H3OU10 - H3E10 Then
        homereqdec = homeelodec + H3E10
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU10 And homeelodec >= H3OU10 - H3E10 Then
        homereqdec = homeelodec + H3E11
        
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU11 And homeelodec < H3OU11 - H3E11 Then
        homereqdec = homeelodec + H3E11
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU11 And homeelodec >= H3OU11 - H3E11 Then
        homereqdec = homeelodec + H3E12
        
    ElseIf week >= H3WL And week <= H3WU And homeoddsdec <= H3OU12 Then
        homereqdec = homeelodec + H3E12
       
    End If
    
    'awaysoddsdec, homeoddsdec
    Sheets(odds_sheet).Cells(odds_current_row, odds_awayreqdec_column).Value = awayreqdec
    Sheets(odds_sheet).Cells(odds_current_row, odds_homereqdec_column).Value = homereqdec
        
    If awayreqdec > 1 Then
    
        awayreqodds = awayreqdec * 100
    
    Else
        
        awayreqodds = -100 / awayreqdec
        
    End If
    
    Sheets(odds_sheet).Cells(odds_current_row, odds_awayreqodds_column).Value = awayreqodds
    
    If homereqdec > 1 Then
    
        homereqodds = homereqdec * 100
    
    Else
        
        homereqodds = -100 / homereqdec
        
    End If
    
    Sheets(odds_sheet).Cells(odds_current_row, odds_homereqodds_column).Value = homereqodds
    
    'Calculate bet units
    Sheets(odds_sheet).Cells(odds_current_row, odds_betunits_column).Value = Round(KPP * Sheets(odds_sheet).Cells(odds_current_row, odds_kelly_column).Value * 100, 2)
    
    'Calculate win amount and betodds, bet edge, edgeedge
    If Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "NONE" Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_winamount_column).Value = 0
        Sheets(odds_sheet).Cells(odds_current_row, odds_betodds_column).Value = 0
        Sheets(odds_sheet).Cells(odds_current_row, odds_betedge_column).Value = 0
        Sheets(odds_sheet).Cells(odds_current_row, odds_edgeedge_column).Value = 0
                
    ElseIf Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY" Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_winamount_column).Value = Round(awayoddsdec * Sheets(odds_sheet).Cells(odds_current_row, odds_betamount_column).Value, 2)
        Sheets(odds_sheet).Cells(odds_current_row, odds_betodds_column).Value = awayoddsdec
        Sheets(odds_sheet).Cells(odds_current_row, odds_betedge_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_awayoddsdec_column).Value - Sheets(odds_sheet).Cells(odds_current_row, odds_awayelodec_column).Value
        Sheets(odds_sheet).Cells(odds_current_row, odds_edgeedge_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_awayoddsdec_column).Value - Sheets(odds_sheet).Cells(odds_current_row, odds_awayreqdec_column).Value
           
    ElseIf Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME" Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_winamount_column).Value = Round(homeoddsdec * Sheets(odds_sheet).Cells(odds_current_row, odds_betamount_column).Value, 2)
        Sheets(odds_sheet).Cells(odds_current_row, odds_betodds_column).Value = homeoddsdec
        Sheets(odds_sheet).Cells(odds_current_row, odds_betedge_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_homeoddsdec_column).Value - Sheets(odds_sheet).Cells(odds_current_row, odds_homeelodec_column).Value
        Sheets(odds_sheet).Cells(odds_current_row, odds_edgeedge_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_homeoddsdec_column).Value - Sheets(odds_sheet).Cells(odds_current_row, odds_homereqdec_column).Value
       
       
    Else
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_winamount_column).Value = "ERROR"
        Sheets(odds_sheet).Cells(odds_current_row, odds_betodds_column).Value = "ERROR"
        Sheets(odds_sheet).Cells(odds_current_row, odds_betedge_column).Value = "ERROR"
        Sheets(odds_sheet).Cells(odds_current_row, odds_edgeedge_column).Value = "ERROR"
            
    End If
      
    'Format cells as required
    If Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "NONE" Then
        
        'away formatting
        If Cells(odds_current_row, odds_awayodds_column).Value > awayreqodds - Abs(0.05 * awayreqodds) Then
        
            Sheets(odds_sheet).Cells(odds_current_row, odds_awayodds_column).Interior.Color = RGB(255, 204, 102)
        
        ElseIf Cells(odds_current_row, odds_awayodds_column).Value > awayreqodds - Abs(0.1 * awayreqodds) Then
        
            Sheets(odds_sheet).Cells(odds_current_row, odds_awayodds_column).Interior.Color = RGB(255, 153, 153)
        
        Else
        
            Sheets(odds_sheet).Cells(odds_current_row, odds_awayodds_column).Interior.Color = RGB(255, 255, 255)
        
        End If
        
        'home formatting
        If Cells(odds_current_row, odds_homeodds_column).Value > homereqodds - Abs(0.05 * homereqodds) Then
        
            Sheets(odds_sheet).Cells(odds_current_row, odds_homeodds_column).Interior.Color = RGB(255, 204, 102)
        
        ElseIf Cells(odds_current_row, odds_homeodds_column).Value > homereqodds - Abs(0.1 * homereqodds) Then
        
            Sheets(odds_sheet).Cells(odds_current_row, odds_homeodds_column).Interior.Color = RGB(255, 153, 153)
        
        Else
        
            Sheets(odds_sheet).Cells(odds_current_row, odds_homeodds_column).Interior.Color = RGB(255, 255, 255)
        
        End If
        
    ElseIf Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY" Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_awayodds_column).Interior.Color = RGB(153, 255, 153)
        Sheets(odds_sheet).Cells(odds_current_row, odds_homeodds_column).Interior.Color = RGB(255, 255, 255)
    
    ElseIf Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME" Then
    
        Sheets(odds_sheet).Cells(odds_current_row, odds_awayodds_column).Interior.Color = RGB(255, 255, 255)
        Sheets(odds_sheet).Cells(odds_current_row, odds_homeodds_column).Interior.Color = RGB(153, 255, 153)
    Else
            
    End If
    
    'Determine if
    For odds_betfind_row = 2 To odds_final_row

        If Sheets(odds_sheet).Cells(odds_betfind_row, odds_gameref_column).Value = gameref Then
            If Sheets(odds_sheet).Cells(odds_betfind_row, odds_betplaced_column).Value = "Y" Then
                
                If Sheets(odds_sheet).Cells(odds_betfind_row, odds_bethan_column).Value = "AWAY" Then
                    Sheets(odds_sheet).Cells(odds_current_row, odds_awayodds_column).Interior.Color = RGB(191, 191, 191)
                    Sheets(odds_sheet).Cells(odds_current_row, odds_betplaced_column).Value = "N"
    
                ElseIf Sheets(odds_sheet).Cells(odds_betfind_row, odds_bethan_column).Value = "HOME" Then
                    Sheets(odds_sheet).Cells(odds_current_row, odds_homeodds_column).Interior.Color = RGB(191, 191, 191)
                    Sheets(odds_sheet).Cells(odds_current_row, odds_betplaced_column).Value = "N"
                Else
    
                End If
    
            Else
    
            End If
        Else
    
        End If

    Next odds_betfind_row
    
    'Work out the odds trends
    Dim odds_trend_row As Integer
    odds_trend_row = odds_current_row - 1 'start from the row above to ensure the first row it checks isn't it's own one
    
    Do Until odds_trend_row < 1
        
        'check gameref to see if it is the same
        If Sheets(odds_sheet).Cells(odds_trend_row, odds_gameref_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_gameref_column).Value Then
            
            'Current away odds are greater
            If Sheets(odds_sheet).Cells(odds_current_row, odds_awayodds_column).Value > Sheets(odds_sheet).Cells(odds_trend_row, odds_awayodds_column).Value Then
                Sheets(odds_sheet).Cells(odds_current_row, odds_awaytrend_column).Value = ChrW(&H2191)
                Sheets(odds_sheet).Cells(odds_current_row, odds_awaytrend_column).Interior.Color = RGB(0, 176, 80)
            
            ElseIf Sheets(odds_sheet).Cells(odds_current_row, odds_awayodds_column).Value < Sheets(odds_sheet).Cells(odds_trend_row, odds_awayodds_column).Value Then
                Sheets(odds_sheet).Cells(odds_current_row, odds_awaytrend_column).Value = ChrW(&H2193)
                Sheets(odds_sheet).Cells(odds_current_row, odds_awaytrend_column).Interior.Color = RGB(255, 0, 0)
                
            Else
                Sheets(odds_sheet).Cells(odds_current_row, odds_awaytrend_column).Value = ChrW(&H2194)
                Sheets(odds_sheet).Cells(odds_current_row, odds_awaytrend_column).Interior.Color = RGB(191, 191, 191)
            End If
            
            'home odds
            If Sheets(odds_sheet).Cells(odds_current_row, odds_homeodds_column).Value > Sheets(odds_sheet).Cells(odds_trend_row, odds_homeodds_column).Value Then
                Sheets(odds_sheet).Cells(odds_current_row, odds_hometrend_column).Value = ChrW(&H2191)
                Sheets(odds_sheet).Cells(odds_current_row, odds_hometrend_column).Interior.Color = RGB(0, 176, 80)
            
            ElseIf Sheets(odds_sheet).Cells(odds_current_row, odds_homeodds_column).Value < Sheets(odds_sheet).Cells(odds_trend_row, odds_homeodds_column).Value Then
                Sheets(odds_sheet).Cells(odds_current_row, odds_hometrend_column).Value = ChrW(&H2193)
                Sheets(odds_sheet).Cells(odds_current_row, odds_hometrend_column).Interior.Color = RGB(255, 0, 0)
                
            Else
                Sheets(odds_sheet).Cells(odds_current_row, odds_hometrend_column).Value = ChrW(&H2194)
                Sheets(odds_sheet).Cells(odds_current_row, odds_hometrend_column).Interior.Color = RGB(191, 191, 191)
            End If
        
            Exit Do
        Else
            Sheets(odds_sheet).Cells(odds_current_row, odds_awaytrend_column).Value = "N"
            Sheets(odds_sheet).Cells(odds_current_row, odds_hometrend_column).Value = "N"
        End If
        
        odds_trend_row = odds_trend_row - 1
    Loop
    
    
    'Select odds sheet to ensure the macro goes back to where it was started from
    Sheets(odds_sheet).Cells(odds_current_row + 1, odds_awayodds_column).Select

Next odds_current_row

GoTo Reset

Reset:
Call O_Switch_On_Functionality

End Sub

Sub C_Results()
Attribute C_Results.VB_ProcData.VB_Invoke_Func = "C\n14"

On Err GoTo Reset:

Call F_Switch_Off_Functionality

'Get current cell co-ordinates
Dim games_current_row As Long
games_current_row = ActiveCell.Row

'Sheet names
Dim elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet As String
Call VA_Worksheet_Names(elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet)

current_sheet = ActiveSheet.Name

If current_sheet = games_sheet Then
Else
    MsgBox games_sheet & " worksheet is not selected. This macro will now exit."
    GoTo Reset:
End If

'Games variables
Sheets(games_sheet).Select
Dim games_winner_column, games_gameref_column, games_ra_column, games_rh_column, games_rhh_column, games_ea_column, games_eh_column, games_year_column, games_week_column, games_yw_column, games_awayteam_column, games_hometeam_column, games_neutral_column, games_prediction_column, games_awayscore_column, games_homescore_column, games_winnerstatus_column, games_predcorrect_column, games_sa_column, games_sh_column, games_rd_column, games_radash_column, games_rhdash_column, games_predtotal_column As Integer
Call VG_Games_Sheet_Variables(games_winner_column, games_gameref_column, games_ra_column, games_rh_column, games_rhh_column, games_ea_column, games_eh_column, games_year_column, games_week_column, games_yw_column, games_awayteam_column, games_hometeam_column, games_neutral_column, games_prediction_column, games_awayscore_column, games_homescore_column, games_winnerstatus_column, games_predcorrect_column, games_sa_column, games_sh_column, games_rd_column, games_radash_column, games_rhdash_column, games_predtotal_column)

'Elo sheet variables
Sheets(elo_sheet).Select
Dim elo_ref_column, elo_year_column, elo_team_column, elo_yw_column, elo_week_column, elo_elo_column As Integer
Call VE_Elo_Sheet_Variables(elo_ref_column, elo_year_column, elo_team_column, elo_yw_column, elo_week_column, elo_elo_column)

Dim elo_start_row As Integer
elo_start_row = 2

Dim elo_final_row As Long
elo_final_row = WorksheetFunction.CountA(Columns(elo_yw_column))

'Odds variables
Sheets(odds_sheet).Select
Dim odds_gameref_column, odds_datetime_column, odds_awayodds_column, odds_homeodds_column, odds_awayelodec_column, odds_homeelodec_column, odds_awayoddsdec_column, odds_homeoddsdec_column, odds_awayedge_column, odds_homeedge_column, odds_betteam_column, odds_bethan_column, odds_kelly_column, odds_betamount_column, odds_winamount_column, odds_awayreqdec_column, odds_homereqdec_column, odds_awayreqodds_column, odds_homereqodds_column, odds_year_column, odds_week_column, odds_betplaced_column, odds_betodds_column, odds_betedge_column, odds_winlose_column, odds_opposeteam_column, odds_edgeedge_column, odds_betunits_column, odds_awaytrend_column, odds_hometrend_column, odds_neutral_column, odds_awayteam_column, odds_hometeam_column, odds_profitloss_column, odds_winnerstatus_column, odds_cumulative_column As Integer
Call VO_Odds_Sheet_Variables(odds_gameref_column, odds_datetime_column, odds_awayodds_column, odds_homeodds_column, odds_awayelodec_column, odds_homeelodec_column, odds_awayoddsdec_column, odds_homeoddsdec_column, odds_awayedge_column, odds_homeedge_column, odds_betteam_column, odds_bethan_column, odds_kelly_column, odds_betamount_column, odds_winamount_column, odds_awayreqdec_column, odds_homereqdec_column, odds_awayreqodds_column, odds_homereqodds_column, odds_year_column, odds_week_column, odds_betplaced_column, odds_betodds_column, odds_betedge_column, odds_winlose_column, odds_opposeteam_column, odds_edgeedge_column, odds_betunits_column, odds_awaytrend_column, odds_hometrend_column, odds_neutral_column, odds_awayteam_column, odds_hometeam_column, odds_profitloss_column, odds_winnerstatus_column, odds_cumulative_column)

'Error messages
Sheets(games_sheet).Select

If IsEmpty(Sheets(games_sheet).Cells(games_current_row, games_awayscore_column)) = True Then
    MsgBox "The AWAYSCORE is not populated against the current game. This macro will now exit."
    GoTo Reset:
Else
End If

If IsEmpty(Sheets(games_sheet).Cells(games_current_row, games_homescore_column)) = True Then
    MsgBox "The HOMESCORE is not populated against the current game. This macro will now exit."
    GoTo Reset:
Else
End If

If IsEmpty(Sheets(games_sheet).Cells(games_current_row, games_winner_column)) = False Then
    MsgBox "The WINNER is populated against the current game. This macro will now exit."
    GoTo Reset:
Else
End If

'Calculate winner and winner status
If Sheets(games_sheet).Cells(games_current_row, games_awayscore_column).Value = Sheets(games_sheet).Cells(games_current_row, games_homescore_column).Value Then
    
    Sheets(games_sheet).Cells(games_current_row, games_winner_column).Value = "TIE"
    Sheets(games_sheet).Cells(games_current_row, games_winnerstatus_column).Value = "TIE"
    
ElseIf Sheets(games_sheet).Cells(games_current_row, games_awayscore_column).Value > Sheets(games_sheet).Cells(games_current_row, games_homescore_column).Value Then

    Sheets(games_sheet).Cells(games_current_row, games_winner_column).Value = Sheets(games_sheet).Cells(games_current_row, games_awayteam_column).Value
    Sheets(games_sheet).Cells(games_current_row, games_winnerstatus_column).Value = "AWAY"
    
ElseIf Sheets(games_sheet).Cells(games_current_row, games_awayscore_column).Value < Sheets(games_sheet).Cells(games_current_row, games_homescore_column).Value Then

    Sheets(games_sheet).Cells(games_current_row, games_winner_column).Value = Sheets(games_sheet).Cells(games_current_row, games_hometeam_column).Value
    Sheets(games_sheet).Cells(games_current_row, games_winnerstatus_column).Value = "HOME"
    
Else

    Sheets(games_sheet).Cells(games_current_row, games_winner_column).Value = "ERROR"
    Sheets(games_sheet).Cells(games_current_row, games_winnerstatus_column).Value = "ERROR"

End If

'Calculate prediction correct
If Sheets(games_sheet).Cells(games_current_row, games_winner_column).Value = Sheets(games_sheet).Cells(games_current_row, games_prediction_column).Value Then

    Sheets(games_sheet).Cells(games_current_row, games_predcorrect_column).Value = "1"
    
Else

    Sheets(games_sheet).Cells(games_current_row, games_predcorrect_column).Value = "0"

End If

Sheets(games_sheet).Cells(games_current_row, games_predtotal_column).Value = "1"

'Calculate SA and SH
If Sheets(games_sheet).Cells(games_current_row, games_winnerstatus_column).Value = "TIE" Then

    Sheets(games_sheet).Cells(games_current_row, games_sa_column).Value = 0.5
    Sheets(games_sheet).Cells(games_current_row, games_sh_column).Value = 0.5
    
ElseIf Sheets(games_sheet).Cells(games_current_row, games_winnerstatus_column).Value = "AWAY" Then

    Sheets(games_sheet).Cells(games_current_row, games_sa_column).Value = 1
    Sheets(games_sheet).Cells(games_current_row, games_sh_column).Value = 0
    
ElseIf Sheets(games_sheet).Cells(games_current_row, games_winnerstatus_column).Value = "HOME" Then

    Sheets(games_sheet).Cells(games_current_row, games_sa_column).Value = 0
    Sheets(games_sheet).Cells(games_current_row, games_sh_column).Value = 1

Else

    Sheets(games_sheet).Cells(games_current_row, games_sa_column).Value = "ERROR"
    Sheets(games_sheet).Cells(games_current_row, games_sh_column).Value = "ERROR"

End If

'Calculate R delta
Dim K As Double
K = Range("K").Value

Sheets(games_sheet).Cells(games_current_row, games_rd_column).Value = Abs(K * (Sheets(games_sheet).Cells(games_current_row, games_sa_column).Value - Sheets(games_sheet).Cells(games_current_row, games_ea_column).Value))

'Calculate R away dash
Sheets(games_sheet).Cells(games_current_row, games_radash_column).Value = Sheets(games_sheet).Cells(games_current_row, games_ra_column).Value + (K * (Sheets(games_sheet).Cells(games_current_row, games_sa_column).Value - Sheets(games_sheet).Cells(games_current_row, games_ea_column).Value))

'Calculate R home dash
Sheets(games_sheet).Cells(games_current_row, games_rhdash_column).Value = Sheets(games_sheet).Cells(games_current_row, games_rh_column).Value + (K * (Sheets(games_sheet).Cells(games_current_row, games_sh_column).Value - Sheets(games_sheet).Cells(games_current_row, games_eh_column).Value))

'Determine Elo strings to populate
Dim eloref_week, eloref_year As Integer
eloref_week = Sheets(games_sheet).Cells(games_current_row, games_week_column).Value + 1
eloref_year = Sheets(games_sheet).Cells(games_current_row, games_year_column).Value

Dim away_eloref As String
away_eloref = "E" & eloref_year & eloref_week & Sheets(games_sheet).Cells(games_current_row, games_awayteam_column).Value

Dim home_eloref As String
home_eloref = "E" & eloref_year & eloref_week & Sheets(games_sheet).Cells(games_current_row, games_hometeam_column).Value

Sheets(elo_sheet).Cells(elo_final_row + 1, elo_ref_column).Value = away_eloref
Sheets(elo_sheet).Cells(elo_final_row + 2, elo_ref_column).Value = home_eloref
Sheets(elo_sheet).Cells(elo_final_row + 1, elo_year_column).Value = eloref_year
Sheets(elo_sheet).Cells(elo_final_row + 2, elo_year_column).Value = eloref_year
Sheets(elo_sheet).Cells(elo_final_row + 1, elo_week_column).Value = eloref_week
Sheets(elo_sheet).Cells(elo_final_row + 2, elo_week_column).Value = eloref_week

If eloref_week < 10 Then
    Sheets(elo_sheet).Cells(elo_final_row + 1, elo_yw_column).Value = eloref_year & "-0" & eloref_week
    Sheets(elo_sheet).Cells(elo_final_row + 2, elo_yw_column).Value = eloref_year & "-0" & eloref_week
Else
    Sheets(elo_sheet).Cells(elo_final_row + 1, elo_yw_column).Value = eloref_year & "-" & eloref_week
    Sheets(elo_sheet).Cells(elo_final_row + 2, elo_yw_column).Value = eloref_year & "-" & eloref_week
End If
    
Sheets(elo_sheet).Cells(elo_final_row + 1, elo_team_column).Value = Sheets(games_sheet).Cells(games_current_row, games_awayteam_column).Value
Sheets(elo_sheet).Cells(elo_final_row + 2, elo_team_column).Value = Sheets(games_sheet).Cells(games_current_row, games_hometeam_column).Value
Sheets(elo_sheet).Cells(elo_final_row + 1, elo_elo_column).Value = Sheets(games_sheet).Cells(games_current_row, games_radash_column).Value
Sheets(elo_sheet).Cells(elo_final_row + 2, elo_elo_column).Value = Sheets(games_sheet).Cells(games_current_row, games_rhdash_column).Value

Dim gameref As String
gameref = Sheets(games_sheet).Cells(games_current_row, games_gameref_column).Value

Sheets(odds_sheet).Select
Dim odds_current_row, odds_start_row, odds_final_row As Integer
odds_final_row = WorksheetFunction.CountA(Columns(odds_gameref_column))
odds_start_row = 2

'Search around odds sheet to find a placed bet
For odds_current_row = odds_start_row To odds_final_row

    If Sheets(odds_sheet).Cells(odds_current_row, odds_gameref_column).Value = gameref Then
    
        'winnerstatus
        Sheets(odds_sheet).Cells(odds_current_row, odds_winnerstatus_column).Value = Sheets(games_sheet).Cells(games_current_row, games_winnerstatus_column).Value
    
        'grey out cells
        Sheets(odds_sheet).Select
        Range(Cells(odds_current_row, 1), Cells(odds_current_row, 36)).Select
        With Selection.Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.249977111117893
            .PatternTintAndShade = 0
        End With
        Cells(odds_current_row, 1).Select
    
        If Sheets(odds_sheet).Cells(odds_current_row, odds_betplaced_column).Value = "Y" Then
        
            'Determine win/loss amount
            If Sheets(odds_sheet).Cells(odds_current_row, odds_betteam_column).Value = Sheets(games_sheet).Cells(games_current_row, games_winner_column).Value Then
            
                Sheets(odds_sheet).Cells(odds_current_row, odds_profitloss_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_winamount_column).Value
                Sheets(odds_sheet).Cells(odds_current_row, odds_winlose_column).Value = "WIN"
                'Sheets(odds_sheet).Cells(odds_current_row, odds_cumulative_column).Value = Sheets(odds_sheet).Cells(odds_current_row - 1, odds_cumulative_column).Value + Sheets(odds_sheet).Cells(odds_current_row, odds_profitloss_column).Value
                
            'Tie tie tie means a 'push' so zero money gained, zero money lost
            ElseIf Sheets(games_sheet).Cells(games_current_row, games_winner_column).Value = "TIE" Then
            
                Sheets(odds_sheet).Cells(odds_current_row, odds_profitloss_column).Value = 0
                Sheets(odds_sheet).Cells(odds_current_row, odds_winlose_column).Value = "PUSH"
                'Sheets(odds_sheet).Cells(odds_current_row, odds_cumulative_column).Value = Sheets(odds_sheet).Cells(odds_current_row - 1, odds_cumulative_column).Value
            
            Else
            
                Sheets(odds_sheet).Cells(odds_current_row, odds_profitloss_column).Value = -Sheets(odds_sheet).Cells(odds_current_row, odds_betamount_column).Value
                Sheets(odds_sheet).Cells(odds_current_row, odds_winlose_column).Value = "LOSE"
                'Sheets(odds_sheet).Cells(odds_current_row, odds_cumulative_column).Value = Sheets(odds_sheet).Cells(odds_current_row - 1, odds_cumulative_column).Value + Sheets(odds_sheet).Cells(odds_current_row, odds_profitloss_column).Value
            
            End If
            
        Else
        
            'set profit/loss to 0 as no bet has been placed
            Sheets(odds_sheet).Cells(odds_current_row, odds_profitloss_column).Value = 0
            Sheets(odds_sheet).Cells(odds_current_row, odds_betplaced_column).Value = "N"
            Sheets(odds_sheet).Cells(odds_current_row, odds_winlose_column).Value = "N/A"
            'Sheets(odds_sheet).Cells(odds_current_row, odds_cumulative_column).Value = Sheets(odds_sheet).Cells(odds_current_row - 1, odds_cumulative_column).Value
        
        End If
    
    Else
    
    End If

Next odds_current_row

'grey out cells
Sheets(games_sheet).Select
Range(Cells(games_current_row, 1), Cells(games_current_row, 24)).Select
With Selection.Interior
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorDark1
    .TintAndShade = -0.249977111117893
    .PatternTintAndShade = 0
End With

GoTo Reset:

Reset:
Call O_Switch_On_Functionality

Sheets(games_sheet).Cells(games_current_row + 1, 14).Select

End Sub
Sub D_Odds_Graph()
Attribute D_Odds_Graph.VB_ProcData.VB_Invoke_Func = "D\n14"
'Macro to create a graph to plot the movement of odds for a single game

Call F_Switch_Off_Functionality

'Sheet names
Dim elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet As String
Call VA_Worksheet_Names(elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet)

'determine if the graph worksheet is actually selected
current_sheet = ActiveSheet.Name

If current_sheet = graph_sheet Then
Else
    MsgBox games_sheet & " worksheet is not selected. This macro will now exit."
    Exit Sub
End If

'graph sheet variables
Sheets(graph_sheet).Select
Dim graph_gameref_row, graph_gameref_column, graph_datetime_column, graph_awayelodec_column, graph_homeelodec_column, graph_awayoddsdec_column, graph_homeoddsdec_column, graph_awayreqdec_column, graph_homereqdec_column, graph_awayplaced_column, graph_homeplaced_column, graph_awayprospect_column, graph_homeprospect_column, graph_week_row, graph_week_column As Integer
Call VH_Graph_Sheet_Variables(graph_gameref_row, graph_gameref_column, graph_datetime_column, graph_awayelodec_column, graph_homeelodec_column, graph_awayoddsdec_column, graph_homeoddsdec_column, graph_awayreqdec_column, graph_homereqdec_column, graph_awayplaced_column, graph_homeplaced_column, graph_awayprospect_column, graph_homeprospect_column, graph_week_row, graph_week_column)

Dim gameref As String
gameref = Sheets(graph_sheet).Cells(graph_gameref_row, graph_gameref_column + 1).Value

Dim graph_start_row As Integer
graph_start_row = 2

Dim graph_current_row As Integer
graph_current_row = graph_start_row

'Odds variables
Sheets(odds_sheet).Select
Dim odds_gameref_column, odds_datetime_column, odds_awayodds_column, odds_homeodds_column, odds_awayelodec_column, odds_homeelodec_column, odds_awayoddsdec_column, odds_homeoddsdec_column, odds_awayedge_column, odds_homeedge_column, odds_betteam_column, odds_bethan_column, odds_kelly_column, odds_betamount_column, odds_winamount_column, odds_awayreqdec_column, odds_homereqdec_column, odds_awayreqodds_column, odds_homereqodds_column, odds_year_column, odds_week_column, odds_betplaced_column, odds_betodds_column, odds_betedge_column, odds_winlose_column, odds_opposeteam_column, odds_edgeedge_column, odds_betunits_column, odds_awaytrend_column, odds_hometrend_column, odds_neutral_column, odds_awayteam_column, odds_hometeam_column, odds_profitloss_column, odds_winnerstatus_column, odds_cumulative_column As Integer
Call VO_Odds_Sheet_Variables(odds_gameref_column, odds_datetime_column, odds_awayodds_column, odds_homeodds_column, odds_awayelodec_column, odds_homeelodec_column, odds_awayoddsdec_column, odds_homeoddsdec_column, odds_awayedge_column, odds_homeedge_column, odds_betteam_column, odds_bethan_column, odds_kelly_column, odds_betamount_column, odds_winamount_column, odds_awayreqdec_column, odds_homereqdec_column, odds_awayreqodds_column, odds_homereqodds_column, odds_year_column, odds_week_column, odds_betplaced_column, odds_betodds_column, odds_betedge_column, odds_winlose_column, odds_opposeteam_column, odds_edgeedge_column, odds_betunits_column, odds_awaytrend_column, odds_hometrend_column, odds_neutral_column, odds_awayteam_column, odds_hometeam_column, odds_profitloss_column, odds_winnerstatus_column, odds_cumulative_column)

Dim odds_final_row As Integer
odds_final_row = WorksheetFunction.CountA(Columns(odds_gameref_column))

Dim odds_start_row As Integer
odds_start_row = 2

Dim odds_current_row As Integer
odds_current_row = odds_start_row

Dim gameref_found As Boolean
gameref_found = False

'loop around rows in odds sheet
For odds_current_row = odds_start_row To odds_final_row

    If Sheets(odds_sheet).Cells(odds_current_row, odds_gameref_column).Value = gameref Then
        gameref_found = True
        Sheets(graph_sheet).Cells(graph_current_row, graph_datetime_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_datetime_column).Value
        Sheets(graph_sheet).Cells(graph_current_row, graph_awayoddsdec_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_awayoddsdec_column).Value
        Sheets(graph_sheet).Cells(graph_current_row, graph_homeoddsdec_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_homeoddsdec_column).Value
        Sheets(graph_sheet).Cells(graph_current_row, graph_awayreqdec_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_awayreqdec_column).Value
        Sheets(graph_sheet).Cells(graph_current_row, graph_homereqdec_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_homereqdec_column).Value
        Sheets(graph_sheet).Cells(graph_current_row, graph_awayelodec_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_awayelodec_column).Value
        Sheets(graph_sheet).Cells(graph_current_row, graph_homeelodec_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_homeelodec_column).Value
        
        Sheets(graph_sheet).Cells(graph_week_row, graph_week_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_week_column).Value
        Sheets(graph_sheet).Cells(graph_week_row + 1, graph_week_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_awayelodec_column).Value
        Sheets(graph_sheet).Cells(graph_week_row + 2, graph_week_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_homeelodec_column).Value
        Sheets(graph_sheet).Cells(graph_week_row + 3, graph_week_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_awayoddsdec_column).Value
        Sheets(graph_sheet).Cells(graph_week_row + 4, graph_week_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_homeoddsdec_column).Value
        Sheets(graph_sheet).Cells(graph_week_row + 5, graph_week_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_awayedge_column).Value
        Sheets(graph_sheet).Cells(graph_week_row + 6, graph_week_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_homeedge_column).Value
        Sheets(graph_sheet).Cells(graph_week_row + 7, graph_week_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_neutral_column).Value
                       
        'bet amounts and bet placed rows
        If Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "NONE" Then
            Sheets(graph_sheet).Cells(graph_current_row, graph_awayprospect_column).Value = ""
            Sheets(graph_sheet).Cells(graph_current_row, graph_homeprospect_column).Value = ""
            Sheets(graph_sheet).Cells(graph_current_row, graph_awayplaced_column).Value = ""
            Sheets(graph_sheet).Cells(graph_current_row, graph_homeplaced_column).Value = ""
        
        ElseIf Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "AWAY" Then
            Sheets(graph_sheet).Cells(graph_current_row, graph_awayprospect_column).Value = 1
            Sheets(graph_sheet).Cells(graph_current_row, graph_homeprospect_column).Value = ""
            
            If Sheets(odds_sheet).Cells(odds_current_row, odds_betplaced_column).Value = "Y" Then
                Sheets(graph_sheet).Cells(graph_current_row, graph_awayplaced_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_awayoddsdec_column).Value
                Sheets(graph_sheet).Cells(graph_current_row, graph_homeplaced_column).Value = ""
            Else
                Sheets(graph_sheet).Cells(graph_current_row, graph_awayplaced_column).Value = ""
                Sheets(graph_sheet).Cells(graph_current_row, graph_homeplaced_column).Value = ""
            End If
        
        ElseIf Sheets(odds_sheet).Cells(odds_current_row, odds_bethan_column).Value = "HOME" Then
            Sheets(graph_sheet).Cells(graph_current_row, graph_awayprospect_column).Value = ""
            Sheets(graph_sheet).Cells(graph_current_row, graph_homeprospect_column).Value = 1
            
            If Sheets(odds_sheet).Cells(odds_current_row, odds_betplaced_column).Value = "Y" Then
                Sheets(graph_sheet).Cells(graph_current_row, graph_awayplaced_column).Value = ""
                Sheets(graph_sheet).Cells(graph_current_row, graph_homeplaced_column).Value = Sheets(odds_sheet).Cells(odds_current_row, odds_homeoddsdec_column).Value
            Else
                Sheets(graph_sheet).Cells(graph_current_row, graph_awayplaced_column).Value = ""
                Sheets(graph_sheet).Cells(graph_current_row, graph_homeplaced_column).Value = ""
            End If
        Else
        
        End If
        
        graph_current_row = graph_current_row + 1
        
    Else
    
    End If
    

Next odds_current_row

'If the game ref has not been found, the user will be informed
If gameref_found = False Then

    MsgBox gameref & " has not been found in worksheet " & odds_sheet & ". This macro will now exit."
    Call O_Switch_On_Functionality
    End

Else

End If

'delete old content
Sheets(graph_sheet).Select
Range(Cells(graph_current_row, graph_datetime_column), Cells(51, graph_homeplaced_column)).Select
Selection.ClearContents

Sheets(graph_sheet).Select
Cells(1, 1).Select

Call O_Switch_On_Functionality

End Sub

Sub E_Year_End()
Attribute E_Year_End.VB_ProcData.VB_Invoke_Func = "E\n14"

current_sheet = ActiveSheet.Name

'Sheet names
Dim elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet As String
Call VA_Worksheet_Names(elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet)

'Check that the elo sheet is selected
If current_sheet = elo_sheet Then
Else
    MsgBox elo_sheet & " worksheet is not selected. This macro will now exit."
    Exit Sub
End If

'Elo sheet variables
Sheets(elo_sheet).Select
Dim elo_ref_column, elo_year_column, elo_team_column, elo_yw_column, elo_week_column, elo_elo_column As Integer
Call VE_Elo_Sheet_Variables(elo_ref_column, elo_year_column, elo_team_column, elo_yw_column, elo_week_column, elo_elo_column)

Dim elo_start_row As Integer
elo_start_row = 2

Dim elo_final_row As Long
elo_final_row = WorksheetFunction.CountA(Columns(elo_yw_column))

Dim elo_current_row As Long
elo_current_row = elo_start_row

'Teams variables
Sheets(teams_sheet).Select
Dim teams_teamref_column, teams_name_column, teams_fullname_column, teams_fbsyear_column, teams_nickname_column, teams_conference_column As Integer
Call VT_Teams_Sheet_Variables(teams_teamref_column, teams_name_column, teams_fullname_column, teams_fbsyear_column, teams_nickname_column, teams_conference_column)

Dim teams_current_row, teams_start_row, teams_final_row As Integer
teams_start_row = 2

teams_final_row = WorksheetFunction.CountA(Columns(teams_name_column))

Dim teams_current_team As String

'FBS variable
Dim FBS As Double
FBS = Range("FBS").Value

'FCS variable
Dim FCS As Double
FCS = Range("FCS").Value

'Find Division variables
Dim isFBS As Boolean

'Go to end of the data
elo_current_row = elo_final_row + 1

'Determine new year
Dim year, new_year As Integer

'find last year in worksheet
year = Sheets(elo_sheet).Cells(elo_final_row, elo_year_column).Value

new_year = year + 1

Dim RTM As Double
RTM = Range("RTM").Value

Dim current_elo, modified_elo As Double

'set up variables for the teams row - checking back to the start of the list of values
Dim elo_team_row, elo_findr_row As Long
elo_team_row = elo_start_row

Dim teams_name As String

'Loop around all teams - number taken from recent count
Do While elo_team_row <= teams_final_row

    'Populate year (+1) and week (0)
    Sheets(elo_sheet).Cells(elo_current_row, elo_year_column).Value = new_year
    Sheets(elo_sheet).Cells(elo_current_row, elo_week_column).Value = 0
    Sheets(elo_sheet).Cells(elo_current_row, elo_yw_column).Value = new_year & "-00"
        
    'Populate team name
    teams_name = Sheets(elo_sheet).Cells(elo_team_row, elo_team_column).Value
    Sheets(elo_sheet).Cells(elo_current_row, elo_team_column).Value = teams_name
    
    'Populate EloRef
    Sheets(elo_sheet).Cells(elo_current_row, elo_ref_column).Value = "E" & new_year & "0" & teams_name
    
    'Find last team Elo - steal from macro A
    elo_findr_row = elo_final_row

    Do Until elo_findr_row < elo_start_row

        If Sheets(elo_sheet).Cells(elo_findr_row, elo_team_column).Value = teams_name Then
            current_elo = Sheets(elo_sheet).Cells(elo_findr_row, elo_elo_column).Value
            Exit Do
        Else
        
        End If
    
        elo_findr_row = elo_findr_row - 1
        
    Loop
    
    'Perform RTM correction
    'Determine if the team is FBS or FCS
    isFBS = False
        
    Call U_Find_Division(teams_name, isFBS, new_year)
    
    If isFBS = True Then
        modified_elo = current_elo + ((FBS - current_elo) * RTM)
    Else
        modified_elo = current_elo + ((FCS - current_elo) * RTM)
    End If
    
    'Populate adjusted Elo
    Sheets(elo_sheet).Cells(elo_current_row, elo_elo_column).Value = modified_elo
    
    'End loop
    elo_current_row = elo_current_row + 1
    elo_team_row = elo_team_row + 1
Loop

End Sub

Sub F_Switch_Off_Functionality()
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
End Sub

Sub O_Switch_On_Functionality()
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True
End Sub
Sub G_Bets_Graph()

'Sheet names
Dim elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet As String
Call VA_Worksheet_Names(elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet)

'Odds variables
Sheets(odds_sheet).Select
Dim odds_gameref_column, odds_datetime_column, odds_awayodds_column, odds_homeodds_column, odds_awayelodec_column, odds_homeelodec_column, odds_awayoddsdec_column, odds_homeoddsdec_column, odds_awayedge_column, odds_homeedge_column, odds_betteam_column, odds_bethan_column, odds_kelly_column, odds_betamount_column, odds_winamount_column, odds_awayreqdec_column, odds_homereqdec_column, odds_awayreqodds_column, odds_homereqodds_column, odds_year_column, odds_week_column, odds_betplaced_column, odds_betodds_column, odds_betedge_column, odds_winlose_column, odds_opposeteam_column, odds_edgeedge_column, odds_betunits_column, odds_awaytrend_column, odds_hometrend_column, odds_neutral_column, odds_awayteam_column, odds_hometeam_column, odds_profitloss_column, odds_winnerstatus_column, odds_cumulative_column As Integer
Call VO_Odds_Sheet_Variables(odds_gameref_column, odds_datetime_column, odds_awayodds_column, odds_homeodds_column, odds_awayelodec_column, odds_homeelodec_column, odds_awayoddsdec_column, odds_homeoddsdec_column, odds_awayedge_column, odds_homeedge_column, odds_betteam_column, odds_bethan_column, odds_kelly_column, odds_betamount_column, odds_winamount_column, odds_awayreqdec_column, odds_homereqdec_column, odds_awayreqodds_column, odds_homereqodds_column, odds_year_column, odds_week_column, odds_betplaced_column, odds_betodds_column, odds_betedge_column, odds_winlose_column, odds_opposeteam_column, odds_edgeedge_column, odds_betunits_column, odds_awaytrend_column, odds_hometrend_column, odds_neutral_column, odds_awayteam_column, odds_hometeam_column, odds_profitloss_column, odds_winnerstatus_column, odds_cumulative_column)

Dim betsgraph_sheet As String
betsgraph_sheet = "Bc"
Sheets(betsgraph_sheet).Select

Call VB_Bets_Graph_Variables


End Sub

Sub P_Predictions()
Attribute P_Predictions.VB_ProcData.VB_Invoke_Func = "P\n14"
'Generate a set of predictions, including an output txt file for inputting into blog

Call F_Switch_Off_Functionality

current_sheet = ActiveSheet.Name

'Sheet names
Dim elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet As String
Call VA_Worksheet_Names(elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet)

'Check that the prediction (Pd) sheet is selected
If current_sheet = pred_sheet Then
Else
    MsgBox pred_sheet & " worksheet is not selected. This macro will now exit."
    Exit Sub
End If

'Games variables
Sheets(games_sheet).Select
Dim games_winner_column, games_gameref_column, games_ra_column, games_rh_column, games_rhh_column, games_ea_column, games_eh_column, games_year_column, games_week_column, games_yw_column, games_awayteam_column, games_hometeam_column, games_neutral_column, games_prediction_column, games_awayscore_column, games_homescore_column, games_winnerstatus_column, games_predcorrect_column, games_sa_column, games_sh_column, games_rd_column, games_radash_column, games_rhdash_column, games_predtotal_column As Integer
Call VG_Games_Sheet_Variables(games_winner_column, games_gameref_column, games_ra_column, games_rh_column, games_rhh_column, games_ea_column, games_eh_column, games_year_column, games_week_column, games_yw_column, games_awayteam_column, games_hometeam_column, games_neutral_column, games_prediction_column, games_awayscore_column, games_homescore_column, games_winnerstatus_column, games_predcorrect_column, games_sa_column, games_sh_column, games_rd_column, games_radash_column, games_rhdash_column, games_predtotal_column)

'Predictions Sheet variables
Sheets(pred_sheet).Select
Dim pred_gameref_column, pred_awayteam_column, pred_hometeam_column, pred_at_column
Call VP_Predictions_Sheet_Variables(pred_gameref_column, pred_awayteam_column, pred_hometeam_column, pred_at_column)

Dim pred_start_row, pred_current_row As Integer
pred_start_row = 2
pred_current_row = pred_start_row

Dim away_team_short, home_team_short, away_team_full, home_team_full, prediction_team As String

Sheets(games_sheet).Select
Dim games_final_row, games_current_row, games_start_row As Long
games_final_row = WorksheetFunction.CountA(Columns(games_gameref_column))

games_start_row = games_final_row
games_current_row = games_start_row

Dim games_week As Integer
games_week = Sheets(games_sheet).Cells(games_final_row, games_week_column).Value

'Find no_games
Dim no_games As Integer
no_games = 0

Do While Sheets(games_sheet).Cells(games_current_row, games_week_column).Value = games_week

    no_games = no_games + 1
    games_current_row = games_current_row - 1
Loop

'Test message
'MsgBox "The number of games in this week is: " & no_games

games_start_row = games_current_row + 1

For games_current_row = games_start_row To games_final_row

    Sheets(pred_sheet).Cells(pred_current_row, pred_gameref_column).Value = Sheets(games_sheet).Cells(games_current_row, games_gameref_column).Value
    away_team_short = Sheets(games_sheet).Cells(games_current_row, games_awayteam_column).Value
    home_team_short = Sheets(games_sheet).Cells(games_current_row, games_hometeam_column).Value
    prediction_team = Sheets(games_sheet).Cells(games_current_row, games_prediction_column).Value
    
    teams_name = away_team_short
    Call T_Teams_Fullname(teams_name, teams_fullname)
    away_team_full = teams_fullname
    
    teams_name = home_team_short
    Call T_Teams_Fullname(teams_name, teams_fullname)
    home_team_full = teams_fullname
    
    If prediction_team = away_team_short Then
        Sheets(pred_sheet).Cells(pred_current_row, pred_awayteam_column).Value = "<b>[" & away_team_full & "]</b>"
        Sheets(pred_sheet).Cells(pred_current_row, pred_hometeam_column).Value = home_team_full & "<br />"
    ElseIf prediction_team = home_team_short Then
        Sheets(pred_sheet).Cells(pred_current_row, pred_awayteam_column).Value = away_team_full
        Sheets(pred_sheet).Cells(pred_current_row, pred_hometeam_column).Value = "<b>[" & home_team_full & "]</b><br />"
    Else
        MsgBox "Cannot resolve the predicted winner for: " & Sheets(games_sheet).Cells(games_current_row, games_gameref_column).Value & ". This macro will now exit."
        Exit Sub
    End If
    
    If Sheets(games_sheet).Cells(games_current_row, games_neutral_column).Value = "N" Then
    
        Sheets(pred_sheet).Cells(pred_current_row, pred_at_column).Value = "at"
    
    Else
    
        Sheets(pred_sheet).Cells(pred_current_row, pred_at_column).Value = "vs"
    
    End If
        
    pred_current_row = pred_current_row + 1

Next games_current_row

'delete old content
Sheets(pred_sheet).Select
Range(Cells(pred_current_row, pred_gameref_column), Cells(100, pred_hometeam_column)).Select
Selection.ClearContents

Cells(1, 1).Select

Call O_Switch_On_Functionality

End Sub

Sub R_Power_Rankings()
Attribute R_Power_Rankings.VB_ProcData.VB_Invoke_Func = "R\n14"

Call F_Switch_Off_Functionality

current_sheet = ActiveSheet.Name

'Sheet names
Dim elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet As String
Call VA_Worksheet_Names(elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet)

'Check that the power ranking sheet is selected
If current_sheet = power_sheet Then
Else
    MsgBox power_sheet & " worksheet is not selected. This macro will now exit."
    Call O_Switch_On_Functionality
    Exit Sub
End If

'Games variables
Sheets(games_sheet).Select
Dim games_winner_column, games_gameref_column, games_ra_column, games_rh_column, games_rhh_column, games_ea_column, games_eh_column, games_year_column, games_week_column, games_yw_column, games_awayteam_column, games_hometeam_column, games_neutral_column, games_prediction_column, games_awayscore_column, games_homescore_column, games_winnerstatus_column, games_predcorrect_column, games_sa_column, games_sh_column, games_rd_column, games_radash_column, games_rhdash_column, games_predtotal_column As Integer
Call VG_Games_Sheet_Variables(games_winner_column, games_gameref_column, games_ra_column, games_rh_column, games_rhh_column, games_ea_column, games_eh_column, games_year_column, games_week_column, games_yw_column, games_awayteam_column, games_hometeam_column, games_neutral_column, games_prediction_column, games_awayscore_column, games_homescore_column, games_winnerstatus_column, games_predcorrect_column, games_sa_column, games_sh_column, games_rd_column, games_radash_column, games_rhdash_column, games_predtotal_column)
Dim games_final_row As Integer
games_final_row = WorksheetFunction.CountA(Columns(games_awayteam_column))

'PR sheet variables
Sheets(power_sheet).Select
Dim power_rank_column, power_elo_column, power_team_column, power_wins_column, power_losses_column, power_ties_column As Integer
Call VR_Rankings_Sheet_Variables(power_rank_column, power_elo_column, power_team_column, power_wins_column, power_losses_column, power_ties_column)

Dim power_start_row, power_final_row, power_current_row As Integer
power_start_row = 2
power_current_row = power_start_row

'Elo sheet variables
Sheets(elo_sheet).Select
Dim elo_ref_column, elo_year_column, elo_team_column, elo_yw_column, elo_week_column, elo_elo_column As Integer
Call VE_Elo_Sheet_Variables(elo_ref_column, elo_year_column, elo_team_column, elo_yw_column, elo_week_column, elo_elo_column)

Dim elo_start_row As Integer
elo_start_row = 2

Dim elo_final_row As Long
elo_final_row = WorksheetFunction.CountA(Columns(elo_ref_column))

Dim elo_current_row As Long
elo_current_row = elo_start_row

'Teams variables
Sheets(teams_sheet).Select
Dim teams_teamref_column, teams_name_column, teams_fullname_column, teams_fbsyear_column, teams_nickname_column, teams_conference_column As Integer
Call VT_Teams_Sheet_Variables(teams_teamref_column, teams_name_column, teams_fullname_column, teams_fbsyear_column, teams_nickname_column, teams_conference_column)

Dim no_teams As Integer
no_teams = WorksheetFunction.CountA(Columns(teams_teamref_column)) - 1

Dim current_elo As Double

Dim team_name As String

'Teams_Fullname variables
'Call Teams_Fullname(teams_name, teams_fullname)
Dim teams_name, teams_fullname As String

Dim no_wins, no_draws, no_losses, games_current_row As Integer

elo_current_row = elo_start_row

'Loop around all teams - number taken from recent count
Do Until elo_current_row = elo_start_row + no_teams

    'populate rank
    Sheets(power_sheet).Cells(power_current_row, power_rank_column).Value = power_current_row - power_start_row + 1

    'Populate team name
    teams_name = Sheets(elo_sheet).Cells(elo_current_row, elo_team_column).Value
    
    'Find teams_fullname
    Call T_Teams_Fullname(teams_name, teams_fullname)
    
    Sheets(power_sheet).Cells(power_current_row, power_team_column).Value = teams_fullname
    
    'Find last team Elo - steal from macro A
    elo_findr_row = elo_final_row

    Do Until elo_findr_row < elo_start_row

        If Sheets(elo_sheet).Cells(elo_findr_row, elo_team_column).Value = teams_name Then
            current_elo = Sheets(elo_sheet).Cells(elo_findr_row, elo_elo_column).Value
            Exit Do
        Else
        
        End If
    
        elo_findr_row = elo_findr_row - 1
        
    Loop

    'Populate Elo
    Sheets(power_sheet).Cells(power_current_row, power_elo_column).Value = current_elo
    
    'determine number of wins, losses and draws
    no_wins = 0
    no_losses = 0
    no_draws = 0
    
    For games_current_row = 1 To games_final_row
        If Sheets(games_sheet).Cells(games_current_row, games_winner_column).Value <> "" Then
            If Sheets(games_sheet).Cells(games_current_row, games_awayteam_column).Value = teams_name Or Sheets(games_sheet).Cells(games_current_row, games_hometeam_column).Value = teams_name Then
                If Sheets(games_sheet).Cells(games_current_row, games_winner_column).Value = teams_name Then
                    no_wins = no_wins + 1
                ElseIf Sheets(games_sheet).Cells(games_current_row, games_winner_column).Value = "TIE" Then
                    no_draws = no_draws + 1
                Else
                    no_losses = no_losses + 1
                End If
            Else
            End If
        Else
        End If
    Next games_current_row
    
    'Populate Wins, Losses, Ties
    Sheets(power_sheet).Cells(power_current_row, power_wins_column).Value = no_wins
    Sheets(power_sheet).Cells(power_current_row, power_losses_column).Value = no_losses
    Sheets(power_sheet).Cells(power_current_row, power_ties_column).Value = no_draws
    
    'End loop
    elo_current_row = elo_current_row + 1
    power_current_row = power_current_row + 1
Loop

Sheets(power_sheet).Select

With ActiveSheet.Sort
     .SortFields.Add Key:=Range(Cells(power_start_row - 1, power_elo_column).Address()), Order:=xlDescending
     .SortFields.Add Key:=Range(Cells(power_start_row - 1, power_team_column).Address()), Order:=xlAscending
     .SetRange Range(Cells(power_start_row - 1, power_team_column).Address(), Cells(power_current_row, power_ties_column).Address())
     .Header = xlYes
     .Apply
End With

Sheets(power_sheet).Sort.SortFields.Clear

Call O_Switch_On_Functionality

End Sub

Sub S_Search(ByRef search_term, ByRef search_row, ByRef search_column)
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

Call O_Switch_On_Functionality

End

End Sub

Sub T_Teams_Fullname(ByRef teams_name, ByRef teams_fullname)
'Search function to find the full name of a team

'Sheet names
Dim elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet As String
Call VA_Worksheet_Names(elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet)

'Teams variables
Sheets(teams_sheet).Select
Dim teams_teamref_column, teams_name_column, teams_fullname_column, teams_fbsyear_column, teams_nickname_column, teams_conference_column As Integer
Call VT_Teams_Sheet_Variables(teams_teamref_column, teams_name_column, teams_fullname_column, teams_fbsyear_column, teams_nickname_column, teams_conference_column)

Dim teams_current_row, teams_start_row, teams_final_row As Integer
teams_start_row = 2

teams_final_row = WorksheetFunction.CountA(Columns(teams_name_column))

For teams_current_row = teams_start_row To teams_final_row
    If Sheets(teams_sheet).Cells(teams_current_row, teams_name_column).Value = teams_name Then
                    
        teams_fullname = Sheets(teams_sheet).Cells(teams_current_row, teams_fullname_column).Value
        Exit Sub

    Else
        
    End If
    
Next teams_current_row

End Sub

Sub U_Find_Division(ByRef teams_name, ByRef isFBS, ByRef new_year)

'Sheet names
Dim elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet As String
Call VA_Worksheet_Names(elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet)

'Teams variables
Sheets(teams_sheet).Select
Dim teams_teamref_column, teams_name_column, teams_fullname_column, teams_fbsyear_column, teams_nickname_column, teams_conference_column As Integer
Call VT_Teams_Sheet_Variables(teams_teamref_column, teams_name_column, teams_fullname_column, teams_fbsyear_column, teams_nickname_column, teams_conference_column)

Dim teams_current_row, teams_start_row, teams_final_row As Integer
teams_start_row = 2

teams_final_row = WorksheetFunction.CountA(Columns(teams_name_column))

Dim teams_current_team As String

'Determine if the team is FBS or FCS
teams_current_team = ""

For teams_current_row = teams_start_row To teams_final_row
    teams_current_team = Sheets(teams_sheet).Cells(teams_current_row, teams_name_column).Value
    
    If teams_name = teams_current_team Then
        If Sheets(teams_sheet).Cells(teams_current_row, teams_fbsyear_column).Value <= new_year Then
            isFBS = True
        Else
            isFBS = False
        End If
            
        Exit Sub

    Else
    End If
   
Next teams_current_row

MsgBox teams_name & " not found in '" & teams_sheet & "' worksheet. This macro will now exit."
End

End Sub

Sub VA_Worksheet_Names(ByRef elo_sheet, ByRef games_sheet, ByRef teams_sheet, ByRef odds_sheet, ByRef graph_sheet, ByRef pred_sheet, ByRef power_sheet)

elo_sheet = "E"
games_sheet = "G"
teams_sheet = "T"
odds_sheet = "O"
graph_sheet = "Oc"
pred_sheet = "Pd"
power_sheet = "Pw"

End Sub

Sub VB_Bets_Graph_Variables()

search_term = "ARI"
Call S_Search(search_term, search_row, search_column)
betsgraph_ari_column = search_column
betsgraph_date_column = search_column - 1
betsgraph_week_row = search_row
betsgraph_week_column = search_column - 1

search_term = "ATL"
Call S_Search(search_term, search_row, search_column)
betsgraph_atl_column = search_column

End Sub


Sub VE_Elo_Sheet_Variables(ByRef elo_ref_column, ByRef elo_year_column, ByRef elo_team_column, ByRef elo_yw_column, ByRef elo_week_column, ByRef elo_elo_column)

search_term = "EloRef"
Call S_Search(search_term, search_row, search_column)
elo_ref_column = search_column

search_term = "Year"
Call S_Search(search_term, search_row, search_column)
elo_year_column = search_column

search_term = "Team"
Call S_Search(search_term, search_row, search_column)
elo_team_column = search_column

search_term = "Y-W"
Call S_Search(search_term, search_row, search_column)
elo_yw_column = search_column

search_term = "Week"
Call S_Search(search_term, search_row, search_column)
elo_week_column = search_column

search_term = "Elo"
Call S_Search(search_term, search_row, search_column)
elo_elo_column = search_column

End Sub

Sub VG_Games_Sheet_Variables(ByRef games_winner_column, ByRef games_gameref_column, ByRef games_ra_column, ByRef games_rh_column, ByRef games_rhh_column, ByRef games_ea_column, ByRef games_eh_column, ByRef games_year_column, ByRef games_week_column, ByRef games_yw_column, ByRef games_awayteam_column, ByRef games_hometeam_column, ByRef games_neutral_column, ByRef games_prediction_column, ByRef games_awayscore_column, ByRef games_homescore_column, ByRef games_winnerstatus_column, ByRef games_predcorrect_column, ByRef games_sa_column, ByRef games_sh_column, games_rd_column, ByRef games_radash_column, ByRef games_rhdash_column, ByRef games_predtotal_column)

search_term = "Winner"
Call S_Search(search_term, search_row, search_column)
games_winner_column = search_column

search_term = "GameRef"
Call S_Search(search_term, search_row, search_column)
games_gameref_column = search_column

search_term = "RA"
Call S_Search(search_term, search_row, search_column)
games_ra_column = search_column

search_term = "RH"
Call S_Search(search_term, search_row, search_column)
games_rh_column = search_column

search_term = "RHH"
Call S_Search(search_term, search_row, search_column)
games_rhh_column = search_column

search_term = "EA"
Call S_Search(search_term, search_row, search_column)
games_ea_column = search_column

search_term = "EH"
Call S_Search(search_term, search_row, search_column)
games_eh_column = search_column

search_term = "Year"
Call S_Search(search_term, search_row, search_column)
games_year_column = search_column

search_term = "Week"
Call S_Search(search_term, search_row, search_column)
games_week_column = search_column

search_term = "Y-W"
Call S_Search(search_term, search_row, search_column)
games_yw_column = search_column

search_term = "AwayTeam"
Call S_Search(search_term, search_row, search_column)
games_awayteam_column = search_column

search_term = "HomeTeam"
Call S_Search(search_term, search_row, search_column)
games_hometeam_column = search_column

search_term = "Neutral"
Call S_Search(search_term, search_row, search_column)
games_neutral_column = search_column

search_term = "Prediction"
Call S_Search(search_term, search_row, search_column)
games_prediction_column = search_column

'search_term = "AwayFBS"
'Call S_Search(search_term, search_row, search_column)
'games_awayfbs_column = search_column

'search_term = "HomeFBS"
'Call S_Search(search_term, search_row, search_column)
'games_homefbs_column = search_column

search_term = "AwayScore"
Call S_Search(search_term, search_row, search_column)
games_awayscore_column = search_column

search_term = "HomeScore"
Call S_Search(search_term, search_row, search_column)
games_homescore_column = search_column

search_term = "Winner"
Call S_Search(search_term, search_row, search_column)
games_winner_column = search_column

search_term = "WinnerStatus"
Call S_Search(search_term, search_row, search_column)
games_winnerstatus_column = search_column

search_term = "PredCorrect"
Call S_Search(search_term, search_row, search_column)
games_predcorrect_column = search_column

search_term = "SA"
Call S_Search(search_term, search_row, search_column)
games_sa_column = search_column

search_term = "SH"
Call S_Search(search_term, search_row, search_column)
games_sh_column = search_column

search_term = "RD"
Call S_Search(search_term, search_row, search_column)
games_rd_column = search_column

search_term = "RA'"
Call S_Search(search_term, search_row, search_column)
games_radash_column = search_column

search_term = "RH'"
Call S_Search(search_term, search_row, search_column)
games_rhdash_column = search_column

search_term = "PredTotal"
Call S_Search(search_term, search_row, search_column)
games_predtotal_column = search_column

End Sub
Sub VH_Graph_Sheet_Variables(ByRef graph_gameref_row, ByRef graph_gameref_column, ByRef graph_datetime_column, ByRef graph_awayelodec_column, ByRef graph_homeelodec_column, ByRef graph_awayoddsdec_column, ByRef graph_homeoddsdec_column, ByRef graph_awayreqdec_column, ByRef graph_homereqdec_column, ByRef graph_awayplaced_column, ByRef graph_homeplaced_column, ByRef graph_awayprospect_column, ByRef graph_homeprospect_column, ByRef graph_week_row, ByRef graph_week_column)

search_term = "GameRef"
Call S_Search(search_term, search_row, search_column)
graph_gameref_row = search_row
graph_gameref_column = search_column

search_term = "DateTime"
Call S_Search(search_term, search_row, search_column)
graph_datetime_column = search_column

search_term = "AwayEloDec"
Call S_Search(search_term, search_row, search_column)
graph_awayelodec_column = search_column

search_term = "HomeEloDec"
Call S_Search(search_term, search_row, search_column)
graph_homeelodec_column = search_column

search_term = "AwayOddsDec"
Call S_Search(search_term, search_row, search_column)
graph_awayoddsdec_column = search_column

search_term = "HomeOddsDec"
Call S_Search(search_term, search_row, search_column)
graph_homeoddsdec_column = search_column

search_term = "AwayReqDec"
Call S_Search(search_term, search_row, search_column)
graph_awayreqdec_column = search_column

search_term = "HomeReqDec"
Call S_Search(search_term, search_row, search_column)
graph_homereqdec_column = search_column

search_term = "AwayPlaced"
Call S_Search(search_term, search_row, search_column)
graph_awayplaced_column = search_column

search_term = "HomePlaced"
Call S_Search(search_term, search_row, search_column)
graph_homeplaced_column = search_column

search_term = "AwayProspect"
Call S_Search(search_term, search_row, search_column)
graph_awayprospect_column = search_column

search_term = "HomeProspect"
Call S_Search(search_term, search_row, search_column)
graph_homeprospect_column = search_column

search_term = "Week"
Call S_Search(search_term, search_row, search_column)
graph_week_row = search_row
graph_week_column = search_column + 1

End Sub
Sub VO_Odds_Sheet_Variables(ByRef odds_gameref_column, ByRef odds_datetime_column, ByRef odds_awayodds_column, ByRef odds_homeodds_column, ByRef odds_awayelodec_column, ByRef odds_homeelodec_column, ByRef odds_awayoddsdec_column, ByRef odds_homeoddsdec_column, ByRef odds_awayedge_column, ByRef odds_homeedge_column, ByRef odds_betteam_column, ByRef odds_bethan_column, ByRef odds_kelly_column, ByRef odds_betamount_column, ByRef odds_winamount_column, ByRef odds_awayreqdec_column, ByRef odds_homereqdec_column, ByRef odds_awayreqodds_column, ByRef odds_homereqodds_column, ByRef odds_year_column, ByRef odds_week_column, ByRef odds_betplaced_column, ByRef odds_betodds_column, ByRef odds_betedge_column, ByRef odds_winlose_column, ByRef odds_opposeteam_column, ByRef odds_edgeedge_column, ByRef odds_betunits_column, ByRef odds_awaytrend_column, ByRef odds_hometrend_column, ByRef odds_neutral_column, ByRef odds_awayteam_column, ByRef odds_hometeam_column, _
ByRef odds_profitloss_column, ByRef odds_winnerstatus_column, ByRef odds_cumulative_column)

search_term = "GameRef"
Call S_Search(search_term, search_row, search_column)
odds_gameref_column = search_column

search_term = "DateTime"
Call S_Search(search_term, search_row, search_column)
odds_datetime_column = search_column

search_term = "AwayOdds"
Call S_Search(search_term, search_row, search_column)
odds_awayodds_column = search_column

search_term = "HomeOdds"
Call S_Search(search_term, search_row, search_column)
odds_homeodds_column = search_column

search_term = "AwayEloDec"
Call S_Search(search_term, search_row, search_column)
odds_awayelodec_column = search_column

search_term = "HomeEloDec"
Call S_Search(search_term, search_row, search_column)
odds_homeelodec_column = search_column

search_term = "AwayOddsDec"
Call S_Search(search_term, search_row, search_column)
odds_awayoddsdec_column = search_column

search_term = "HomeOddsDec"
Call S_Search(search_term, search_row, search_column)
odds_homeoddsdec_column = search_column

search_term = "AwayEdge"
Call S_Search(search_term, search_row, search_column)
odds_awayedge_column = search_column

search_term = "HomeEdge"
Call S_Search(search_term, search_row, search_column)
odds_homeedge_column = search_column

search_term = "BetTeam"
Call S_Search(search_term, search_row, search_column)
odds_betteam_column = search_column

search_term = "BetHAN"
Call S_Search(search_term, search_row, search_column)
odds_bethan_column = search_column

search_term = "Kelly%"
Call S_Search(search_term, search_row, search_column)
odds_kelly_column = search_column

search_term = "BetAmount"
Call S_Search(search_term, search_row, search_column)
odds_betamount_column = search_column

search_term = "WinAmount"
Call S_Search(search_term, search_row, search_column)
odds_winamount_column = search_column

search_term = "AwayReqDec"
Call S_Search(search_term, search_row, search_column)
odds_awayreqdec_column = search_column

search_term = "HomeReqDec"
Call S_Search(search_term, search_row, search_column)
odds_homereqdec_column = search_column

search_term = "AwayReqOdds"
Call S_Search(search_term, search_row, search_column)
odds_awayreqodds_column = search_column

search_term = "HomeReqOdds"
Call S_Search(search_term, search_row, search_column)
odds_homereqodds_column = search_column

search_term = "Year"
Call S_Search(search_term, search_row, search_column)
odds_year_column = search_column

search_term = "Week"
Call S_Search(search_term, search_row, search_column)
odds_week_column = search_column

search_term = "BetPlaced"
Call S_Search(search_term, search_row, search_column)
odds_betplaced_column = search_column

search_term = "BetOdds"
Call S_Search(search_term, search_row, search_column)
odds_betodds_column = search_column

search_term = "BetEdge"
Call S_Search(search_term, search_row, search_column)
odds_betedge_column = search_column

search_term = "WIN/LOSE"
Call S_Search(search_term, search_row, search_column)
odds_winlose_column = search_column

search_term = "OpposeTeam"
Call S_Search(search_term, search_row, search_column)
odds_opposeteam_column = search_column

search_term = "EdgeEdge"
Call S_Search(search_term, search_row, search_column)
odds_edgeedge_column = search_column

search_term = "BetUnits"
Call S_Search(search_term, search_row, search_column)
odds_betunits_column = search_column

search_term = "AT"
Call S_Search(search_term, search_row, search_column)
odds_awaytrend_column = search_column

search_term = "HT"
Call S_Search(search_term, search_row, search_column)
odds_hometrend_column = search_column

search_term = "Neutral"
Call S_Search(search_term, search_row, search_column)
odds_neutral_column = search_column

search_term = "AwayTeam"
Call S_Search(search_term, search_row, search_column)
odds_awayteam_column = search_column

search_term = "HomeTeam"
Call S_Search(search_term, search_row, search_column)
odds_hometeam_column = search_column

search_term = "Profit"
Call S_Search(search_term, search_row, search_column)
odds_profitloss_column = search_column

search_term = "WinnerStatus"
Call S_Search(search_term, search_row, search_column)
odds_winnerstatus_column = search_column

search_term = "Cumulative"
Call S_Search(search_term, search_row, search_column)
odds_cumulative_column = search_column

End Sub

Sub VP_Predictions_Sheet_Variables(ByRef pred_gameref_column, ByRef pred_awayteam_column, ByRef pred_hometeam_column, ByRef pred_at_column)

search_term = "GameRef"
Call S_Search(search_term, search_row, search_column)
pred_gameref_column = search_column

search_term = "AwayTeam"
Call S_Search(search_term, search_row, search_column)
pred_awayteam_column = search_column

search_term = "HomeTeam"
Call S_Search(search_term, search_row, search_column)
pred_hometeam_column = search_column

search_term = "At"
Call S_Search(search_term, search_row, search_column)
pred_at_column = search_column

End Sub

Sub VR_Rankings_Sheet_Variables(ByRef power_rank_column, ByRef power_elo_column, ByRef power_team_column, ByRef power_wins_column, ByRef power_losses_column, ByRef power_ties_column)

search_term = "Rank"
Call S_Search(search_term, search_row, search_column)
power_rank_column = search_column

search_term = "Elo"
Call S_Search(search_term, search_row, search_column)
power_elo_column = search_column

search_term = "Team"
Call S_Search(search_term, search_row, search_column)
power_team_column = search_column

search_term = "Wins"
Call S_Search(search_term, search_row, search_column)
power_wins_column = search_column

search_term = "Losses"
Call S_Search(search_term, search_row, search_column)
power_losses_column = search_column

search_term = "Ties"
Call S_Search(search_term, search_row, search_column)
power_ties_column = search_column

End Sub

Sub VT_Teams_Sheet_Variables(ByRef teams_teamref_column, ByRef teams_name_column, ByRef teams_fullname_column, ByRef teams_fbsyear_column, ByRef teams_nickname_column, ByRef teams_conference_column)

search_term = "TeamRef"
Call S_Search(search_term, search_row, search_column)
teams_teamref_column = search_column

search_term = "Name"
Call S_Search(search_term, search_row, search_column)
teams_name_column = search_column

search_term = "FullName"
Call S_Search(search_term, search_row, search_column)
teams_fullname_column = search_column

search_term = "Nickname"
Call S_Search(search_term, search_row, search_column)
teams_nickname_column = search_column

search_term = "Conference"
Call S_Search(search_term, search_row, search_column)
teams_conference_column = search_column

search_term = "FBSYear"
Call S_Search(search_term, search_row, search_column)
teams_fbsyear_column = search_column

End Sub

Sub ZA_Populate_Historical_Data()

Call F_Switch_Off_Functionality

'Sheet names
Dim elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet As String
Call VA_Worksheet_Names(elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet)

'Elo sheet variables
Sheets(elo_sheet).Select
Dim elo_ref_column, elo_year_column, elo_team_column, elo_yw_column, elo_week_column, elo_elo_column As Integer
Call VE_Elo_Sheet_Variables(elo_ref_column, elo_year_column, elo_team_column, elo_yw_column, elo_week_column, elo_elo_column)

'Games variables
Sheets(games_sheet).Select
Dim games_winner_column, games_gameref_column, games_ra_column, games_rh_column, games_rhh_column, games_ea_column, games_eh_column, games_year_column, games_week_column, games_yw_column, games_awayteam_column, games_hometeam_column, games_neutral_column, games_prediction_column, games_awayscore_column, games_homescore_column, games_winnerstatus_column, games_predcorrect_column, games_sa_column, games_sh_column, games_rd_column, games_radash_column, games_rhdash_column, games_predtotal_column As Integer
Call VG_Games_Sheet_Variables(games_winner_column, games_gameref_column, games_ra_column, games_rh_column, games_rhh_column, games_ea_column, games_eh_column, games_year_column, games_week_column, games_yw_column, games_awayteam_column, games_hometeam_column, games_neutral_column, games_prediction_column, games_awayscore_column, games_homescore_column, games_winnerstatus_column, games_predcorrect_column, games_sa_column, games_sh_column, games_rd_column, games_radash_column, games_rhdash_column, games_predtotal_column)

Dim start_row, current_row, end_row As Integer
start_row = 2
Sheets(games_sheet).Select
end_row = WorksheetFunction.CountA(Columns(games_year_column))
current_row = start_row

Dim current_year, start_year, start_week As Integer
start_year = Sheets(games_sheet).Cells(start_row, games_year_column).Value
current_year = start_year
start_week = 0 'Sheets(games_sheet).Cells(2, 4).Value

'Teams variables
Sheets(teams_sheet).Select
Dim teams_teamref_column, teams_name_column, teams_fullname_column, teams_fbsyear_column, teams_nickname_column, teams_conference_column As Integer
Call VT_Teams_Sheet_Variables(teams_teamref_column, teams_name_column, teams_fullname_column, teams_fbsyear_column, teams_nickname_column, teams_conference_column)

Dim teams_current_row, teams_start_row, teams_final_row As Integer
teams_start_row = 2
teams_final_row = WorksheetFunction.CountA(Columns(teams_name_column))

Dim teams_current_team As String

'FBS variable
Dim FBS As Double
FBS = Range("FBS").Value

'FCS variable
Dim FCS As Double
FCS = Range("FCS").Value

'Find Division variables
Dim isFBS As Boolean
Dim teams_name As String
Dim new_year As Integer



'Delete G data
Sheets(games_sheet).Select
Range(Cells(2, 1), Cells(1000000, 2)).ClearContents
Range(Cells(2, 8), Cells(1000000, 13)).ClearContents
Range(Cells(2, 16), Cells(1000000, 26)).ClearContents



'Main loop
For current_row = start_row To end_row
    
    Sheets(games_sheet).Select
    new_year = Cells(current_row, games_year_column).Value

    If new_year = current_year Then
    
    Else
        
        Sheets(elo_sheet).Select
        Call E_Year_End
        current_year = new_year
        Sheets(games_sheet).Select
                
    End If
    
    Sheets(games_sheet).Select
    Cells(current_row, games_gameref_column).Select
    Call A_Populate_R
    
    Call F_Switch_Off_Functionality
    Call ZZ_Calculate
    
    Cells(current_row, games_gameref_column).Select
    Call C_Results
    
    Call F_Switch_Off_Functionality
    Call ZZ_Calculate
    
    Call F_Switch_Off_Functionality
    
    'If current_row / 1000 = Int(current_row / 1000) Then
        'Stop
    'Else
    'End If

Next current_row

Call O_Switch_On_Functionality

Sheets(games_sheet).Select

MsgBox "Historical Data Macro Complete!"

End Sub
Sub ZB_Find_Closing_Odds()

Dim closing_sheet As String
closing_sheet = "ClosingOdds"

'Sheet names
Dim elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet As String
Call VA_Worksheet_Names(elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet)

'Odds variables
Sheets(odds_sheet).Select
Dim odds_gameref_column, odds_datetime_column, odds_awayodds_column, odds_homeodds_column, odds_awayelodec_column, odds_homeelodec_column, odds_awayoddsdec_column, odds_homeoddsdec_column, odds_awayedge_column, odds_homeedge_column, odds_betteam_column, odds_bethan_column, odds_kelly_column, odds_betamount_column, odds_winamount_column, odds_awayreqdec_column, odds_homereqdec_column, odds_awayreqodds_column, odds_homereqodds_column, odds_year_column, odds_week_column, odds_betplaced_column, odds_betodds_column, odds_betedge_column, odds_winlose_column, odds_opposeteam_column, odds_edgeedge_column, odds_betunits_column, odds_awaytrend_column, odds_hometrend_column, odds_neutral_column, odds_awayteam_column, odds_hometeam_column, odds_profitloss_column, odds_winnerstatus_column, odds_cumulative_column As Integer
Call VO_Odds_Sheet_Variables(odds_gameref_column, odds_datetime_column, odds_awayodds_column, odds_homeodds_column, odds_awayelodec_column, odds_homeelodec_column, odds_awayoddsdec_column, odds_homeoddsdec_column, odds_awayedge_column, odds_homeedge_column, odds_betteam_column, odds_bethan_column, odds_kelly_column, odds_betamount_column, odds_winamount_column, odds_awayreqdec_column, odds_homereqdec_column, odds_awayreqodds_column, odds_homereqodds_column, odds_year_column, odds_week_column, odds_betplaced_column, odds_betodds_column, odds_betedge_column, odds_winlose_column, odds_opposeteam_column, odds_edgeedge_column, odds_betunits_column, odds_awaytrend_column, odds_hometrend_column, odds_neutral_column, odds_awayteam_column, odds_hometeam_column, odds_profitloss_column, odds_winnerstatus_column, odds_cumulative_column)

'Games variables
Sheets(games_sheet).Select
Dim games_winner_column, games_gameref_column, games_ra_column, games_rh_column, games_rhh_column, games_ea_column, games_eh_column, games_year_column, games_week_column, games_yw_column, games_awayteam_column, games_hometeam_column, games_neutral_column, games_prediction_column, games_awayscore_column, games_homescore_column, games_winnerstatus_column, games_predcorrect_column, games_sa_column, games_sh_column, games_rd_column, games_radash_column, games_rhdash_column, games_predtotal_column As Integer
Call VG_Games_Sheet_Variables(games_winner_column, games_gameref_column, games_ra_column, games_rh_column, games_rhh_column, games_ea_column, games_eh_column, games_year_column, games_week_column, games_yw_column, games_awayteam_column, games_hometeam_column, games_neutral_column, games_prediction_column, games_awayscore_column, games_homescore_column, games_winnerstatus_column, games_predcorrect_column, games_sa_column, games_sh_column, games_rd_column, games_radash_column, games_rhdash_column, games_predtotal_column)

Sheets(games_sheet).Select

Dim no_games As Integer
no_games = WorksheetFunction.CountA(Columns(games_gameref_column))

Sheets(odds_sheet).Select
Dim no_odds As Integer
no_odds = WorksheetFunction.CountA(Columns(odds_gameref_column))

Dim current_game As Integer

Dim current_new As Integer
current_new = 2

Dim game_found As Boolean
game_found = False

For current_game = 2 To no_games
    current_ref = Sheets(games_sheet).Cells(current_game, 1).Value
    
    current_odds = no_odds
    
    Do Until current_odds = 2
        If Sheets(odds_sheet).Cells(current_odds, 1).Value = current_ref Then
            
            Sheets(closing_sheet).Cells(current_new, 1).Value = current_ref
            Sheets(closing_sheet).Cells(current_new, 2).Value = Sheets(odds_sheet).Cells(current_odds, odds_awayodds_column).Value
            Sheets(closing_sheet).Cells(current_new, 3).Value = Sheets(odds_sheet).Cells(current_odds, odds_homeodds_column).Value
            game_found = True
            Exit Do
        Else
        
            game_found = False
        
        End If
    
        current_odds = current_odds - 1
    Loop
    
    If game_found = True Then
        current_new = current_new + 1
    Else
    End If
    
Next current_game

End Sub

Sub ZC_Edges()

'Speed up macro
Call F_Switch_Off_Functionality

'Declare variables
Dim results_sheet, h_sheet As String
results_sheet = "Edge"
h_sheet = "H1_16"

Dim h_input_row, h_lowero_column, h_uppero_column, h_awayedge_column, h_awayprofit_column, h_homeprofit_column, h_homeedge_column As Integer
h_input_row = 1
h_lowero_column = 26
h_uppero_column = 28
h_awayprofit_column = 18
h_awayedge_column = 19
h_homeprofit_column = 20
h_homeedge_column = 21

Dim edge_start_row, edge_current_row, edge_final_row, edge_lowero_column, edge_uppero_column, edge_edge_column, edge_profit_column As Integer
Sheets(results_sheet).Select
edge_current_row = edge_start_row
edge_lowero_column = 1
edge_uppero_column = 2
edge_awayedge_column = 3
edge_awayprofit_column = 4
edge_homeedge_column = 5
edge_homeprofit_column = 6
edge_start_row = WorksheetFunction.CountA(Columns(edge_awayedge_column)) + 1
edge_final_row = WorksheetFunction.CountA(Columns(edge_lowero_column))

If edge_start_row = edge_final_row + 1 Then
    MsgBox "No rows to be filled. This macro will now exit."
    Call O_Switch_On_Functionality
    Exit Sub
Else

End If

Dim StartTime As Double
Dim SecondsElapsed As Double

'Main loop
For edge_current_row = edge_start_row To edge_final_row

    'Paste values into H sheet
    Sheets(h_sheet).Cells(h_input_row, h_lowero_column).Value = Sheets(results_sheet).Cells(edge_current_row, edge_lowero_column).Value
    Sheets(h_sheet).Cells(h_input_row, h_uppero_column).Value = Sheets(results_sheet).Cells(edge_current_row, edge_uppero_column).Value
          
    'calculate
    StartTime = Timer
    Calculate
        
    Do While Application.CalculationState <> xlDone
         SecondsElapsed = Round(Timer - StartTime, 0)
         DoEvents
         If SecondsElapsed > 60 Then
            MsgBox "Too much time (" & SecondsElapsed & " seconds) has elapsed in this calculation, this macro will exit."
            Call O_Switch_On_Functionality
            Exit Sub
         Else
         
         End If
    Loop
    
    'copy results into workbook
    Sheets(results_sheet).Cells(edge_current_row, edge_awayedge_column).Value = Sheets(h_sheet).Cells(h_input_row, h_awayedge_column).Value
    Sheets(results_sheet).Cells(edge_current_row, edge_homeedge_column).Value = Sheets(h_sheet).Cells(h_input_row, h_homeedge_column).Value
    
    Sheets(results_sheet).Cells(edge_current_row, edge_awayprofit_column).Value = Sheets(h_sheet).Cells(h_input_row, h_awayprofit_column).Value
    Sheets(results_sheet).Cells(edge_current_row, edge_homeprofit_column).Value = Sheets(h_sheet).Cells(h_input_row, h_homeprofit_column).Value

Next edge_current_row

'Reset defaults
Call O_Switch_On_Functionality

MsgBox "Ta-da!"

End Sub

Sub ZF_Populate_FBS_Columns()

Call F_Switch_Off_Functionality

'Sheet names
Dim elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet As String
Call VA_Worksheet_Names(elo_sheet, games_sheet, teams_sheet, odds_sheet, graph_sheet, pred_sheet, power_sheet)

'Games variables
Sheets(games_sheet).Select
Dim games_winner_column, games_gameref_column, games_ra_column, games_rh_column, games_rhh_column, games_ea_column, games_eh_column, games_year_column, games_week_column, games_yw_column, games_awayteam_column, games_hometeam_column, games_neutral_column, games_prediction_column, games_awayscore_column, games_homescore_column, games_winnerstatus_column, games_predcorrect_column, games_sa_column, games_sh_column, games_rd_column, games_radash_column, games_rhdash_column, games_predtotal_column As Integer
Call VG_Games_Sheet_Variables(games_winner_column, games_gameref_column, games_ra_column, games_rh_column, games_rhh_column, games_ea_column, games_eh_column, games_year_column, games_week_column, games_yw_column, games_awayteam_column, games_hometeam_column, games_neutral_column, games_prediction_column, games_awayscore_column, games_homescore_column, games_winnerstatus_column, games_predcorrect_column, games_sa_column, games_sh_column, games_rd_column, games_radash_column, games_rhdash_column, games_predtotal_column)

Dim games_current_row, games_start_row, games_final_row As Integer
games_final_row = WorksheetFunction.CountA(Columns(games_year_column))

games_start_row = 2

Dim games_current_team As String

Dim i As Integer

'Teams variables
Sheets(teams_sheet).Select
Dim teams_teamref_column, teams_name_column, teams_fullname_column, teams_fbsyear_column, teams_nickname_column, teams_conference_column As Integer
Call VT_Teams_Sheet_Variables(teams_teamref_column, teams_name_column, teams_fullname_column, teams_fbsyear_column, teams_nickname_column, teams_conference_column)

Dim teams_current_row, teams_start_row, teams_final_row As Integer
teams_start_row = 2

teams_final_row = WorksheetFunction.CountA(Columns(teams_name_column))

Dim teams_current_team As String

'Loop around each game
For games_current_row = games_start_row To games_final_row

    For i = 0 To 1
        games_current_team = Sheets(games_sheet).Cells(games_current_row, games_awayteam_column + i).Value
    
        'loop around team
        teams_current_row = teams_start_row
        teams_current_team = ""
        
        Do Until teams_current_row > teams_final_row
            teams_current_team = Sheets(teams_sheet).Cells(teams_current_row, teams_name_column).Value
            
            If games_current_team = teams_current_team Then
                If Sheets(teams_sheet).Cells(teams_current_row, teams_fbsyear_column).Value <= Sheets(games_sheet).Cells(games_current_row, games_year_column).Value Then
                    Sheets(games_sheet).Cells(games_current_row, games_awayfbs_column + i).Value = "Y"
                Else
                    Sheets(games_sheet).Cells(games_current_row, games_awayfbs_column + i).Value = "N"
                End If
                    
                Exit Do
            Else
                 Sheets(games_sheet).Cells(games_current_row, games_awayfbs_column + i).Value = "N"
            End If
        
            teams_current_row = teams_current_row + 1
        Loop
        
    Next i
    
Next games_current_row

Call O_Switch_On_Functionality

MsgBox "Finished!"

End Sub

Sub ZZ_Calculate()
    'calculate
    StartTime = Timer
    Calculate
        
    Do While Application.CalculationState <> xlDone
         SecondsElapsed = Round(Timer - StartTime, 0)
         DoEvents
         If SecondsElapsed > 60 Then
            MsgBox "Too much time (" & SecondsElapsed & " seconds) has elapsed in this calculation, this macro will exit."
            Call O_Switch_On_Functionality
            Exit Sub
         Else
         
         End If
    Loop
End Sub
