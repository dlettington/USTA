Attribute VB_Name = "Validation"
Option Explicit

'Set to true to see each individual validation message (helpful for debugging)
Const CONSOLIDATE_MESSAGES As Boolean = True

'Callback for validateButton onAction
Sub ValidateMatches(control As IRibbonControl)
    USTAValidation
End Sub

Function USTAValidation() As Boolean
    'Update the match team IDs using the (possibly modified) team names prior to any validation
    If Not UpdateMatchTeamIDs Then
        MsgBox "Could not update team IDs from current team names. Please check that the matches table is not corrupt.", _
            vbExclamation, ADDIN_NAME
            
        Exit Function
    End If
    
    'Update the facility IDs using the (possibly modified) facility names prior to any validation
    If Not UpdateFacilityIDs Then
        MsgBox "Could not update facility IDs from current facility names. Please check that the matches table is not corrupt.", _
            vbExclamation, ADDIN_NAME
            
        Exit Function
    End If
    
    Dim totalErrors As Integer
    
    totalErrors = 0
    
    ClearValidationErrorBorder Range(MATCHES_TABLE_NAME)
    
    'Run the validations and collect the number of issues found
    ValidateTeamName totalErrors
    ValidateHomeAndVisitor totalErrors
    ValidateNumberOfMatches totalErrors
    ValidateFacilityName totalErrors
    ValidateFacility totalErrors
    ValidateMatchDate totalErrors

    If totalErrors = 0 Then
        MsgBox "Validation found no errors.", vbInformation, ADDIN_NAME
        USTAValidation = True
    Else
        MsgBox "Validation found " & totalErrors & " errors in the match information.", vbExclamation, ADDIN_NAME
        USTAValidation = False
    End If
End Function

'Check that home team and visiting team are not the same each match
Function ValidateHomeAndVisitor(ByRef totalErrors As Integer) As Boolean
    On Error GoTo Err

    ValidateHomeAndVisitor = True
    
    Dim errors As Integer
    Dim errorsSame As Integer
    Dim errorsBye As Integer
    errors = 0
    errorsSame = 0
    errorsBye = 0
    
    Dim errorString As String
    Const ERROR_SAME_TEAMS As String = "Home and visiting teams must be different "
    Const ERROR_BYE_TEAMS As String = "Home and visiting teams cannot both have a bye week "
    Const ERROR_BOTH As String = "Home and visiting teams must be different and cannot both have a bye week "
    
    Dim homeTeamID As String
    Dim visitingTeamID As String
    Dim matchID As String
    
    'Loop through team IDs from teams lookup table
    Dim i As Integer
    
    For i = 1 To Range(MATCHES_TABLE_NAME).Rows.Count
        homeTeamID = Range(HOME_TEAM_ID_RANGE_NAME).Value2(i, 1)
        visitingTeamID = Range(VISITING_TEAM_ID_RANGE_NAME).Value2(i, 1)
        
        If homeTeamID = visitingTeamID Then
            'Highlight match ID before showing message to make it more clear to the user
            SetValidationErrorBorder Range(MATCHES_TABLE_NAME).Cells(i, 1)
            SetValidationErrorBorder Range(HOME_TEAM_NAME_RANGE_NAME).Cells(i, 1)
            SetValidationErrorBorder Range(VISITING_TEAM_NAME_RANGE_NAME).Cells(i, 1)
            
            matchID = Range(MATCH_ID_RANGE_NAME).Value2(i, 1)
                    
            If homeTeamID <> BYE_WEEK_TEAM_ID Then
                'Home and visiting team are the same
                errorString = ERROR_SAME_TEAMS
                errorsSame = errorsSame + 1
            Else
                'Home and visiting team are both bye
                errorString = ERROR_BYE_TEAMS
                errorsBye = errorsBye + 1
            End If
                
            If Not CONSOLIDATE_MESSAGES Then
                MsgBox errorString & "[match ID: " & matchID & "].", vbInformation, ADDIN_NAME
            End If
                
            errors = errors + 1
            totalErrors = totalErrors + 1
            ValidateHomeAndVisitor = False
        End If
    Next i
    
    If CONSOLIDATE_MESSAGES And errors > 0 Then
        If errorsSame > 0 And errorsBye > 0 Then
            errorString = ERROR_BOTH
        End If
            
        MsgBox errorString & GetErrString(errors) & ".", vbInformation, ADDIN_NAME
    End If
   
    Exit Function
    
Err:
    DisplayError "ValidateHomeAndVisitor", Err
End Function

Function ValidateNumberOfMatches(ByRef totalErrors As Integer) As Boolean
    On Error GoTo Err
    
    ValidateNumberOfMatches = True
    
    Dim validationMessage As String
    Dim matchesPerTeam As Integer
    Dim teamMatches As Integer
    Dim teamID As String
    Dim teamName As String
    
    matchesPerTeam = -1
    
    'Loop through team IDs from teams lookup table
    Dim i As Integer
    
    For i = 1 To Range(TEAMS_TABLE_NAME).Rows.Count
        'Get number of matches for each team
        teamID = Range(TEAM_ID_RANGE_NAME).Value2(i, 1)
        teamName = Range(TEAM_NAME_RANGE_NAME).Value2(i, 1)
        teamMatches = Application.WorksheetFunction.CountIf(Range(MATCH_TEAM_INFO_RANGE_NAME), teamID)
        
        validationMessage = validationMessage & vbNewLine & "'" & teamName & "' has " & teamMatches & " matches"
        
        If matchesPerTeam = -1 Then
            matchesPerTeam = teamMatches
        Else
            If matchesPerTeam <> teamMatches Then
                'This team does not have the same number of matches as the previous teams checked
                ValidateNumberOfMatches = False
            End If
        End If
    Next i
    
    If Not ValidateNumberOfMatches Then
        MsgBox "All teams must play the same number of matches: " & vbNewLine & _
            validationMessage, vbInformation, ADDIN_NAME
            
        totalErrors = totalErrors + 1
    End If
    
    Exit Function
    
Err:
    DisplayError "ValidateNumberOfMatches", Err
End Function

Function ValidateFacility(ByRef totalErrors As Integer) As Boolean
    On Error GoTo Err
    
    ValidateFacility = True
    
    Dim errors As Integer
    errors = 0
    
    Dim matchID As String
    Dim homeTeamID As String
    Dim visitingTeamID As String
    Dim facilityID As String
    Dim facilityName As String
    
    'Loop through facility IDs from matches table
    Dim i As Integer
    
    For i = 1 To Range(MATCHES_TABLE_NAME).Rows.Count
        'Facility should only be missing (0) if the match is a bye for at least one of the teams
        'However, if facility name is TBD then allow missing facility since that means the TBD
        'will be cleared up in the Tennis Link web site
        facilityID = Range(FACILITY_ID_RANGE_NAME).Value2(i, 1)
        facilityName = Range(FACILITY_NAME_RANGE_NAME).Value2(i, 1)
        homeTeamID = Range(HOME_TEAM_ID_RANGE_NAME).Value2(i, 1)
        
        If facilityID = BYE_FACILITY_ID Then
            If homeTeamID <> BYE_WEEK_TEAM_ID Then
                visitingTeamID = Range(VISITING_TEAM_ID_RANGE_NAME).Value2(i, 1)
                        
                If visitingTeamID <> BYE_WEEK_TEAM_ID And facilityName <> MISSING_FACILITY_NAME Then
                    'Highlight match ID before showing message to make it more clear to the user
                    SetValidationErrorBorder Range(MATCHES_TABLE_NAME).Cells(i, 1)
                    SetValidationErrorBorder Range(FACILITY_NAME_RANGE_NAME).Cells(i, 1)
                    
                    matchID = Range(MATCH_ID_RANGE_NAME).Value2(i, 1)
                    
                    If Not CONSOLIDATE_MESSAGES Then
                        MsgBox "A facility for the match has not been defined [match ID: " & matchID & "].", _
                            vbInformation, ADDIN_NAME
                    End If
                
                    ValidateFacility = False
                    
                    errors = errors + 1
                    totalErrors = totalErrors + 1
                End If
            End If
        End If
    Next i
    
    If CONSOLIDATE_MESSAGES And errors > 0 Then
        MsgBox "A facility for a match has not been defined " & GetErrString(errors) & ".", _
            vbInformation, ADDIN_NAME
    End If

    Exit Function
    
Err:
    DisplayError "ValidateFacility", Err
End Function

'This situation shouldn't happen in normal circumstances since data validation
'should check for a valid team name during input.  However, files generated
'using other mechanisms may contain team names that are not in the Teams table.
Function ValidateTeamName(ByRef totalErrors As Integer) As Boolean
    On Error GoTo Err
    
    ValidateTeamName = True
    
    Dim errors As Integer
    errors = 0
    
    Dim homeTeamName As String
    Dim homeTeamID As String
    Dim visitingTeamName As String
    Dim visitingTeamID As String
    Dim matchID As String
    
    'Loop through team IDs from matches table
    Dim i As Integer
    
    For i = 1 To Range(MATCHES_TABLE_NAME).Rows.Count
        'Facility should only be missing (0) if the match is a bye for at least one of the teams
        homeTeamName = Range(HOME_TEAM_NAME_RANGE_NAME).Value2(i, 1)
        homeTeamID = Range(HOME_TEAM_ID_RANGE_NAME).Value2(i, 1)
        
        visitingTeamName = Range(VISITING_TEAM_NAME_RANGE_NAME).Value2(i, 1)
        visitingTeamID = Range(VISITING_TEAM_ID_RANGE_NAME).Value2(i, 1)
        
        'Since team ID update is run prior to validation, any invalid team name will
        'now have its team ID set to 0.  Error situation will have a team ID of 0 and
        'a non-blank team name.
        If homeTeamID = BYE_WEEK_TEAM_ID And Len(homeTeamName) > 0 Then
            SetValidationErrorBorder Range(MATCHES_TABLE_NAME).Cells(i, 1)
            SetValidationErrorBorder Range(HOME_TEAM_NAME_RANGE_NAME).Cells(i, 1)
                    
            matchID = Range(MATCH_ID_RANGE_NAME).Value2(i, 1)
                    
            If Not CONSOLIDATE_MESSAGES Then
                MsgBox "Home team name '" & homeTeamName & "' is invalid [match ID: " & matchID & "].", _
                    vbInformation, ADDIN_NAME
            End If
                
            ValidateTeamName = False
            
            errors = errors + 1
            totalErrors = totalErrors + 1
        End If
        
        If visitingTeamID = BYE_WEEK_TEAM_ID And Len(visitingTeamName) > 0 Then
            SetValidationErrorBorder Range(MATCHES_TABLE_NAME).Cells(i, 1)
            SetValidationErrorBorder Range(VISITING_TEAM_NAME_RANGE_NAME).Cells(i, 1)
                    
            matchID = Range(MATCH_ID_RANGE_NAME).Value2(i, 1)
                    
            If Not CONSOLIDATE_MESSAGES Then
                MsgBox "Visiting team name '" & visitingTeamName & "' is invalid [match ID: " & matchID & "].", _
                    vbInformation, ADDIN_NAME
            End If
                
            ValidateTeamName = False
            
            errors = errors + 1
            totalErrors = totalErrors + 1
        End If
    Next i
    
    If CONSOLIDATE_MESSAGES And errors > 0 Then
        MsgBox "Found invalid team names " & GetErrString(errors) & ".", _
            vbInformation, ADDIN_NAME
    End If

    Exit Function
    
Err:
    DisplayError "ValidateTeamName", Err
End Function

Function ValidateFacilityName(ByRef totalErrors As Integer) As Boolean
    On Error GoTo Err
    
    ValidateFacilityName = True
    
    Dim errors As Integer
    errors = 0
    
    Dim facilityName As String
    Dim facilityID As String
    Dim matchID As String
    
    'Loop through facility IDs from matches table
    Dim i As Integer
    
    For i = 1 To Range(MATCHES_TABLE_NAME).Rows.Count
        'Facility should only be missing (0) if the match is a bye for at least one of the teams
        facilityName = Range(FACILITY_NAME_RANGE_NAME).Value2(i, 1)
        facilityID = Range(FACILITY_ID_RANGE_NAME).Value2(i, 1)
        
        'Since facility ID update is run prior to validation, any invalid facility name will
        'now have its facility ID set to 0.  Error situation will have a facility ID of 0 and
        'a TBD facility name.
        If facilityID = INVALID_FACILITY_ID And facilityName <> MISSING_FACILITY_NAME And Len(facilityName) > 0 Then
            SetValidationErrorBorder Range(MATCHES_TABLE_NAME).Cells(i, 1)
            SetValidationErrorBorder Range(FACILITY_NAME_RANGE_NAME).Cells(i, 1)
                    
            matchID = Range(MATCH_ID_RANGE_NAME).Value2(i, 1)
                    
            If Not CONSOLIDATE_MESSAGES Then
                MsgBox "Facility name '" & facilityName & "' is invalid [match ID: " & matchID & "].", _
                    vbInformation, ADDIN_NAME
            End If
                
            ValidateFacilityName = False
            
            errors = errors + 1
            totalErrors = totalErrors + 1
        End If
    Next i

    If CONSOLIDATE_MESSAGES And errors > 0 Then
        MsgBox "Found invalid facility names " & GetErrString(errors) & ".", _
            vbInformation, ADDIN_NAME
    End If
            
    Exit Function
    
Err:
    DisplayError "ValidateFacilityName", Err
End Function

'Dates exported from Access appear as text in Excel cells. This function attempts
'to convert them to dates so they may be used in date comparisons during validation.
Function GetMatchDate(matchDate As Variant, ByRef isInvalid As Boolean) As Date
    On Error GoTo Err
    
    isInvalid = False
    
    GetMatchDate = DateValue(matchDate)
    
    Exit Function
    
Err:
    isInvalid = True
End Function

Function ValidateMatchDate(ByRef totalErrors As Integer) As Boolean
    On Error GoTo Err
    
    ValidateMatchDate = True
    
    Dim errors As Integer
    errors = 0
    
    Dim matchDate As Date
    Dim matchDateString As String
    Dim seasonStartDate As Date
    Dim seasonStartDateString As String
    Dim seasonEndDate As Date
    Dim seasonEndDateString As String
    Dim isInvalidMatchDate As Boolean
    Dim isInvalidStartDate As Boolean
    Dim isInvalidEndDate As Boolean
    
    'Should only be one row in the header table
    seasonStartDateString = Range(HEADER_START_DATE_RANGE_NAME)
    seasonStartDate = GetMatchDate(seasonStartDateString, isInvalidStartDate)
    seasonEndDateString = Range(HEADER_END_DATE_RANGE_NAME)
    seasonEndDate = GetMatchDate(seasonEndDateString, isInvalidEndDate)
    
    If isInvalidStartDate Or isInvalidEndDate Or seasonStartDate > seasonEndDate Then
        MsgBox "Season date range (" & seasonStartDateString & " to " & seasonEndDateString & _
            ") is invalid. Match date validation will not be performed.", _
            vbInformation, ADDIN_NAME
        
        Exit Function
    End If

    'Loop through match dates from matches table
    Dim i As Integer
    
    For i = 1 To Range(MATCHES_TABLE_NAME).Rows.Count
        matchDateString = Range(MATCH_DATE_RANGE_NAME).Value2(i, 1)
        matchDate = GetMatchDate(matchDateString, isInvalidMatchDate)
        
        'Check that match dates fall in the range of the season
        If isInvalidMatchDate Or matchDate < seasonStartDate Or matchDate > seasonEndDate Then
            'Highlight match date before showing message to make it more clear to the user
            SetValidationErrorBorder Range(MATCHES_TABLE_NAME).Cells(i, 1)
            SetValidationErrorBorder Range(MATCH_DATE_RANGE_NAME).Cells(i, 1)
                    
            If Not CONSOLIDATE_MESSAGES Then
                If isInvalidMatchDate Then
                    MsgBox "Match date (" & matchDateString & ") is not a valid date.", _
                    vbInformation, ADDIN_NAME
                Else
                    MsgBox "Match date (" & matchDate & ") is not within the current season (" & _
                    seasonStartDate & " to " & seasonEndDate & ").", _
                    vbInformation, ADDIN_NAME
                End If
            End If
                
            ValidateMatchDate = False
            
            errors = errors + 1
            totalErrors = totalErrors + 1
        End If
    Next i
    
    If CONSOLIDATE_MESSAGES And errors > 0 Then
        MsgBox "Found invalid match date " & GetErrString(errors) & ".", _
            vbInformation, ADDIN_NAME
    End If
    
    Exit Function
    
Err:
    DisplayError "ValidateMatchDate", Err
End Function


Sub ClearAllValidationErrorBorders()
    ClearValidationErrorBorder Range(MATCHES_TABLE_NAME)
End Sub

Sub SetValidationErrorBorder(errorRange As Range)
    With errorRange.Borders
        .LineStyle = xlDouble
        .Color = vbRed
        .Weight = xlThick
    End With
End Sub

Sub ClearValidationErrorBorder(errorRange As Range)
    errorRange.Borders.LineStyle = xlNone
End Sub

Function GetErrString(numErrors As Integer) As String
    If numErrors = 1 Then
        GetErrString = "[one error]"
    Else
        GetErrString = "[" & numErrors & " errors]"
    End If
End Function
