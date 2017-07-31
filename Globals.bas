Attribute VB_Name = "Globals"
Option Explicit

Public CurrentAccessFilePath As String

Global Const ADDIN_NAME As String = "USTA Tennis Link"
Global Const DEFAULT_ACCESS_FOLDER As String = "%USERPROFILE%\Downloads\"
'Global Const DEFAULT_ACCESS_FOLDER As String = "%USERPROFILE%\Documents\USTA\"  'For Drew's testing

Global Const DEFAULT_ACCESS_EXPORT_FILE As String = "_TennisLink.accdb"

Global Const BYE_WEEK_TEAM_NAME As String = ""
Global Const BYE_WEEK_TEAM_ID As String = "0"
Global Const BYE_FACILITY_ID As String = "0"
Global Const INVALID_FACILITY_ID As String = "-1"
Global Const MISSING_FACILITY_NAME As String = "TBD"

Global Const FACILITIES_TABLE_NAME As String = "Facilities"
Global Const FACILITIES_SHEET_NAME As String = "tFacilities"
Global Const FACILITIES_ID_COLUMN_NAME As String = "FacilitiesID"
Global Const FACILITIES_NAME_COLUMN_NAME As String = "FacilitiesName"
Global Const FACILITIES_ID_RANGE_NAME As String = FACILITIES_TABLE_NAME & "[" & FACILITIES_ID_COLUMN_NAME & "]"
Global Const FACILITIES_NAME_RANGE_NAME As String = FACILITIES_TABLE_NAME & "[" & FACILITIES_NAME_COLUMN_NAME & "]"

Global Const HEADER_TABLE_NAME As String = "Header"
Global Const HEADER_SHEET_NAME As String = "tHeader"
Global Const HEADER_START_DATE_COLUMN_NAME As String = "StartDate"
Global Const HEADER_END_DATE_COLUMN_NAME As String = "EndDate"
Global Const HEADER_START_DATE_RANGE_NAME As String = HEADER_TABLE_NAME & "[" & HEADER_START_DATE_COLUMN_NAME & "]"
Global Const HEADER_END_DATE_RANGE_NAME As String = HEADER_TABLE_NAME & "[" & HEADER_END_DATE_COLUMN_NAME & "]"

Global Const TEAMS_TABLE_NAME As String = "Teams"
Global Const TEAMS_SHEET_NAME As String = "tTeams"
Global Const TEAM_ID_COLUMN_NAME As String = "TeamID"
Global Const TEAM_NAME_COLUMN_NAME As String = "TeamName"
Global Const TEAM_ID_RANGE_NAME As String = TEAMS_TABLE_NAME & "[" & TEAM_ID_COLUMN_NAME & "]"
Global Const TEAM_NAME_RANGE_NAME As String = TEAMS_TABLE_NAME & "[" & TEAM_NAME_COLUMN_NAME & "]"

Global Const MATCHES_TABLE_NAME As String = "Matches"
Global Const MATCHES_SHEET_NAME As String = "tMatches"
Global Const MATCH_ID_COLUMN_NAME As String = "MatchID"
Global Const MATCH_ID_RANGE_NAME As String = MATCHES_TABLE_NAME & "[" & MATCH_ID_COLUMN_NAME & "]"

Global Const MATCH_DATE_COLUMN_NAME As String = "MatchDate"
Global Const MATCH_TIME_COLUMN_NAME As String = "MatchTime"
Global Const MATCH_DATE_RANGE_NAME As String = MATCHES_TABLE_NAME & "[" & MATCH_DATE_COLUMN_NAME & "]"
Global Const MATCH_TIME_RANGE_NAME As String = MATCHES_TABLE_NAME & "[" & MATCH_TIME_COLUMN_NAME & "]"
Global Const MATCH_DATE_TIME_RANGE_NAME As String = MATCHES_TABLE_NAME & "[[" & MATCH_DATE_COLUMN_NAME & "]:[" & MATCH_TIME_COLUMN_NAME & "]]"

Global Const HOME_TEAM_ID_COLUMN_NAME As String = "HomeTeamID"
Global Const HOME_TEAM_NAME_COLUMN_NAME As String = "HomeTeamName"
Global Const HOME_TEAM_ID_RANGE_NAME As String = MATCHES_TABLE_NAME & "[" & HOME_TEAM_ID_COLUMN_NAME & "]"
Global Const HOME_TEAM_NAME_RANGE_NAME As String = MATCHES_TABLE_NAME & "[" & HOME_TEAM_NAME_COLUMN_NAME & "]"

Global Const VISITING_TEAM_ID_COLUMN_NAME As String = "VisitingTeamID"
Global Const VISITING_TEAM_NAME_COLUMN_NAME As String = "VisitingTeamName"
Global Const VISITING_TEAM_ID_RANGE_NAME As String = MATCHES_TABLE_NAME & "[" & VISITING_TEAM_ID_COLUMN_NAME & "]"
Global Const VISITING_TEAM_NAME_RANGE_NAME As String = MATCHES_TABLE_NAME & "[" & VISITING_TEAM_NAME_COLUMN_NAME & "]"

Global Const FACILITY_ID_COLUMN_NAME As String = "FacilityID"
Global Const FACILITY_NAME_COLUMN_NAME As String = "FacilityName"
Global Const FACILITY_ID_RANGE_NAME As String = MATCHES_TABLE_NAME & "[" & FACILITY_ID_COLUMN_NAME & "]"
Global Const FACILITY_NAME_RANGE_NAME As String = MATCHES_TABLE_NAME & "[" & FACILITY_NAME_COLUMN_NAME & "]"

Global Const MATCH_TEAM_INFO_RANGE_NAME As String = MATCHES_TABLE_NAME & "[[" & HOME_TEAM_NAME_COLUMN_NAME & _
    "]:[" & VISITING_TEAM_ID_COLUMN_NAME & "]]"

Function GetTeamID(teamName As String) As String
    GetTeamID = IndexMatch(teamName, Range(TEAM_ID_RANGE_NAME), Range(TEAM_NAME_RANGE_NAME), BYE_WEEK_TEAM_ID)
End Function

Function GetTeamName(teamID As String) As String
    GetTeamName = IndexMatch(teamID, Range(TEAM_NAME_RANGE_NAME), Range(TEAM_ID_RANGE_NAME), "N/A")
End Function

Function GetFacilityID(facilityName As String) As String
    GetFacilityID = IndexMatch(facilityName, Range(FACILITIES_ID_RANGE_NAME), Range(FACILITIES_NAME_RANGE_NAME), INVALID_FACILITY_ID)
End Function

Function GetFacilityName(facilityID As String) As String
    GetFacilityName = IndexMatchNumber(Val(facilityID), Range(FACILITIES_NAME_RANGE_NAME), Range(FACILITIES_ID_RANGE_NAME), "N/A")
End Function

Function IndexMatch(lookupValue As String, indexRange As Range, matchRange As Range, notFoundValue As String) As String
    'See http://eimagine.com/say-goodbye-to-vlookup-and-hello-to-index-match/
    On Error GoTo Err

    Dim matchResult As Double
    Const MATCH_TYPE_EXACT As Integer = 0  '0: Exact, -1: Nearest less than, 1: Nearest greater than
    
    matchResult = Application.WorksheetFunction.Match(lookupValue, matchRange, MATCH_TYPE_EXACT)
    IndexMatch = Application.WorksheetFunction.Index(indexRange, matchResult)
    
    Exit Function
    
Err:
    'Don't display error to user, the error is handled by returning notFoundValue and the
    'root cause may be discovered later in the validation routines.  Not found is Err.Number = 1004
    'DisplayError "IndexMatch", Err
    
    IndexMatch = notFoundValue
End Function

Function IndexMatchNumber(lookupValue As Double, indexRange As Range, matchRange As Range, notFoundValue As String) As String
    On Error GoTo Err

    Dim matchResult As Double
    Const MATCH_TYPE_EXACT As Integer = 0  '0: Exact, -1: Nearest less than, 1: Nearest greater than
    
    matchResult = Application.WorksheetFunction.Match(lookupValue, matchRange, MATCH_TYPE_EXACT)
    IndexMatchNumber = Application.WorksheetFunction.Index(indexRange, matchResult)
    
    Exit Function
    
Err:
    
    IndexMatchNumber = notFoundValue
End Function

Function VLookup(lookupValue As String, tableLookup As Range, colOffset As Integer, exactMatch As Boolean) As String
    On Error GoTo Err
    
    VLookup = Application.WorksheetFunction.VLookup(lookupValue, tableLookup, colOffset, exactMatch)
       
    Exit Function
    
Err:
    If Err.Number = 1004 Then
        MsgBox "VLookup Error: " & "Lookup '" & lookupValue & "' is missing from the lookup table", vbExclamation, ADDIN_NAME
    Else
        DisplayError "VLookup", Err
    End If
    
    VLookup = "Not Found"
End Function

Function GetAccessFilePath(InitialFileName As String) As String
    On Error GoTo Err
    
    With Application.FileDialog(msoFileDialogOpen)
        .Title = "Select a Tennis Link Access file"
        .InitialFileName = InitialFileName
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Microsoft Access", "*.accdb, *.mdb"
    End With

    If Application.FileDialog(msoFileDialogOpen).Show <> 0 Then
        GetAccessFilePath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
    Else
        GetAccessFilePath = vbNullString
    End If
    
    Exit Function
    
Err:
    MsgBox "GetAccessFilePath Error: " & Err.Description, vbExclamation, ADDIN_NAME
    
    GetAccessFilePath = vbNullString
End Function

Function ContainsExport(TheWorkbook As Excel.Workbook) As Boolean
    ContainsExport = SheetExists(HEADER_SHEET_NAME, TheWorkbook) Or _
        SheetExists(TEAMS_SHEET_NAME, TheWorkbook) Or _
        SheetExists(MATCHES_SHEET_NAME, TheWorkbook) Or _
        SheetExists(FACILITIES_SHEET_NAME, TheWorkbook)
End Function

Function SheetExists(sheetName As String, TheWorkbook As Excel.Workbook) As Boolean
    Dim testSheet As Excel.Worksheet
    
    On Error Resume Next
    
    Set testSheet = TheWorkbook.Sheets(sheetName)
    
    On Error GoTo 0
    
    SheetExists = Not testSheet Is Nothing
End Function

Function FolderFromPath(fullPath As String) As String
    'See http://vba-tutorial.com/parsing-a-file-string-into-path-filename-and-extension/ for more
     FolderFromPath = Left(fullPath, InStrRev(fullPath, Application.PathSeparator))
End Function

Sub DisplayError(FunctionName As String, ErrCode As ErrObject)
    MsgBox FunctionName & " Error: " & ErrCode.Description, vbExclamation, ADDIN_NAME
End Sub
