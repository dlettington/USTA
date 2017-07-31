Attribute VB_Name = "Import"
Option Explicit

'Callback for importButton onAction
Sub ImportMatchesFromAccess(control As IRibbonControl)
    USTAImport
End Sub

Sub USTAImport()
    Application.ScreenUpdating = False
    
    Dim AccessFilePath As String
    
    If ActiveWorkbook Is Nothing Then
        'MsgBox "You must open an Excel workbook before importing.", vbInformation, ADDIN_NAME
        'Exit Sub
        
        'Open a blank workbook
        Workbooks.Add
    End If
    
    If MsgBox("Import match information from Access?", vbYesNo, ADDIN_NAME) = vbYes Then
        'Check that an import has not already been done
        If ContainsExport(ActiveWorkbook) Then
            If MsgBox("This workbook already contains match information. Would you like to replace it?", vbYesNo, ADDIN_NAME) = vbYes Then
                'Delete the previously imported worksheets
                On Error Resume Next
            
                Application.DisplayAlerts = False
                
                Sheets(HEADER_SHEET_NAME).Delete
                Sheets(TEAMS_SHEET_NAME).Delete
                Sheets(MATCHES_SHEET_NAME).Delete
                Sheets(FACILITIES_SHEET_NAME).Delete
                
                Application.DisplayAlerts = True
            Else
                Exit Sub
            End If
        End If
    
        CurrentAccessFilePath = vbNullString
        AccessFilePath = GetAccessFilePath(DEFAULT_ACCESS_FOLDER)
        
        If Len(AccessFilePath) > 0 Then
            'User selected an access file, import and process it
            If Import(AccessFilePath) Then
                If AddTeamAndFacilityNames Then
                    'Success, save file path for use in export
                    'Consider persisting in case the user closes the Excel file
                    CurrentAccessFilePath = AccessFilePath
                End If
            End If
        End If
    End If
    
    Application.ScreenUpdating = True
End Sub

Function Import(AccessFilePath As String) As Boolean
    On Error GoTo Err
    
    Dim newSheet As Worksheet
        
    'DREWDREW - Need to check that we have an active workbook, and handle cases where sheet names alread exist
    
    'Since facilities information is not included in the Access export from Tennis Link, create it here
    If Not AddFacilitiesTable() Then
        Import = False
        
        Exit Function
    End If
    
    'Get the header table
    Set newSheet = ActiveWorkbook.Worksheets.Add
    
    With newSheet.ListObjects.Add(SourceType:=0, Source:=Array( _
        "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=" & AccessFilePath & ";Mo" _
        , _
        "de=Read;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password=""" _
        , _
        """;Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactio" _
        , _
        "ns=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't " _
        , _
        "Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=F" _
        , _
        "alse;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False" _
        , ""), Destination:=Range("$A$1")).QueryTable
        
        .CommandType = xlCmdTable
        .CommandText = Array(HEADER_SHEET_NAME)
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .SourceDataFile = AccessFilePath
        .Refresh BackgroundQuery:=False
        '.ListObject.DisplayName = HEADER_TABLE_NAME
        
        'Cannot change the display name directly since we are now opening the Access DB as read-only
        'Let the import into Excel create the named range with a default name (Table_ExternalData_nnn) and then rename it
        newSheet.ListObjects(.ListObject.DisplayName).name = HEADER_TABLE_NAME
    End With
    
    
    newSheet.name = HEADER_SHEET_NAME
    
    'Get the teams table
    Set newSheet = ActiveWorkbook.Worksheets.Add
    
    With newSheet.ListObjects.Add(SourceType:=0, Source:=Array( _
        "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=" & AccessFilePath & ";Mo" _
        , _
        "de=Read;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password=""" _
        , _
        """;Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactio" _
        , _
        "ns=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't " _
        , _
        "Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=F" _
        , _
        "alse;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False" _
        , ""), Destination:=Range("$A$1")).QueryTable
        
        .CommandType = xlCmdTable
        .CommandText = Array(TEAMS_SHEET_NAME)
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .SourceDataFile = AccessFilePath
        .Refresh BackgroundQuery:=False
        '.ListObject.DisplayName = TEAMS_TABLE_NAME
        
        'Cannot change the display name directly since we are now opening the Access DB as read-only
        'Let the import into Excel create the named range with a default name (Table_ExternalData_nnn) and then rename it
        newSheet.ListObjects(.ListObject.DisplayName).name = TEAMS_TABLE_NAME
    End With
    
    newSheet.name = TEAMS_SHEET_NAME
    
    'Get matches table
    Set newSheet = ActiveWorkbook.Worksheets.Add
    
    With newSheet.ListObjects.Add(SourceType:=0, Source:=Array( _
        "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=" & AccessFilePath & ";Mo" _
        , _
        "de=Read;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password=""" _
        , _
        """;Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactio" _
        , _
        "ns=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't " _
        , _
        "Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=F" _
        , _
        "alse;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False" _
        , ""), Destination:=Range("$A$1")).QueryTable
        
        .CommandType = xlCmdTable
        .CommandText = Array(MATCHES_SHEET_NAME)
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .SourceDataFile = AccessFilePath
        .Refresh BackgroundQuery:=False
        '.ListObject.DisplayName = MATCHES_TABLE_NAME
        
        'Cannot change the display name directly since we are now opening the Access DB as read-only
        'Let the import into Excel create the named range with a default name (Table_ExternalData_nnn) and then rename it
        newSheet.ListObjects(.ListObject.DisplayName).name = MATCHES_TABLE_NAME
    End With
    
    newSheet.name = MATCHES_SHEET_NAME
    
    'Explicitly set date and time columns to text as that is what Tennis Link Access file expects
    Range(MATCH_DATE_TIME_RANGE_NAME).NumberFormat = "@"
    
    'Make information easier to view
    Sheets(HEADER_SHEET_NAME).Columns.AutoFit
    Sheets(TEAMS_SHEET_NAME).Columns.AutoFit
    Sheets(MATCHES_SHEET_NAME).Columns.AutoFit
    Sheets(FACILITIES_SHEET_NAME).Columns.AutoFit
    
    Import = True
    
    Exit Function
    
Err:
    DisplayError "Import", Err
    
    Import = False
End Function

Function AddFacilitiesTable() As Boolean
    On Error GoTo Err
    '6.10.17 - Added 9 new facilities, updated one that was TBA with a valid facilities ID
    '6.11.17 - Added 1 new facility
    '6.19.17 - Added 1 new facility

    'Currently there are 52 facilities defined, update this array if there are changes in Tennis Link
    'Note: defining the names alphabetical order makes data validation friendlier
    'Array size is the 54 (number of facilities plus 1 (for TBD) plus 1 (for the header))
    'If the number of facilities changes, update the array dimensions (setting base-0 upper bound)
    'and the range for the table, as well as the population values (ID is number, name is string)
    Dim facilitiesArray(53, 1)
    Const FACILITIES_RANGE As String = "A1:B54"
    
    facilitiesArray(0, 0) = FACILITIES_ID_COLUMN_NAME
    facilitiesArray(0, 1) = FACILITIES_NAME_COLUMN_NAME
    facilitiesArray(1, 0) = 919359696
    facilitiesArray(1, 1) = "Balboa Tennis Club"
    facilitiesArray(2, 0) = 919899368
    facilitiesArray(2, 1) = "Barnes Tennis Center"
    facilitiesArray(3, 0) = 919364093
    facilitiesArray(3, 1) = "Bobby Riggs Tennis Club" 'Added 6.10.17
    facilitiesArray(4, 0) = 2010778698
    facilitiesArray(4, 1) = "BRENGLE TERRACE PARK"
    facilitiesArray(5, 0) = 2010005061
    facilitiesArray(5, 1) = "Carmel Valley Tennis" 'Added 6.11.17
    facilitiesArray(6, 0) = 2010924271
    facilitiesArray(6, 1) = "Chantemar HOA Tennis Courts" 'Added 6.10.17
    facilitiesArray(7, 0) = 2010905213
    facilitiesArray(7, 1) = "Coronado Cays" 'Added 6.10.17
    facilitiesArray(8, 0) = 919367768
    facilitiesArray(8, 1) = "Coronado Island Marriott Resort Tennis Club"
    facilitiesArray(9, 0) = 919359702
    facilitiesArray(9, 1) = "Coronado Tennis Association"
    facilitiesArray(10, 0) = 2010178251
    facilitiesArray(10, 1) = "Coronado Tennis Center"
    facilitiesArray(11, 0) = 2010881325
    facilitiesArray(11, 1) = "Del Cerro Tennis Club" 'Added 6.10.17
    facilitiesArray(12, 0) = 2010905226
    facilitiesArray(12, 1) = "Del Rayo Downs" 'Updated from TBA 6.10.17
    facilitiesArray(13, 0) = 922072780
    facilitiesArray(13, 1) = "Eccta At Helix High School"
    facilitiesArray(14, 0) = 920300711
    facilitiesArray(14, 1) = "El Camino Country Club"
    facilitiesArray(15, 0) = 920875890
    facilitiesArray(15, 1) = "Fairbanks"
    facilitiesArray(16, 0) = 2010917536
    facilitiesArray(16, 1) = "Fairbanks Community" 'Added 6.10.17
    facilitiesArray(17, 0) = 919363033
    facilitiesArray(17, 1) = "Fairbanks Ranch Country Club"
    facilitiesArray(18, 0) = 919363352
    facilitiesArray(18, 1) = "Fallbrook Tennis Club"
    facilitiesArray(19, 0) = 2010905222
    facilitiesArray(19, 1) = "Fit Athletic" 'Updated from TBA 6.10.17
    facilitiesArray(20, 0) = 2010778686
    facilitiesArray(20, 1) = "KIT CARSON PARK"
    facilitiesArray(21, 0) = 919359647
    facilitiesArray(21, 1) = "La Jolla Beach Tennis Club"
    facilitiesArray(22, 0) = 919359650
    facilitiesArray(22, 1) = "La Jolla Tennis Club"
    facilitiesArray(23, 0) = 919366662
    facilitiesArray(23, 1) = "Lake Murray Tennis Club"
    facilitiesArray(24, 0) = 919366906
    facilitiesArray(24, 1) = "Lomas Santa Fe Country Club"
    facilitiesArray(25, 0) = 921005446
    facilitiesArray(25, 1) = "Martin Luther King Park" 'Added 6.10.17
    facilitiesArray(26, 0) = 919370570
    facilitiesArray(26, 1) = "Morgan Run Resort and Club"
    facilitiesArray(27, 0) = 919371867
    facilitiesArray(27, 1) = "Mountain View Sports Racquet Association"
    facilitiesArray(28, 0) = 2010778692
    facilitiesArray(28, 1) = "OMNI LA COSTA RESORT AND SPA"
    facilitiesArray(29, 0) = 919363862
    facilitiesArray(29, 1) = "Pacific Beach Tennis Club"
    facilitiesArray(30, 0) = 2010905221
    facilitiesArray(30, 1) = "Park Hyatt Aviara"
    facilitiesArray(31, 0) = 919362208
    facilitiesArray(31, 1) = "Peninsula Tennis Club"
    facilitiesArray(32, 0) = 2005272799
    facilitiesArray(32, 1) = "Rancho Arbolitos Swim and Tennis Club"
    facilitiesArray(33, 0) = 919367280
    facilitiesArray(33, 1) = "Rancho Bernardo Community Tennis Center"
    facilitiesArray(34, 0) = 919363863
    facilitiesArray(34, 1) = "Rancho Bernardo Swim/tnns Club"
    facilitiesArray(35, 0) = 919363976
    facilitiesArray(35, 1) = "Rancho Penasquitos Tennis Center"
    facilitiesArray(36, 0) = 920950464
    facilitiesArray(36, 1) = "Rancho Santa Fe Tennis Club"
    facilitiesArray(37, 0) = 919369274
    facilitiesArray(37, 1) = "Rancho Valencia Resort"
    facilitiesArray(38, 0) = 920329947
    facilitiesArray(38, 1) = "Rb Westwood Club"
    facilitiesArray(39, 0) = 2004526466
    facilitiesArray(39, 1) = "Riviera Oaks Resort and Racquet Club" 'Added 6.10.17
    facilitiesArray(40, 0) = 919994863
    facilitiesArray(40, 1) = "San Diego Tennis and Racquet Club"
    facilitiesArray(41, 0) = 919359899
    facilitiesArray(41, 1) = "San Dieguito Tennis Club"
    facilitiesArray(42, 0) = 919363675
    facilitiesArray(42, 1) = "Scripps Ranch & Racquet Club"
    facilitiesArray(43, 0) = 922851229
    facilitiesArray(43, 1) = "Scripps Ranch High School"
    facilitiesArray(44, 0) = 2010924794
    facilitiesArray(44, 1) = "Scripps Trails Tennis Courts" 'Added 6.10.17
    facilitiesArray(45, 0) = 2010778697
    facilitiesArray(45, 1) = "STONERIDGE COUNTRY CLUB"
    facilitiesArray(46, 0) = 919435248
    facilitiesArray(46, 1) = "Surf & Turf Tennis Club"
    facilitiesArray(47, 0) = 922841443
    facilitiesArray(47, 1) = "The Santaluz Club"
    facilitiesArray(48, 0) = 2010482287
    facilitiesArray(48, 1) = "Tierrasanta Tennis Club"
    facilitiesArray(49, 0) = 919359779
    facilitiesArray(49, 1) = "University City Racquet Club"
    facilitiesArray(50, 0) = 2010778710
    facilitiesArray(50, 1) = "VALLEY CENTER TENNIS ADAMS PARK"
    facilitiesArray(51, 0) = 919363681
    facilitiesArray(51, 1) = "Vista Tennis Club" 'Added 6.10.17
    facilitiesArray(52, 0) = 919359783
    facilitiesArray(52, 1) = "Winner's Tennis Club"
    facilitiesArray(53, 0) = 0
    facilitiesArray(53, 1) = MISSING_FACILITY_NAME

    Dim facilitiesSheet As Worksheet
    
    Set facilitiesSheet = ActiveWorkbook.Sheets.Add
    facilitiesSheet.name = FACILITIES_SHEET_NAME
    
    facilitiesSheet.Range(FACILITIES_RANGE) = facilitiesArray
        
    facilitiesSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=facilitiesSheet.Range(FACILITIES_RANGE), _
        XlListObjectHasHeaders:=xlYes).name = FACILITIES_TABLE_NAME

    AddFacilitiesTable = True
    
    Exit Function
    
Err:
    DisplayError "AddFacilitiesTable", Err
    
    AddFacilitiesTable = False
End Function

'Add columns useful for making modifications to the matches table but are not
'exported from Tennis Link into Access
Function InsertColumns() As Boolean
    On Error GoTo Err
    
    Dim matchesTable As ListObject
    Dim newColNum As Integer
    
    Set matchesTable = Sheets(MATCHES_SHEET_NAME).ListObjects(MATCHES_TABLE_NAME)
    
    newColNum = Range(HOME_TEAM_ID_RANGE_NAME).Column
    matchesTable.ListColumns.Add(newColNum).name = HOME_TEAM_NAME_COLUMN_NAME
    
    newColNum = Range(VISITING_TEAM_ID_RANGE_NAME).Column
    matchesTable.ListColumns.Add(newColNum).name = VISITING_TEAM_NAME_COLUMN_NAME
    
    newColNum = Range(FACILITY_ID_RANGE_NAME).Column
    matchesTable.ListColumns.Add(newColNum).name = FACILITY_NAME_COLUMN_NAME
    
    InsertColumns = True
    
    Exit Function
    
Err:
    DisplayError "InsertColumns", Err
    
    InsertColumns = False
End Function

Function AddTeamAndFacilityNames() As Boolean
    On Error GoTo Err
    
    'Create the new columns in the match table
    If InsertColumns() = False Then
        AddTeamAndFacilityNames = False
        
        Exit Function
    End If
    
    'Get the facility names
    If GetMatchFacilityNames() = False Then
        AddTeamAndFacilityNames = False
        
        Exit Function
    End If
    
    'Get the home and visiting team names
    If GetMatchTeamNames() = False Then
        AddTeamAndFacilityNames = False
        
        Exit Function
    End If

    AddTeamAndFacilityNames = True
    
    Exit Function
    
Err:
    DisplayError "AddTeamAndFacilityNames", Err
    
    AddTeamAndFacilityNames = True
End Function

Function GetMatchTeamNames() As Boolean
    On Error GoTo Err

    Dim teamID As String
    Dim teamName As String
    
    'Get home team names from teams lookup table
    Dim i As Integer
    
    For i = 1 To Range(HOME_TEAM_ID_RANGE_NAME).Rows.Count
        'First do the home team
        teamID = Range(HOME_TEAM_ID_RANGE_NAME).Value2(i, 1)
        
        If teamID <> BYE_WEEK_TEAM_ID Then
            'TeamName = VLookup(TeamID, Range(TEAMS_TABLE_NAME), 4, False)
            teamName = GetTeamName(teamID)
        Else
            teamName = BYE_WEEK_TEAM_NAME
        End If
        
        Range(HOME_TEAM_NAME_RANGE_NAME).Cells(i, 1) = teamName
        
        'Then do the visiting team
        teamID = Range(VISITING_TEAM_ID_RANGE_NAME).Value2(i, 1)
        
        If teamID <> BYE_WEEK_TEAM_ID Then
            teamName = GetTeamName(teamID)
        Else
            teamName = BYE_WEEK_TEAM_NAME
        End If
        
        Range(VISITING_TEAM_NAME_RANGE_NAME).Cells(i, 1) = teamName
        
    Next i
    
    Sheets(MATCHES_SHEET_NAME).Columns.AutoFit
    Range(HOME_TEAM_ID_RANGE_NAME).Columns.EntireColumn.Hidden = True
    Range(VISITING_TEAM_ID_RANGE_NAME).Columns.EntireColumn.Hidden = True
    Range(FACILITY_ID_RANGE_NAME).Columns.EntireColumn.Hidden = True
    
    'Add data validation for team names
    DefineNameDataValidation
        
    GetMatchTeamNames = True
    
    Exit Function
    
Err:
    DisplayError "GetMatchTeamNames", Err
    
    GetMatchTeamNames = False
End Function

Function GetMatchFacilityNames() As Boolean
    On Error GoTo Err

    Dim facilityID As String
    Dim facilityName As String
       
    'Get home team names from teams lookup table
    Dim i As Integer
    
    For i = 1 To Range(FACILITY_ID_RANGE_NAME).Rows.Count
        'First do the home team
        facilityID = Range(FACILITY_ID_RANGE_NAME).Value2(i, 1)
        
        If facilityID <> BYE_FACILITY_ID Then
            facilityName = GetFacilityName(facilityID)
        Else
            facilityName = MISSING_FACILITY_NAME
        End If
        
        Range(FACILITY_NAME_RANGE_NAME).Cells(i, 1) = facilityName
    Next i
    
    Sheets(MATCHES_SHEET_NAME).Columns.AutoFit
    Range(HOME_TEAM_ID_RANGE_NAME).Columns.EntireColumn.Hidden = True
    Range(VISITING_TEAM_ID_RANGE_NAME).Columns.EntireColumn.Hidden = True
    Range(FACILITY_ID_RANGE_NAME).Columns.EntireColumn.Hidden = True
    
    'Add data validation for team names
    DefineNameDataValidation
        
    GetMatchFacilityNames = True
    
    Exit Function
    
Err:
    DisplayError "GetMatchFacilityNames", Err
    
    GetMatchFacilityNames = False
End Function

Sub DefineNameDataValidation()
    On Error GoTo Err

    With Range(HOME_TEAM_NAME_RANGE_NAME).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
            Formula1:="='" & TEAMS_SHEET_NAME & "'!" & Range(TEAM_NAME_RANGE_NAME).Address
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ADDIN_NAME
        .InputMessage = ""
        .ErrorMessage = "Please enter a valid home team name"
        .ShowInput = False
        .ShowError = True
    End With

    With Range(VISITING_TEAM_NAME_RANGE_NAME).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
            Formula1:="='" & TEAMS_SHEET_NAME & "'!" & Range(TEAM_NAME_RANGE_NAME).Address
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ADDIN_NAME
        .InputMessage = ""
        .ErrorMessage = "Please enter a valid visiting team name"
        .ShowInput = False
        .ShowError = True
    End With
    
    With Range(FACILITY_NAME_RANGE_NAME).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
            Formula1:="='" & FACILITIES_SHEET_NAME & "'!" & Range(FACILITIES_NAME_RANGE_NAME).Address
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ADDIN_NAME
        .InputMessage = ""
        .ErrorMessage = "Please enter a valid facility name"
        .ShowInput = False
        .ShowError = True
    End With
    
    Exit Sub
    
Err:
    DisplayError "DefineNameDataValidation", Err
End Sub
