Attribute VB_Name = "Facilities"
Option Explicit

Const USTA_NUMBER_DELIM As String = "USTA Number: "
Const FACILITY_HEADER As String = "Facility"
Const CONTACT_HEADER As String = "Contact Information:"
Const PHONE_HEADER As String = "Phone:"
Const EMAIL_HEADER As String = "E-Mail:"

'Header in Tennis Link export; e.g. "Dates: All Sundays from 06/16/2017"
'Having problems with inconsistent Find behavior. Removed spaces in hopes
'of resolving the issue. More likely to get a false positive but should be okay.
Const DATES_HEADER As String = "Dates: All "
Const DATES_HEADER_FIND As String = "Dates:"

'Column headers in the usage table (in order), dates (variable number) are between Start Time and Home
Const TEAM_USAGE_HEADER As String = "Team Name"
Const FLIGHT_USAGE_HEADER As String = "Sub Flight/Gender"
Const START_TIME_USAGE_HEADER As String = "Normal Start Time"
Const START_TIME_USAGE_HEADER_REPLACE As String = "Start Time"
Const NUMBER_COURTS_USAGE_HEADER As String = "Normal # Courts Required"
Const NUMBER_COURTS_USAGE_HEADER_REPLACE As String = "Courts"
Const HOME_USAGE_HEADER As String = "Home"
Const AWAY_USAGE_HEADER As String = "Away"
Const NEUTRAL_USAGE_HEADER As String = "Neutral"
Const BYE_USAGE_HEADER As String = "Bye"

Const TABLE_FORMAT As String = "TableStyleMedium9"  'Originally was TableStyleMedium2

'Three collections holds the USTA Number to facilities worksheet name mapping
Dim facilitiesSheetMap As Collection  'USTA number is the key, sheet name is the item
Dim facilitiesNameMap As Collection 'USTA number is the key, facility name is the item
Dim facilitiesNumberMap As Collection 'Sheet name is the key, USTA number is the item
Const FACILITIES_MAP_SHEET_NAME As String = "_FacilitiesMap_"

Const USTA_EMAIL_NAME As String = "San Diego USTA"
Const USTA_EMAIL_ADDRESS As String = "sandiegousta@gmail.com"
Const USTA_EMAIL_FROM As String = """" & USTA_EMAIL_NAME & """ <" & USTA_EMAIL_ADDRESS & ">"

'Workbook containing information about each facility
Const FACILITY_INFO_PATH As String = "C:\Users\Drew\Documents\USTA\Excel Add-In\Content\FacilityInfo.xlsm"
Const FACILITY_INFO_SHEET_NAME As String = "FacilityInfo"

'Table headers and column numbers in the facility information workbook
Const USTA_NUMBER_COLUMN As String = "USTA Number"
Const FACILITY_NAME_COLUMN As String = "Facility Name"
Const FACILITY_EMAIL_COLUMN As String = "Facility Email"
Const FACILITY_PHONE_COLUMN As String = "Facility Phone"
Const FACILITY_ADDRESS_COLUMN As String = "Facility Address"
Const FACILITY_NOTES_COLUMN As String = "Notes"


'Holds the number of worksheets in the active workbook at the start of processing (used in alpha sorting)
Dim numExistingSheets As Integer


'Callback for usageButton onAction
Sub ImportFacilitiesUsage(control As IRibbonControl)
    On Error GoTo Err
    
    If ActiveWorkbook Is Nothing Then
        'Open a blank workbook.  Must be done before setting calculation to manual
        'so do it here rather than in ProcessFacilities()
        Workbooks.Add
    End If
    
    numExistingSheets = ActiveWorkbook.Sheets.Count
      
    Dim currentSheet As Worksheet
    Set currentSheet = ActiveWorkbook.ActiveSheet
    
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Importing USTA Facilities Usage Data..."
    Application.ScreenUpdating = False
    
    'Process the facilities usage data for all facilities for all days of the week
    ProcessFacilities
    
    'Return the user to where he was before adding the facilities worksheets
    currentSheet.Activate

Err:
    'Always fall through so clean up is done
    DeleteAllFacilitiesSheets  'Cleanup
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = vbNullString
    
    'Message is for debugging only, user should have already been notified
    'DisplayError "ImportFacilitiesUsage", Err
End Sub

'Callback for emailButton and emailSplitButton onAction
Sub EmailAllFacilities(control As IRibbonControl)
    On Error Resume Next
    
    If Not ActiveWorkbook Is Nothing Then
        EmailFacilities
    End If
End Sub

'Callback for emailSplitMenuSendCurrent onAction
Sub EmailCurrentFacility(control As IRibbonControl)
    On Error Resume Next
    'MsgBox "Not implemented yet Debra!", Title:=ADDIN_NAME
    
    If Not ActiveWorkbook.ActiveSheet Is Nothing Then
        EmailFacility ActiveWorkbook.ActiveSheet.name, False
    End If
End Sub


Sub UpdateStatusBar(statusMessage As String)
    Dim restoreState As Boolean
    restoreState = Application.ScreenUpdating
    
    Application.ScreenUpdating = True
    
    Application.StatusBar = statusMessage
    
    Application.ScreenUpdating = restoreState
End Sub

Sub ProcessFacilities()
    On Error GoTo Err
    
    'Initialize globals
    Set facilitiesSheetMap = New Collection
    Set facilitiesNameMap = New Collection
    Set facilitiesNumberMap = New Collection
    
    'Get facilities report into clipboard, then proceed
    Dim res As VbMsgBoxResult
    
    Dim day As Integer  '0:Sunday - 6:Saturday
    day = 0  'Start on Sunday
    
    Dim daysToProcess As Integer  'For Debugging
    daysToProcess = 7  'Normal processing, all seven days of the week, change to debug
    
    'Get the facilities for each day of the week using copy/paste
    Do
        UpdateStatusBar "Importing usage data for " & DayOfWeekString(day) & "s..."

        'Get facilities report into clipboard, then proceed
        res = MsgBox("Copy " & DayOfWeekString(day) & " Facilities Usage Report data from TennisLink to the clipboard." & _
            vbNewLine & vbNewLine & _
            "Once you have copied the data, select Ok to continue.", _
            vbOKCancel, "Facilities Usage Report [" & DayOfWeekString(day) & "]")
            
        If res <> vbOK Then
            'User cancelled, we are done
            Exit Sub
        End If

        If GetFacilitiesUsageDay(DayOfWeekString(day)) Then
            'Success, process the next day
            day = day + 1
        Else
            'Failure, ask the user to retry or quitt
            res = MsgBox("Retry copying " & DayOfWeekString(day) & " Facilities Usage Report data?", _
                vbYesNo, "Facilities Usage Report [" & DayOfWeekString(day) & "]")
                
            If res = vbNo Then
                'User doesn't want to retry, we are done
                Exit Sub
            End If
        End If
    Loop While day < daysToProcess
    
    'Once we have the facilities for all the days of the week for all facilities
    'Loop through the days worksheets and put the usage data for each facility on its own worksheet
    For day = 0 To daysToProcess - 1
        UpdateStatusBar "Formatting usage data for " & DayOfWeekString(day) & "s..."
        
        If Not ProcessFacilityDay(DayOfWeekString(day)) Then
            'TODO: Continue or stop
            'For now, stop
            Exit Sub
        End If
    Next day
    
    'Store the mapping for possible later use
    PersistFacilitiesMap
    
    'All done, try and make printing more painless
    OrientForPrinting
    
    Exit Sub
    
Err:
    DisplayError "ProcessFacilities", Err
End Sub

Sub OrientForPrinting()
    On Error Resume Next
    
    'Helps with printing - landscape and all on one page
    Dim sheetName As Variant
    Dim facilitySheet As Worksheet
    
    For Each sheetName In facilitiesSheetMap
        Set facilitySheet = ActiveWorkbook.Sheets(sheetName)
        
        facilitySheet.PageSetup.Orientation = xlLandscape
        facilitySheet.PageSetup.FitToPagesWide = 1
        'facilitySheet.PageSetup.FitToPagesTall = 1
    Next sheetName
End Sub

Function PasteUsageData(facilitiesSheet As Worksheet) As Boolean
    On Error GoTo Err
    
    'Paste starting at A1
    facilitiesSheet.Range("A1").Select
    facilitiesSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:=False, NoHTMLFormatting:=True
    
    PasteUsageData = True
    
    Exit Function
    
Err:
    PasteUsageData = False
End Function

Sub DeleteAllFacilitiesSheets()
    On Error Resume Next
    
    Dim day As Integer
    
    For day = 0 To 6
        DeleteFacilitiesSheet DayOfWeekSheetName(DayOfWeekString(day))
    Next day

End Sub

Sub DeleteFacilitiesSheet(facilitiesSheetName As String)
    On Error GoTo Err
    
    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets(facilitiesSheetName).Delete
    Application.DisplayAlerts = True
        
    Exit Sub
    
Err:
    'Error display is for debugging only
    'DisplayError "DeleteFacilitiesSheet", Err
End Sub

Function DayOfWeekString(dayNumber As Integer) As String
    DayOfWeekString = "InvalidDayNumber"
    
    Select Case dayNumber
        Case 0
            DayOfWeekString = "Saturday"
        Case 1
            DayOfWeekString = "Sunday"
        Case 2
            DayOfWeekString = "Monday"
        Case 3
            DayOfWeekString = "Tuesday"
        Case 4
            DayOfWeekString = "Wednesday"
        Case 5
            DayOfWeekString = "Thursday"
        Case 6
            DayOfWeekString = "Friday"
    End Select
End Function

Function DayOfWeekSheetName(dayName As String) As String
    DayOfWeekSheetName = "InvalidWorksheetName"
    
    DayOfWeekSheetName = "_" & dayName & "Usage"
End Function

'Get the facilities usage for the specified day from the clipboard and add it to the day worksheet
Function GetFacilitiesUsageDay(dayOfWeek As String) As Boolean
    On Error GoTo Err
    
    GetFacilitiesUsageDay = True
    
    Dim newSheet As Worksheet
    Set newSheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    
    newSheet.name = DayOfWeekSheetName(dayOfWeek)
    
    'Paste report data into new worksheet
    If Not PasteUsageData(newSheet) Then
        MsgBox "Clipboard does not contain facilities usage data.", Title:=ADDIN_NAME
        
        'We are not going to use the new worksheet so delete it
        DeleteFacilitiesSheet newSheet.name
        
        GetFacilitiesUsageDay = False
        Exit Function
    End If
    
    'Confirm we pasted the correct data
    'Look for a header identifying the day of the week for the paste; e.g. "Dates: All Sundays from 06/15/2017"
    'Due to how different browsers create the copy, the header will not be in a fixed location
    'Since the header prefix is unique enough a simple search should suffice to find it
    Dim dayRange As Range
    Set dayRange = newSheet.Cells.Find(DATES_HEADER_FIND, MatchCase:=True, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, SearchFormat:=False, LookIn:=xlValues, LookAt:=xlPart)
    
    If dayRange Is Nothing Then
        MsgBox "Clipboard does not contain facilities usage data.", Title:=ADDIN_NAME
        
        'We are not going to use the new worksheet so delete it
        DeleteFacilitiesSheet newSheet.name
        
        GetFacilitiesUsageDay = False
        Exit Function
    Else
        If InStr(dayRange.Value, dayOfWeek) <> (Len(DATES_HEADER) + 1) Then
            MsgBox "Clipboard does not contain facilities usage data for " & dayOfWeek & ".", Title:=ADDIN_NAME
        
            'We are not going to use the new worksheet so delete it
            DeleteFacilitiesSheet newSheet.name
        
            GetFacilitiesUsageDay = False
            Exit Function
        End If
    End If
    
    'Have a valid paste from Tennis Link, hide the raw usage data worksheet
    'newSheet.Visible = xlSheetHidden  'No longer need this since we delete the sheet at the end of processing
    
    Exit Function
    
Err:
    GetFacilitiesUsageDay = False
    
    DisplayError "GetFacilitiesUsageDay", Err
End Function

'Extract the facilities information from the specified day worksheet and add it to each facility worksheet
Function ProcessFacilityDay(dayOfWeek As String) As Boolean
    On Error GoTo Err
    
    ProcessFacilityDay = True
    
    'Get the worksheet facilities worksheet for the specified day of the week
    Dim newSheet As Worksheet
    Set newSheet = ActiveWorkbook.Sheets(DayOfWeekSheetName(dayOfWeek))
    
    'Loop through all the facilities reported for this specified day of the week
    Dim currFacility As Range
    Dim firstFacility As Range
    Dim allFacilities As Range
    
    'Since Excel cannot have nested finds, find get all the facilities for the day
    'in one loop, then loop through them and do the processing (which also uses find)
    With newSheet.Cells
        'Set currFacility = .Find(USTA_NUMBER_DELIM, SearchDirection:=xlNext, MatchCase:=True, LookIn:=xlValues, LookAt:=xlText)
        Set currFacility = .Find(USTA_NUMBER_DELIM, MatchCase:=True, SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, SearchFormat:=False, LookIn:=xlValues, LookAt:=xlPart)
            
        Set firstFacility = currFacility
        
        Set allFacilities = currFacility
               
        If Not currFacility Is Nothing Then
            Do
                'Get the location of each facility's usage data for the specified day of the week
                Set currFacility = .FindNext(currFacility)
                
                If Not currFacility Is Nothing Then
                    Set allFacilities = Union(allFacilities, currFacility)
                End If
                
            'Loop using search find next until we arrive at the first facility again
            Loop While Not currFacility Is Nothing And currFacility.Address <> firstFacility.Address
        End If
    End With
    
    'Loop through each facility with a data usage table for the specified day
    For Each currFacility In allFacilities.Cells
        'Get each facility's information for the specified day of the week and
        'add it to the facility's worksheet
        ProcessFacility currFacility, dayOfWeek
    Next currFacility
    
    Exit Function
    
Err:
    ProcessFacilityDay = False
    
    DisplayError "ProcessFacilityDay", Err
End Function

'Take an individual facility's information and add it to the facility's worksheet
Sub ProcessFacility(facilityRange As Range, dayOfWeek As String)
    On Error GoTo Err
    
    If facilityRange Is Nothing Then Exit Sub
    
    'If the facility's worksheet doesn't exist yet, create it
    'Get the facility's unique USTA number
    Dim ustaNumber As String
    Dim facilityName As String
    
    GetFacilityAndNumber facilityRange.Value, facilityName, ustaNumber

    Dim facilitySheet As Worksheet
    Dim facilitySheetName As String
    
    facilitySheetName = GetFacilitySheetName(ustaNumber)
    
    Dim newFacility As Boolean
    
    If Len(facilitySheetName) > 0 Then
        'Facility's worksheet already exists
        Set facilitySheet = ActiveWorkbook.Sheets(facilitySheetName)
        
        newFacility = False
    Else
        'Create a new worksheet and add it to the map
        Set facilitySheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        
        'Name the new worksheet using the facility name
        facilitySheetName = GetValidFacilitySheetName(facilityName)
        facilitySheet.name = facilitySheetName
        
        'Add the USTA number/worksheet name mappings
        facilitiesSheetMap.Add facilitySheetName, ustaNumber
        facilitiesNameMap.Add facilityName, ustaNumber
        facilitiesNumberMap.Add ustaNumber, facilitySheetName

        'Order the facilities worksheets alphabetically
        SheetAlphaOrder facilitySheet
        
        newFacility = True
    End If

    Dim pasteAddress As String
    Dim facilityDataRange As Range

    If newFacility Then
        'This is a new facility worksheet, copy the header and the usage table
        
        facilityRange.CurrentRegion.Copy
        facilitySheet.Paste facilitySheet.Range("A1")  'Paste starting in top left cell
        
        'Data range includes header and usage table
        Set facilityDataRange = facilitySheet.Range("A1").CurrentRegion
        
        'Header rows are variable so find the usage table by searching for the first column header
        Dim headerRangeNew As Range
        Set headerRangeNew = facilitySheet.Range("A1").CurrentRegion.Find(TEAM_USAGE_HEADER, MatchCase:=True)
        
        'Add a blank row and a row to insert the day of the week header
        headerRangeNew.EntireRow.Insert
        headerRangeNew.EntireRow.Insert
        
        With headerRangeNew.Offset(-1)
            .Value = dayOfWeek
            'With .Interior
            '    'Add a gray background so the day header stands out
            '    .Pattern = xlSolid
            '    .PatternColorIndex = xlAutomatic
            '    .ThemeColor = xlThemeColorDark1
            '    .TintAndShade = -0.14996795556505
            '    .PatternTintAndShade = 0
            'End With
        End With
    Else
        'Existing facility, adding a new day's usage table (no header)
        
        'Add a blank row, a day header, and then paste only the usage table
        pasteAddress = facilitySheet.Cells(Rows.Count, 1).End(xlUp).Offset(3, 0).Address
        facilitySheet.Range(pasteAddress).Value = "paste here"
        
        'Adding another day of data to an existing facility
        'Add a blank row and then copy only the usage table
        Dim facilityRangeUsageTable As Range
        
        'Trim off the header rows, leaving only the usage table
        'Header rows are variable so find the usage table by searching for the first column header
        Dim headerRangeExisting As Range
        Set headerRangeExisting = facilityRange.CurrentRegion.Find(TEAM_USAGE_HEADER, MatchCase:=True)

        'Usage table starts at the first column header and ends at the last cell in the range
        Set facilityRangeUsageTable = Range(headerRangeExisting, facilityRange.CurrentRegion.Cells(facilityRange.CurrentRegion.Cells.Count))

        facilityRangeUsageTable.Copy
        facilitySheet.Paste facilitySheet.Range(pasteAddress)
        
        Set facilityDataRange = facilitySheet.Range(pasteAddress).CurrentRegion
        
        'Add the header once the paste is done so we could use CurrentRegion to select the paste
        'Insert a day of the week header before pasting the usage table
        facilitySheet.Range(pasteAddress).Offset(-1).Value = dayOfWeek
    End If
    
    'Deselect the pasted range by selecting the top-left cell
    'facilitySheet.Range("A1").Select
    Application.CutCopyMode = False
    
    FormatFacilityRange facilitySheet, facilityName, ustaNumber, newFacility, dayOfWeek, facilityDataRange
    
    Exit Sub

Err:
    DisplayError "ProcessFacility", Err
End Sub

Sub GetFacilityAndNumber(mergedInfo As String, ByRef facilityName As String, ByRef ustaNumber As String)
    On Error GoTo Err
    
    Dim pos As Integer
    
    pos = InStr(mergedInfo, USTA_NUMBER_DELIM)
    
    If pos > 0 Then
        facilityName = Left(mergedInfo, pos - 1)
        ustaNumber = Right(mergedInfo, Len(mergedInfo) - pos - Len(USTA_NUMBER_DELIM) + 1)
    End If
    
    Exit Sub

Err:
    DisplayError "GetFacilityAndNumber", Err
End Sub

Function GetFacilitySheetName(ustaNumber As String) As String
    On Error GoTo Err
    
    Dim useMap As Boolean
    useMap = True
    
    If facilitiesSheetMap Is Nothing Then
        useMap = False
    ElseIf facilitiesSheetMap.Count = 0 Then
        useMap = False
    End If
    
    If useMap Then
        GetFacilitySheetName = facilitiesSheetMap.Item(ustaNumber)
    Else
        'Not in memory, use the persisted map
        Dim mapRange As Range
        Set mapRange = ActiveWorkbook.Sheets(FACILITIES_MAP_SHEET_NAME).UsedRange
    
        Dim i As Integer
    
        'Map has a header so start at row 2
        For i = 2 To mapRange.Rows.Count - 1
            If mapRange.Rows(i).Cells(1).Value = ustaNumber Then
                GetFacilitySheetName = mapRange.Rows(i).Cells(2).Value
                
                Exit Function
            End If
        Next i
        
        'TODO: Use Vlookup instead of a search loop
        'GetFacilitySheetName = Application.WorksheetFunction.VLookup(facilitySheetName, mapRange, 2, False)
        
        GetFacilitySheetName = vbNullString
    End If
    
    Exit Function

Err:
    GetFacilitySheetName = vbNullString
End Function

Function GetFacilityUSTANumber(facilitySheetName As String) As String
    On Error GoTo Err
    
    Dim useMap As Boolean
    useMap = True
    
    If facilitiesNumberMap Is Nothing Then
        useMap = False
    ElseIf facilitiesNumberMap.Count = 0 Then
        useMap = False
    End If
    
    If useMap Then
        GetFacilityUSTANumber = facilitiesNumberMap.Item(facilitySheetName)
    Else
        'Not in memory, use the persisted map
        Dim mapRange As Range
        Set mapRange = ActiveWorkbook.Sheets(FACILITIES_MAP_SHEET_NAME).UsedRange
    
        Dim i As Integer
    
        'Map has a header so start at row 2
        For i = 2 To mapRange.Rows.Count
            If mapRange.Rows(i).Cells(2).Value = facilitySheetName Then
                GetFacilityUSTANumber = mapRange.Rows(i).Cells(1).Value
                
                Exit Function
            End If
        Next i
        
        GetFacilityUSTANumber = vbNullString
    End If
    
    Exit Function
    
Err:
    GetFacilityUSTANumber = vbNullString
End Function

Sub FormatFacilityRange(facilitySheet As Worksheet, facilityName As String, ustaNumber As String, _
                        formatHeader As Boolean, dayOfWeek As String, dataRange As Range)
    On Error GoTo Err
    
    Dim dataUsageTable As ListObject
    
    If formatHeader Then
        Dim headerRange As Range  'Cell containing the specified header
        
        'Separate merged cell containing facility name and USTA number for readbility
        Set headerRange = dataRange.Find(FACILITY_HEADER, MatchCase:=True)
        headerRange.Offset(0, 1).Value = facilityName
        
        'Display raw number so auto-fitting the column looks better
        'Prefix with single quote so it is handled as text and not scientific
        headerRange.Offset(0, 2).Value = ustaNumber
        
        'Set the facility's email as a hyperlink
        Dim facilityEmail As String
        
        'Find behaves strangely so using a different method
        'Email is the second column in the last row of the header block
        Set headerRange = dataRange.Find(EMAIL_HEADER, MatchCase:=True)
        facilityEmail = headerRange.Offset(0, 1).Value
        
        If InStr(facilityEmail, "@") > 1 Then
            headerRange.Offset(0, 1).Hyperlinks.Add Anchor:=headerRange.Offset(0, 1), Address:="mailto:" & facilityEmail & "?subject=USTA Facilities Report", TextToDisplay:=facilityEmail
        End If
        
        'Give the header a friendly range name for lookup.  Note, do this after unmerging facility name and USTA Number
        Dim headerInfoRange As Range
        Set headerInfoRange = facilitySheet.Range("A1").CurrentRegion
        
        Dim facilityHeaderName As String
        facilityHeaderName = GetValidFacilityTableName(facilityName, "Info")
    
        If Len(facilityHeaderName) > 0 Then
            headerInfoRange.name = facilityHeaderName
        End If
    
        'Format table portion as an Excel table
        'Set headerRange = dataRange.Find(EMAIL_HEADER, MatchCase:=True)
        Set headerRange = dataRange.Find(TEAM_USAGE_HEADER, MatchCase:=True)
        Set dataUsageTable = facilitySheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range(headerRange, facilitySheet.Cells.SpecialCells(xlCellTypeLastCell)), XlListObjectHasHeaders:=xlYes, TableStyleName:=TABLE_FORMAT)
    Else
        'No header, just a data table and its range is specified in the dataRange parameter
        Set dataUsageTable = facilitySheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=dataRange, XlListObjectHasHeaders:=xlYes, TableStyleName:=TABLE_FORMAT)
    End If
    
    'If we want to hide filter buttons (could be useful for copy/paste into final email)
    dataUsageTable.ShowAutoFilterDropDown = False
    
    'Before auto-fit, replace some long column header names with shorter versions for readability
    With dataUsageTable.Range.Cells
        .Replace What:=START_TIME_USAGE_HEADER, Replacement:=START_TIME_USAGE_HEADER_REPLACE, LookAt:=xlWhole, MatchCase:=True
        .Replace What:=NUMBER_COURTS_USAGE_HEADER, Replacement:=NUMBER_COURTS_USAGE_HEADER_REPLACE, LookAt:=xlWhole, MatchCase:=True
    End With
    
    'Make dates more readable - Export have dates of the format two digit month, space, one or two digit day; e.g. '07 8'
    'Replace space with / and Excel will format as short data; e.g. 8-Jul unless prefixed with a single quote
    Dim datesRange As Range
    Set datesRange = dataUsageTable.Range.Find(TEAM_USAGE_HEADER, LookAt:=xlWhole, MatchCase:=True)
        
    Dim i As Integer
    Dim dateVal As Variant
        
    'Skip over Team, Flight, Start, and Courts columns, then look until finding Home column
    'Everything in between is a match date
    For i = 4 To 30
        dateVal = datesRange.Offset(0, i)
            
        If dateVal = HOME_USAGE_HEADER Then
            Exit For  'Found the Home column, done with dates
        Else
            dateVal = Replace(dateVal, " ", "/", Compare:=vbTextCompare)
                
            'Remove any leading zeros from the month (Tennis Link does not prepend zero with single digit days)
            If Left(dateVal, 1) = "0" Then
                dateVal = Right(dateVal, Len(dateVal) - 1)
            End If
                
            'Single quote prefix displays like 7/8, otherwise like 8-Jul
            datesRange.Offset(0, i).Value = "'" & dateVal
        End If
    Next i
    
    Dim pos As Integer
    Dim teamsVal As Variant
    
    'Strip USTA numbers from teams to compress the Team column; e.g. SD-4.0F SDTRC/Lettington(6518241842) -> SD-4.0F SDTRC/Lettington
    For i = 2 To dataUsageTable.Range.Rows.Count
        teamsVal = dataUsageTable.Range.Rows(i).Cells(1).Value
        
        pos = InStr(teamsVal, "(")
        If (pos > 0) Then
            dataUsageTable.Range.Rows(i).Cells(1).Value = Left(teamsVal, pos - 1)
        End If
    Next i

    'AutoFit Columns for readability
    facilitySheet.Columns.AutoFit
    
    'Give the table a friendly name
    Dim usageTableName As String
    usageTableName = GetValidFacilityTableName(facilityName, dayOfWeek)
    
    If Len(usageTableName) > 0 Then
        dataUsageTable.name = usageTableName
    End If
    
    Exit Sub
    
Err:
    DisplayError "FormatFacilityRange", Err
End Sub

Sub PersistFacilitiesMap()
    On Error GoTo Err

    Dim mapSheet As Worksheet
    
    If Not SheetExists(FACILITIES_MAP_SHEET_NAME) Then
        'Create the hidden worksheet
        Set mapSheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        
        mapSheet.name = FACILITIES_MAP_SHEET_NAME
        
        mapSheet.Visible = xlSheetVeryHidden
    Else
        Set mapSheet = ActiveWorkbook.Sheets(FACILITIES_MAP_SHEET_NAME)
        
        'Clear any old map data
        mapSheet.UsedRange.ClearContents
    End If
    
    'Write the map
    Dim i As Integer
    
    'First a header
    mapSheet.Range("A1").Offset(0, 0).Value = "USTA Number"
    mapSheet.Range("A1").Offset(0, 1).Value = "Sheet Name"
    mapSheet.Range("A1").Offset(0, 2).Value = "Facility Name"
    
    For i = 1 To facilitiesSheetMap.Count
        mapSheet.Range("A1").Offset(i, 0).Value = facilitiesNumberMap.Item(i)
        mapSheet.Range("A1").Offset(i, 1).Value = facilitiesSheetMap.Item(i)
        mapSheet.Range("A1").Offset(i, 2).Value = facilitiesNameMap.Item(i)
    Next i
    
    'Map is written, clear the map objects
    Set facilitiesSheetMap = Nothing
    Set facilitiesNumberMap = Nothing
    Set facilitiesNameMap = Nothing
        
    Exit Sub
    
Err:
    'For debugging only, if map doesn't persist then there will be no navigation
    'DisplayError "PersistFacilitiesMap", Err
End Sub

'Get a valid Excel table name from a facility name and day of the week
'Since we are using a facility name (which should be unique) we won't bother checking for duplicates
'We are also appending with _<day of week> so there should be no chance of generating a name
'that's a valid cell reference so we won't check that either
Function GetValidFacilityTableName(facilityName As String, dayOfWeek As String) As String
    On Error GoTo Err
    
    GetValidFacilityTableName = vbNullString
    
    Dim checkChar As String
    Dim proposedName As String
    
    'First character must be a letter or underscore
    checkChar = Left(facilityName, 1)
    If IsAlpha(checkChar) Or checkChar = "_" Then
        proposedName = vbNullString
    Else
        proposedName = "_"  'Start the range name with an underscore
    End If
    
    Const REPLACE_CHAR As String = "."
    
    Dim i As Integer
    
    For i = 1 To Len(facilityName)
        checkChar = Mid(facilityName, i, 1)
        
        'Only letters, numbers, period or underscore are allowed
        If IsAlpha(checkChar) Or IsNumeric(checkChar) Or checkChar = "." Or checkChar = "_" Then
            proposedName = proposedName & checkChar
        Else
            'Handle a few special cases
            If checkChar = " " Then
                proposedName = proposedName & REPLACE_CHAR
            ElseIf checkChar = "&" Then
                proposedName = proposedName & "and"
            End If
        End If
    Next i
    
    proposedName = proposedName & "_" & dayOfWeek
    
    'Just in case
    If Len(proposedName) > 255 Then
        proposedName = Left(proposedName, 255)
    End If
    
    GetValidFacilityTableName = proposedName
    
    Exit Function

Err:
    'Error message for debugging only, if there's a failure here the table will just use its default name
    'DisplayError "GetValidFacilityTableName", Err
    
    GetValidFacilityTableName = vbNullString
End Function

Function IsAlpha(checkChar As String) As Boolean
    On Error GoTo Err
    
    IsAlpha = False
    
    Dim i As Integer
    
    For i = 1 To Len(checkChar)
        Select Case Asc(Mid(checkChar, i, 1))
            Case 65 To 90, 97 To 122  'Upper and lower-case ASCII letters
                IsAlpha = True
            Case Else
                IsAlpha = False
                
                Exit For
        End Select
    Next
    
    Exit Function
    
Err:
    'Unexpected issue so err on the side of caution
    IsAlpha = False
End Function

'Get a valid Excel worksheet name from a facility name
'Remove invalid characters, trim to length, and check for duplicates
Function GetValidFacilitySheetName(facilityName As String) As String
    On Error GoTo Err
    
    GetValidFacilitySheetName = vbNullString
    
    Dim proposedName As String
    proposedName = facilityName
    
    Dim invalidChars As Variant
    Dim invalidChar As Variant
    
    'Not allowed in worksheet names
    invalidChars = Array(":", "/", "\", "?", "*", "[", "]")
     
    For Each invalidChar In invalidChars
        Select Case invalidChar
            Case ":"
                proposedName = Replace(proposedName, invalidChar, vbNullString)
            Case "/"
                proposedName = Replace(proposedName, invalidChar, "-")
            Case "\"
                proposedName = Replace(proposedName, invalidChar, "-")
            Case "?"
                proposedName = Replace(proposedName, invalidChar, vbNullString)
            Case "*"
                proposedName = Replace(proposedName, invalidChar, vbNullString)
            Case "["
                proposedName = Replace(proposedName, invalidChar, "(")
            Case "]"
                proposedName = Replace(proposedName, invalidChar, ")")
        End Select
    Next invalidChar
     
    'Maximum worksheet name length is 31 characters
    proposedName = Left(proposedName, 31)
    
    Dim i As Integer
    
    For i = 1 To 9
        'We have a valid name, now check for duplicates (only do nine for simplicity)
        If Not SheetExists(proposedName) Then
            GetValidFacilitySheetName = proposedName  'Have a unique name
            Exit Function
        Else
            'Try again, adding a number to the end
            proposedName = Left(proposedName, Len(proposedName) - 1) & i
        End If
    Next i

    Exit Function

Err:
    DisplayError "GetValidFacilitySheetName", Err
    
    GetValidFacilitySheetName = vbNullString
End Function

Function SheetExists(sheetName As String) As Boolean
    On Error GoTo DoesNotExist
    
    'If sheet exists then we can get the name
    Dim name As String
    name = Worksheets(sheetName).name
    
    SheetExists = True
    
    Exit Function

DoesNotExist:
    'Exception referencing the sheet by name, it doesn't exist
    SheetExists = False
End Function

Sub SheetAlphaOrder(newSheet As Worksheet)
    On Error GoTo Err
    
    Dim newSheetName As String
    newSheetName = newSheet.name
    
    Dim i As Integer
    
    'Loop through the existing worksheets and then move the specified sheet to its alphabetically correct position
    'Start after any (unsorted) worksheets that existed before adding facilities usage data
    For i = numExistingSheets + 1 To ActiveWorkbook.Sheets.Count
        If ActiveWorkbook.Sheets(i).name >= newSheetName Then
            Exit For
        End If
    Next i
    
    newSheet.Move Before:=ActiveWorkbook.Sheets(i)
    
    Exit Sub
Err:
    'Only show error for debugging
    'DisplayError "SheetAlphaOrder", Err
End Sub


Sub EmailFacilities()
    On Error GoTo Err
    
    If Not SheetExists(FACILITIES_MAP_SHEET_NAME) Then
        MsgBox "The active workbook does not appear to contain any facility usage reports.", vbInformation, ADDIN_NAME
        
        Exit Sub
    End If
    
    Dim mapRange As Range
    Set mapRange = ActiveWorkbook.Sheets(FACILITIES_MAP_SHEET_NAME).UsedRange
    
    Dim facilitySheetName As String
    
    Dim i As Integer
    
    'Clear tab colors
    For i = 2 To mapRange.Rows.Count  'Map has a header so start at row 2
        facilitySheetName = mapRange.Rows(i).Cells(2).Value
        ColorEmailTab facilitySheetName, False
    Next i
    
    Dim currentSheet As Worksheet
    Set currentSheet = ActiveWorkbook.ActiveSheet
    
    For i = 2 To mapRange.Rows.Count
        facilitySheetName = mapRange.Rows(i).Cells(2).Value
        
        ActiveWorkbook.Sheets(facilitySheetName).Activate
            
        If Not EmailFacility(facilitySheetName, True) Then
            Exit For  'User cancelled bulk operation
        End If
    Next i
    
    'Return the user to where he was before emailing all the facilities
    currentSheet.Activate
    
    Exit Sub
    
Err:
    DisplayError "EmailFacilities", Err
End Sub

Function EmailFacility(facilitySheetName As String, ask As Boolean) As Boolean
    On Error Resume Next
    
    ColorEmailTab facilitySheetName, False
    
    'Outlook must be running to send email
    Dim outlookApplication As Object
    Set outlookApplication = GetObject(Class:="Outlook.Application")
    
    On Error GoTo Err

    If outlookApplication Is Nothing Then
        'Try to open it
        Shell "outlook.exe", vbNormalNoFocus
    End If

    EmailFacility = True
            
    Dim facilitySheet As Worksheet
    Dim ustaNumber As String
    Dim facilityEmail As String
    Dim facilityName As String
    
    'This routine may be called from the loop reading the facilities map or using
    'the active worksheet name. Verify the name is in the map before proceeding
    ustaNumber = GetFacilityUSTANumber(facilitySheetName)
    
    If ustaNumber = vbNullString Then
        'Worksheet is not found in the map, we are done
        MsgBox "The current worksheet does not appear to be a valid facility usage report.", vbInformation, ADDIN_NAME
        
        Exit Function
    End If

    'Worksheet name is a valid facility, get the worksheet object
    If SheetExists(facilitySheetName) Then
        Set facilitySheet = ActiveWorkbook.Sheets(facilitySheetName)
    Else
        Exit Function
    End If
    
    'TODO: Error handling
    facilityName = facilitySheet.Range("B1").Value
            
    'Lookup email since it is not always B5
    facilityEmail = Application.WorksheetFunction.VLookup(EMAIL_HEADER, facilitySheet.Range("A1").CurrentRegion, 2, False)
    
    'Get email from facilities information workbook
    'facilityEmail = GetFacilityInfo(ustaNumber, FACILITY_EMAIL_COLUMN)

    'Simple validation check
    If InStr(facilityEmail, "@") < 2 Then
        facilityEmail = vbNullString
    End If
            
    'TEMPORARY - Ask before each email
    Dim res As VbMsgBoxResult
    
    If ask Then
        res = MsgBox("Do you want to send an email to " & facilityName & "?", vbYesNoCancel, Title:=ADDIN_NAME)
    Else
        res = vbYes
    End If
            
    If res = vbYes Then
        Dim mailingRange As Range
        Set mailingRange = facilitySheet.UsedRange  'To send everything
        
        EmailFacility = SendEmail("USTA Court Usage Report " & facilityName, facilityEmail, mailingRange)
    ElseIf res = vbCancel Then
        'If we are mailing all facilities this breaks out of the loop
        EmailFacility = False
        
        Exit Function
    End If
    
    If res = vbYes Then
        'Only color the worksheet tab if the user attempted to send mail (and it was successful)
        ColorEmailTab facilitySheetName, EmailFacility
    End If
            
    Exit Function
    
Err:
    DisplayError "EmailFacility", Err
End Function

Function SendEmail(mailSubject As String, toAddress As String, mailBody As Range) As Boolean
    On Error GoTo Err
    
    SendEmail = True
    
    Dim outlookApp As Object
    Dim outlookMailItem As Object

    'Late binding to help with possibly different Outlook versions
    Set outlookApp = CreateObject("Outlook.Application")
    Set outlookMailItem = outlookApp.CreateItem(0)

    Dim outlookInspector As Object
    Dim wordDocument As Object
    Dim outlookRange As Object
    
    With outlookMailItem
        .To = toAddress
        .subject = mailSubject
        .SentOnBehalfOfName = USTA_EMAIL_FROM

        Dim i As Integer
        
        'See if there is a San Diego USTA account to send from in Outlook
        For i = 1 To outlookApp.Session.Accounts.Count
            If LCase(outlookApp.Session.Accounts.Item(i).SmtpAddress) = LCase(USTA_EMAIL_ADDRESS) Then
                Set .SendUsingAccount = outlookApp.Session.Accounts.Item(i)
                
                Exit For
            End If
        Next i
        
        'See https://answers.microsoft.com/en-us/office/forum/office_2013_release-outlook/paste-clipboard-to-outlook-with-vba/a1a27b25-1534-42ac-ae82-281e66ba30b6
        Set outlookInspector = .GetInspector
        Set wordDocument = outlookInspector.WordEditor
        
        If wordDocument Is Nothing Then
            MsgBox "You must enable Word as your email editor in Outlook to send mail.", vbInformation, ADDIN_NAME
            
            SendEmail = False  'Cancel bulk operation
            
            Exit Function
        Else
            Set outlookRange = wordDocument.Range
            outlookRange.collapse 1  'Collapse Word to the starting point
            
            'Insert the header into the email body
            outlookRange.InsertBefore GetEmailHeader
            
            outlookRange.collapse 0 'Collapse Word to the end point
        
            'Copy/paste the message body from Excel to Outlook
            mailBody.Copy
            outlookRange.Paste
        
            Application.CutCopyMode = False
            
            .BodyFormat = 2  '0: olFormatUnspecified, 1: olFormatPlain, 2: olFormatHTML, 3: olFormatRichText
            
            .Display  'Display the mail
            '.Send  'Send the mail
        End If
    End With

    Set outlookMailItem = Nothing
    Set outlookApp = Nothing
    Set outlookInspector = Nothing
    Set wordDocument = Nothing
    Set outlookRange = Nothing
    
    Exit Function
    
Err:
    DisplayError "EmailFacility", Err
End Function

Function GetEmailHeader() As String
    GetEmailHeader = _
        "Hello Club Tennis Directors and League Coordinators!" & _
        vbNewLine & vbNewLine & _
        "I want to share a new reporting tool with you in hopes it makes your court blocking process go more smoothly. The report is organized by day of the week and shows when you will need to reserve 'home' " & _
        "courts. Feedback from a few clubs I tested in Spring felt that it was very useful. I will run this information for you at the start of each season (next time you will have it sooner but the first time " & _
        "takes the longest). This is a snapshot of the schedule as of June 18, 2017. Any changes after this date to match dates or scheduling of 'floating matches' by your captains will need to be shared with " & _
        "you directly. This report has the few remaining Spring Season matches as well as the current Summer season." & _
        vbNewLine & vbNewLine & _
        "Key to Terms:" & vbNewLine & _
        "H: Home (please book courts)" & vbNewLine & _
        "A: Away" & vbNewLine & _
        "B: Bye" & vbNewLine & _
        "X: Floating match - ALL matches on 7/4 (4th of July) are floating matches. These are matches that could not fit into the regular timeframe of the league. Captains are expected to contact each other " & _
        "and set up a new date to play. I will update in the USTA system when that happens but expect them to coordinate the courts with the ‘home’ club. DO NOT book courts for 7/4 unless requested by the captain." & _
        vbNewLine & vbNewLine & _
        "Thanks for any feedback you may have to make this new tool better since it was created for you!" & _
        vbNewLine & vbNewLine & _
        "Randie Lettington" & vbNewLine & _
        "USTA San Diego District/Area League Coordinator" & vbNewLine & _
        "SanDiegoUSTA@gmail.com" & vbNewLine & _
        "619.251.0103" & vbNewLine & vbNewLine & vbNewLine
End Function

Sub ColorEmailTab(facilitySheetName As String, sent As Boolean)
    On Error Resume Next
    
    With ActiveWorkbook.Sheets(facilitySheetName).Tab
        If sent Then
            'Color the worksheet tab (green) to indicate email sent
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.599993896298105
        Else
            'Cear the worksheet tab color
            .ColorIndex = xlColorIndexNone
        End If
    End With
End Sub


Function GetFacilityInfo(ustaNumber As String, infoColumn As String) As String
    On Error Resume Next

    Const ADO_OPEN_STATIC = 3
    Const ADO_LOCK_OPTIMISTIC = 3
    Const ADO_COMMAND_TEXT = &H1
    
    'If not found or error, return empty string
    GetFacilityInfo = vbNullString

    Dim adoConnection As Object
    Dim adoRecordset As Object

    Set adoConnection = CreateObject("ADODB.Connection")
    Set adoRecordset = CreateObject("ADODB.Recordset")

    'Use Microsoft.jet.oledb.4.0 for 32-bit systems?
    adoConnection.Open "Provider=microsoft.ace.oledb.12.0;Data Source=" & FACILITY_INFO_PATH & ";Extended Properties=""Excel 8.0;HDR=Yes;"";"

    'Sheet name must be contained within square brackets and have a trailing $
    adoRecordset.Open "select [" & infoColumn & "] from [" & FACILITY_INFO_SHEET_NAME & "$] where [" & USTA_NUMBER_COLUMN & "]=" & ustaNumber, _
        adoConnection, ADO_OPEN_STATIC, ADO_LOCK_OPTIMISTIC, ADO_COMMAND_TEXT
    
    'Only one value returned
    If adoRecordset.Fields.Count > 0 Then
        GetFacilityInfo = adoRecordset.Fields.Item(0)
    End If
End Function
