Attribute VB_Name = "Export"
Option Explicit

'Callback for exportButton onAction
Sub ExportMatchesToAccess(control As IRibbonControl)
    USTAExport
End Sub

Sub USTAExport()
    Dim validated As Boolean
    validated = False
    
    If MsgBox("Export match information into Access?", vbYesNo, ADDIN_NAME) = vbYes Then
        'DREWDREW - For now, always validate before exporting to avoid potential errors
        'If MsgBox("Validate before exporting?", vbYesNo, ADDIN_NAME) = vbYes Then
            If Not USTAValidation() Then
                'Validation failed, allow the user the option to still export
                If MsgBox("There were validation errors, do you still want to export?", vbYesNo, ADDIN_NAME) = vbNo Then
                    Exit Sub
                Else
                    ClearAllValidationErrorBorders
                End If
            Else
                validated = True
            End If
        'End If

        If Not validated Then
            'Update the match team IDs using the (possibly modified) team names prior to saving
            'Can skip this step if validated since the validation process updates the team IDs
            If Not UpdateMatchTeamIDs Then
                MsgBox "Could not update team IDs from current team names. Please check that the matches table is not corrupt.", _
                    vbExclamation, ADDIN_NAME
                Exit Sub
            End If
        End If
        
        'Excel file must be saved to disk before writing to Access
        If (Not ActiveWorkbook.Saved) Or (Len(ActiveWorkbook.Path) = 0) Then
            If MsgBox("You must save the current Excel workbook before exporting to Access.  Do you want to save the workbook now?", _
                vbYesNo, ADDIN_NAME) = vbYes Then
                'Save the workbook
                ActiveWorkbook.Save
            Else
                'Not saved and the user doesn't want to save so we're done
                Exit Sub
            End If
        End If
        
        Dim exportFilePath As String
        
        If (Len(CurrentAccessFilePath) > 0) Then
            'We know where the location of the Access file imported into this Excel workbook
            'Default to export from Excel into the same folder with a default name
            exportFilePath = FolderFromPath(CurrentAccessFilePath) & Application.PathSeparator & DEFAULT_ACCESS_EXPORT_FILE
        Else
            'Use the current workbook's path as the default location
            'exportFilePath = "%TEMP%" & Application.PathSeparator & DEFAULT_ACCESS_EXPORT_FILE
            exportFilePath = ActiveWorkbook.Path & Application.PathSeparator & DEFAULT_ACCESS_EXPORT_FILE
        End If
        
        'exportFilePath = GetAccessFilePath(exportFilePath)
        
        If Export(exportFilePath) Then
            MsgBox "Successfully exported match information to '" & exportFilePath & "'.", vbInformation, ADDIN_NAME
        End If
    End If
End Sub

Function Export(AccessFilePath As String) As Boolean
    On Error GoTo Err
    
    If Not CreateAccessDatabase(AccessFilePath) Then
        Export = False
        Exit Function
    End If
    
    Export = WriteTables(AccessFilePath)
    
    Exit Function
    
Err:
    DisplayError "Export", Err
    
    Export = False
End Function

Function CreateAccessDatabase(AccessFilePath As String) As Boolean
    On Error GoTo Err
    
    Dim msAccess As New access.Application

    'Delete the file first since (should we ask?)
    If Len(Dir$(AccessFilePath)) > 0 Then
        Kill AccessFilePath
    End If
    
    msAccess.NewCurrentDatabase AccessFilePath
    
    CreateAccessDatabase = True
        
    Exit Function
    
Err:
    DisplayError "CreateAccessDatabase", Err
    
    CreateAccessDatabase = False
End Function

Function WriteTables(AccessFilePath As String) As Boolean
    On Error GoTo Err
    
    Dim msAccess As New access.Application
    Dim excelFilePath As String

    excelFilePath = Excel.Application.ActiveWorkbook.FullName

    msAccess.OpenCurrentDatabase (AccessFilePath)
    
    'Write the Excel tables to Access - Access requires relative (A1:B2) addresses and not absolute ($A1:$B2)
    'Also, the range address must include the worksheet name or Access will use the active worksheet
    'Include [#All] in the table range of the header row will not be exported
    Call msAccess.DoCmd.TransferSpreadsheet(acImport, acSpreadsheetTypeExcel12, MATCHES_SHEET_NAME, excelFilePath, True, _
        MATCHES_SHEET_NAME & "!" & Range(MATCHES_TABLE_NAME & "[#All]").Address(RowAbsolute:=False, ColumnAbsolute:=False))
        
    Call msAccess.DoCmd.TransferSpreadsheet(acImport, acSpreadsheetTypeExcel12, TEAMS_SHEET_NAME, excelFilePath, True, _
        TEAMS_SHEET_NAME & "!" & Range(TEAMS_TABLE_NAME & "[#All]").Address(RowAbsolute:=False, ColumnAbsolute:=False))
        
    Call msAccess.DoCmd.TransferSpreadsheet(acImport, acSpreadsheetTypeExcel12, HEADER_SHEET_NAME, excelFilePath, True, _
        HEADER_SHEET_NAME & "!" & Range(HEADER_TABLE_NAME & "[#All]").Address(RowAbsolute:=False, ColumnAbsolute:=False))
           
    'msAccess.Visible = True
    
    WriteTables = True
    
    Exit Function
    
Err:
    DisplayError "WriteTables", Err
    
    WriteTables = False
End Function

Function UpdateMatchTeamIDs() As Boolean
    On Error GoTo Err

    Dim teamID As String
    Dim teamName As String
    
    'Get home team IDs from teams lookup table
    Dim i As Integer
    
    For i = 1 To Range(HOME_TEAM_NAME_RANGE_NAME).Rows.Count
        teamName = Range(HOME_TEAM_NAME_RANGE_NAME).Value2(i, 1)
        
        If Len(teamName) > 0 Then
            teamID = GetTeamID(teamName)
        Else
            teamID = BYE_WEEK_TEAM_ID
        End If
        
        Range(HOME_TEAM_ID_RANGE_NAME).Cells(i, 1) = teamID
    Next i
    
    'Get visiting team IDs from teams lookup table
    For i = 1 To Range(VISITING_TEAM_NAME_RANGE_NAME).Rows.Count
        teamName = Range(VISITING_TEAM_NAME_RANGE_NAME).Value2(i, 1)
        
        If Len(teamName) > 0 Then
            teamID = GetTeamID(teamName)
        Else
            teamID = BYE_WEEK_TEAM_ID
        End If
        
        Range(VISITING_TEAM_ID_RANGE_NAME).Cells(i, 1) = teamID
    Next i
        
    UpdateMatchTeamIDs = True
    
    Exit Function
    
Err:
    DisplayError "UpdateMatchTeamIDs", Err
    
    UpdateMatchTeamIDs = False
End Function

Function UpdateFacilityIDs() As Boolean
    On Error GoTo Err

    Dim facilityID As String
    Dim facilityName As String
    
    'Get facility IDs from facilities lookup table
    Dim i As Integer
    
    For i = 1 To Range(FACILITY_NAME_RANGE_NAME).Rows.Count
        facilityName = Range(FACILITY_NAME_RANGE_NAME).Value2(i, 1)
              
        If facilityName = MISSING_FACILITY_NAME Or Len(facilityName) = 0 Then
            facilityID = BYE_FACILITY_ID
        Else
            facilityID = GetFacilityID(facilityName)
        End If
        
        Range(FACILITY_ID_RANGE_NAME).Cells(i, 1) = facilityID
    Next i
        
    UpdateFacilityIDs = True
    
    Exit Function
    
Err:
    DisplayError "UpdateFacilityIDs", Err
    
    UpdateFacilityIDs = False
End Function










Sub Test1()
    'Add Reference to Microsoft ActiveX Data Objects 2.x Library
    Dim strConnectString        As String
    Dim objConnection           As ADODB.Connection
    Dim strDbPath               As String
 
    'Set database name and DB connection string--------
    strDbPath = "%USERPROFILE%\Documents\USTA\ExportTest.accdb"
    '==================================================
    
    'Microsoft.ACE.OLEDB.12.0
    'strConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDbPath & ";"
    strConnectString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strDbPath & ";"
 
    'Connect Database; insert a new table
    Set objConnection = New ADODB.Connection
    With objConnection
        .Open strConnectString
        .Execute "CREATE TABLE MyTable ([EmpName] text(50) WITH Compression, " & _
                 "[Address1] text(150) WITH Compression, " & _
                 "[Address2] text(150) WITH Compression, " & _
                 "[City] text(50) WITH Compression, " & _
                 "[State] text(2) WITH Compression, " & _
                 "[PIN] text(6) WITH Compression, " & _
                 "[SIN] decimal(6))"
    End With
 
    Set objConnection = Nothing
 
End Sub





Public Sub Test2()
  Dim cn As Object
  Dim dbPath As String
  Dim dbWb As String
  Dim dbWs As String
  Dim scn As String
  Dim dsh As String
  Dim ssql As String
  
  Set cn = CreateObject("ADODB.Connection")
  dbPath = "%USERPROFILE%\Documents\USTA\ExportTest.accdb"
  dbWb = Application.ActiveWorkbook.FullName
  dbWs = Application.ActiveSheet.name
  scn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
  dsh = "[" & Application.ActiveSheet.name & "$]"
  cn.Open scn

  ssql = "INSERT INTO fdFolio ([fdName], [fdOne], [fdTwo]) "
  ssql = ssql & "SELECT * FROM [Excel 12.0;HDR=YES;DATABASE=" & dbWb & "]." & dsh

  cn.Execute ssql

End Sub







