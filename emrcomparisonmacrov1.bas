Option Explicit

' ============================================================
'  ACDC Data Migration - EMR (Oracle Millennium) -> Bookwise
'  Migration Check Macro  v1
'
'  Sub: CheckEMRtoBookwiseMigration
'  One-way check only: EMR extract is validated AGAINST the
'  consolidated Bookwise master. Bookwise bookings missing from
'  the EMR are NOT flagged (migrations are progressive).
'
'  Deploy as an Excel Add-In (.xlam); run from a Quick Access
'  Toolbar button while the EMR extract is the active workbook.
'  ActiveSheet.Parent is used throughout (NOT ThisWorkbook) so
'  output sheets land in the user's workbook, not the add-in.
'
'  Bookwise layout: consolidated multi-site sheet, headers in
'  ROW 1, data from ROW 2 (no search-parameter / separator rows).
'  Must include a 'Location' column populated with the site each
'  block of rows came from: Box Hill / Maroondah / Yarra.
'
'  Site check: for compared rows, the EMR LOCATION value implies
'  an expected Bookwise site, which is verified against the
'  Bookwise 'Location' column (trimmed, case-insensitive). A
'  mismatch is treated like any other field discrepancy.
'
'  Year handling: EMR 2-digit years all map to 2000s (26 -> 2026).
'  STATUS gate: case-insensitive match to "Booked(Confirmed)".
'
'  Phase 2 hook: iPM-location rows are routed to manual review
'  via LocationCategory(). To automate the iPM comparison later,
'  add a second read-only picker and replace the iPM branch only.
' ============================================================

' --- Highlight colours -------------------------------------
Private Const CLR_YELLOW As Long = 65535          ' vbYellow  - field mismatch
Private Const CLR_ORANGE As Long = 39423          ' RGB(255,153,0) - ID / location issues

Public Sub CheckEMRtoBookwiseMigration()

    Dim wbEMR        As Workbook
    Dim wsEMR        As Worksheet
    Dim wbBook       As Workbook
    Dim wsBook       As Worksheet
    Dim wsMis        As Worksheet      ' Bookwise Mismatches to Review
    Dim wsUnk        As Worksheet      ' EMR Appt Unkn ID
    Dim wsAudit      As Worksheet

    Dim bookDict     As Object
    Dim dupTrack     As Object
    Dim idCount      As Object

    Dim bookPath     As String
    Dim lastRowEMR   As Long
    Dim lastColEMR   As Long
    Dim lastRowBook  As Long
    Dim lastColBook  As Long
    Dim bookDataStart As Long

    ' EMR header columns (resolved by name in row 1)
    Dim mrCol        As Long           ' Manual Review Status (always far left, col 1)
    Dim idColE       As Long           ' ACDC_BOOKING_ID
    Dim mrnColE      As Long           ' MRN
    Dim locColE      As Long           ' LOCATION
    Dim dateColE     As Long           ' APPOINTMENT_DATE
    Dim timeColE     As Long           ' APPOINTMENT_TIME
    Dim statColE     As Long           ' STATUS

    ' Bookwise header columns (resolved by name in ROW 1)
    Dim bkBookCol    As Long           ' Book No.
    Dim bkUrCol      As Long           ' UR Number
    Dim bkDateCol    As Long           ' Date
    Dim bkTimeCol    As Long           ' Start Time
    Dim bkLocCol     As Long           ' Location (site)

    Dim srcCols()    As Long           ' EMR source columns for the Unknown-ID sheet (all except mrCol)
    Dim nSrc         As Long

    Dim r            As Long
    Dim i            As Long
    Dim runTime      As Date

    ' Counters
    Dim cTotal       As Long
    Dim cCompared    As Long
    Dim cClean       As Long
    Dim cMismatch    As Long
    Dim cUnknownID   As Long
    Dim cMrTotal     As Long
    Dim cMissingID   As Long
    Dim cDupID       As Long
    Dim cUnknLoc     As Long
    Dim cIPM         As Long
    Dim cStatus      As Long

    Dim dupMsg       As String
    Dim dk           As Variant

    runTime = Now()

    ' --------------------------------------------------------
    ' 0. SHEET NAME CHECK  (absolute first - no flags, no change)
    ' --------------------------------------------------------
    If ActiveSheet.Name <> "EMR Extract" Then
        MsgBox "This macro must be run on the EMR extract." & vbCrLf & vbCrLf & _
               "Please rename the active sheet to exactly  EMR Extract  and try again.", _
               vbCritical, "Wrong Sheet"
        Exit Sub
    End If

    Set wsEMR = ActiveSheet
    Set wbEMR = ActiveSheet.Parent

    ' --------------------------------------------------------
    ' 1. FILE PICKER  (before any change - cancel = untouched)
    ' --------------------------------------------------------
    bookPath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx;*.xlsm;*.xls;*.xlsb),*.xlsx;*.xlsm;*.xls;*.xlsb", _
        Title:="Select the CONSOLIDATED BOOKWISE MASTER sheet to compare against")

    If bookPath = "False" Then
        MsgBox "No Bookwise master selected. Macro cancelled - nothing was changed.", _
               vbExclamation, "Cancelled"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo CleanFail

    ' --------------------------------------------------------
    ' 2. OPEN BOOKWISE (read-only). Active sheet of the file.
    '    Consolidated layout: headers ROW 1, data from ROW 2.
    ' --------------------------------------------------------
    Set wbBook = Workbooks.Open(Filename:=bookPath, ReadOnly:=True)
    Set wsBook = wbBook.ActiveSheet

    bookDataStart = 2

    lastColBook = wsBook.Cells(1, wsBook.Columns.Count).End(xlToLeft).Column
    If lastColBook < 1 Then
        MsgBox "Could not read the Bookwise header row (row 1).", vbCritical, "Bookwise Error"
        GoTo CleanFail
    End If

    bkBookCol = FindHeaderColumn(wsBook, "Book No.", 1, lastColBook)
    bkUrCol = FindHeaderColumn(wsBook, "UR Number", 1, lastColBook)
    bkDateCol = FindHeaderColumn(wsBook, "Date", 1, lastColBook)
    bkTimeCol = FindHeaderColumn(wsBook, "Start Time", 1, lastColBook)

    If bkBookCol = 0 Or bkUrCol = 0 Or bkDateCol = 0 Or bkTimeCol = 0 Then
        MsgBox "The Bookwise master is missing one or more required columns in row 1:" & vbCrLf & _
               "  Book No., UR Number, Date, Start Time." & vbCrLf & vbCrLf & _
               "Verify the selected file and try again.", _
               vbCritical, "Missing Bookwise Column"
        GoTo CleanFail
    End If

    ' Location (site) column - required for the site accuracy check
    bkLocCol = FindHeaderColumn(wsBook, "Location", 1, lastColBook)
    If bkLocCol = 0 Then
        MsgBox "The consolidated Bookwise master is missing a 'Location' column in row 1." & vbCrLf & vbCrLf & _
               "Add a 'Location' column populated with the source site for each block of rows" & vbCrLf & _
               "(Box Hill / Maroondah / Yarra) and try again.", _
               vbCritical, "Missing Bookwise Location Column"
        GoTo CleanFail
    End If

    lastRowBook = wsBook.Cells(wsBook.Rows.Count, bkBookCol).End(xlUp).Row
    If lastRowBook < bookDataStart Then
        MsgBox "No booking data found in the Bookwise master 'Book No.' column.", _
               vbExclamation, "No Bookwise Data"
        GoTo CleanFail
    End If

    ' --------------------------------------------------------
    ' 3. HARD STOP: duplicate Book No. in Bookwise master.
    '    Runs BEFORE any change to the EMR extract.
    ' --------------------------------------------------------
    Set dupTrack = CreateObject("Scripting.Dictionary")
    dupTrack.CompareMode = vbTextCompare

    For i = bookDataStart To lastRowBook
        Dim kb As String
        kb = TrimText(wsBook.Cells(i, bkBookCol).Value)
        If kb <> "" Then
            If dupTrack.Exists(kb) Then
                dupTrack(kb) = dupTrack(kb) & ", " & i
            Else
                dupTrack.Add kb, CStr(i)
            End If
        End If
    Next i

    dupMsg = ""
    For Each dk In dupTrack.Keys
        If InStr(CStr(dupTrack(dk)), ",") > 0 Then
            dupMsg = dupMsg & "  Book No. " & dk & "  ->  Rows: " & dupTrack(dk) & vbCrLf
        End If
    Next dk

    If Len(dupMsg) > 0 Then
        MsgBox "HARD STOP - Duplicate 'Book No.' values found in the Bookwise master:" & vbCrLf & vbCrLf & _
               dupMsg & vbCrLf & _
               "These must be resolved before running. No changes have been made to either workbook.", _
               vbCritical, "Duplicate Book No. in Bookwise"
        GoTo CleanFail
    End If

    ' Build Book No. -> row lookup (no duplicates remain)
    Set bookDict = CreateObject("Scripting.Dictionary")
    bookDict.CompareMode = vbTextCompare
    For Each dk In dupTrack.Keys
        bookDict.Add dk, CLng(dupTrack(dk))
    Next dk
    Set dupTrack = Nothing

    ' --------------------------------------------------------
    ' 3b. HARD STOP: Bookwise 'Location' column must be fully
    '     populated and contain only Box Hill / Maroondah /
    '     Yarra (trimmed, case-insensitive). Checked on rows
    '     that carry a Book No. Runs BEFORE any EMR change.
    ' --------------------------------------------------------
    Dim blankLocMsg As String, badLocMsg As String
    Dim bn As String, lv As String
    blankLocMsg = ""
    badLocMsg = ""

    For i = bookDataStart To lastRowBook
        bn = TrimText(wsBook.Cells(i, bkBookCol).Value)
        If bn <> "" Then
            lv = TrimText(wsBook.Cells(i, bkLocCol).Value)
            If lv = "" Then
                blankLocMsg = blankLocMsg & i & ", "
            ElseIf Not IsValidSite(lv) Then
                badLocMsg = badLocMsg & "  Row " & i & ": """ & lv & """" & vbCrLf
            End If
        End If
    Next i

    If Len(blankLocMsg) > 0 Or Len(badLocMsg) > 0 Then
        Dim locMsg As String
        locMsg = "HARD STOP - Problem with the Bookwise master 'Location' column." & vbCrLf & vbCrLf
        If Len(blankLocMsg) > 0 Then
            blankLocMsg = Left(blankLocMsg, Len(blankLocMsg) - 2)
            locMsg = locMsg & "Blank Location on rows:" & vbCrLf & "  " & blankLocMsg & vbCrLf & vbCrLf
        End If
        If Len(badLocMsg) > 0 Then
            locMsg = locMsg & "Unexpected Location values (must be Box Hill, Maroondah or Yarra):" & vbCrLf & _
                     badLocMsg & vbCrLf
        End If
        locMsg = locMsg & "Resolve these before running. No changes have been made to either workbook."
        MsgBox locMsg, vbCritical, "Bookwise Location Validation"
        GoTo CleanFail
    End If

    ' --------------------------------------------------------
    ' 4. PREP EMR EXTRACT (first point of modification)
    '    Ensure a single 'Manual Review Status' column (idempotent).
    ' --------------------------------------------------------
    mrCol = FindHeaderColumn(wsEMR, "Manual Review Status", 1, _
                wsEMR.Cells(1, wsEMR.Columns.Count).End(xlToLeft).Column)

    If mrCol = 0 Then
        wsEMR.Columns(1).Insert Shift:=xlToRight
        wsEMR.Cells(1, 1).Value = "Manual Review Status"
        mrCol = 1
    End If

    ' Resolve EMR data dimensions and headers AFTER the column exists.
    lastRowEMR = LastUsedRow(wsEMR)
    lastColEMR = LastUsedCol(wsEMR)

    If lastRowEMR < 2 Then
        MsgBox "No data rows found below the header in the EMR extract.", vbExclamation, "No Data"
        GoTo CleanFail
    End If

    idColE = FindHeaderColumn(wsEMR, "ACDC_BOOKING_ID", 1, lastColEMR)
    mrnColE = FindHeaderColumn(wsEMR, "MRN", 1, lastColEMR)
    locColE = FindHeaderColumn(wsEMR, "LOCATION", 1, lastColEMR)
    dateColE = FindHeaderColumn(wsEMR, "APPOINTMENT_DATE", 1, lastColEMR)
    timeColE = FindHeaderColumn(wsEMR, "APPOINTMENT_TIME", 1, lastColEMR)
    statColE = FindHeaderColumn(wsEMR, "STATUS", 1, lastColEMR)

    If idColE = 0 Or mrnColE = 0 Or locColE = 0 Or _
       dateColE = 0 Or timeColE = 0 Or statColE = 0 Then
        MsgBox "The EMR extract is missing one or more required headers in row 1:" & vbCrLf & _
               "  ACDC_BOOKING_ID, MRN, LOCATION, APPOINTMENT_DATE," & vbCrLf & _
               "  APPOINTMENT_TIME, STATUS." & vbCrLf & vbCrLf & _
               "Verify the extract and try again.", _
               vbCritical, "Missing EMR Column"
        GoTo CleanFail
    End If

    ' --------------------------------------------------------
    ' 5. FRESH SLATE: clear prior highlights and Manual Review
    '    values across the data region.
    ' --------------------------------------------------------
    wsEMR.Range(wsEMR.Cells(2, 1), wsEMR.Cells(lastRowEMR, lastColEMR)).Interior.ColorIndex = xlNone
    wsEMR.Range(wsEMR.Cells(2, mrCol), wsEMR.Cells(lastRowEMR, mrCol)).ClearContents

    ' --------------------------------------------------------
    ' 6. (Re)build output sheets in the EMR workbook.
    '    Unknown-ID sheet reproduces EMR data columns EXCEPT the
    '    inserted Manual Review Status column.
    ' --------------------------------------------------------
    ReDim srcCols(1 To lastColEMR)
    nSrc = 0
    For i = 1 To lastColEMR
        If i <> mrCol Then
            nSrc = nSrc + 1
            srcCols(nSrc) = i
        End If
    Next i
    ReDim Preserve srcCols(1 To nSrc)

    Set wsMis = SetupSheetClear(wbEMR, "Bookwise Mismatches to Review")
    For i = 1 To lastColBook
        wsMis.Cells(1, i).Value = SafeText(wsBook.Cells(1, i).Value)
    Next i
    wsMis.Cells(1, lastColBook + 1).Value = "Dt/Tm Added by Macro"
    wsMis.Rows(1).Font.Bold = True

    Set wsUnk = SetupSheetClear(wbEMR, "EMR Appt Unkn ID")
    For i = 1 To nSrc
        wsUnk.Cells(1, i).Value = SafeText(wsEMR.Cells(1, srcCols(i)).Value)
    Next i
    wsUnk.Cells(1, nSrc + 1).Value = "Dt/Tm Added by Macro"
    wsUnk.Rows(1).Font.Bold = True

    ' --------------------------------------------------------
    ' 7. Count EMR booking IDs (non-blank) for duplicate flag.
    '    Duplicates do NOT stop the macro - they are flagged.
    ' --------------------------------------------------------
    Set idCount = CreateObject("Scripting.Dictionary")
    idCount.CompareMode = vbTextCompare
    For r = 2 To lastRowEMR
        Dim idv0 As String
        idv0 = TrimText(wsEMR.Cells(r, idColE).Value)
        If idv0 <> "" Then
            If idCount.Exists(idv0) Then
                idCount(idv0) = idCount(idv0) + 1
            Else
                idCount.Add idv0, 1
            End If
        End If
    Next r

    ' --------------------------------------------------------
    ' 8. MAIN LOOP
    ' --------------------------------------------------------
    For r = 2 To lastRowEMR

        Dim loc As String, idv As String, statusv As String, locCat As String
        Dim flags As String
        Dim fIPM As Boolean, fMiss As Boolean, fDup As Boolean
        Dim fStat As Boolean, fUnkLoc As Boolean, hasFlag As Boolean

        loc = TrimText(wsEMR.Cells(r, locColE).Value)
        idv = TrimText(wsEMR.Cells(r, idColE).Value)
        statusv = TrimText(wsEMR.Cells(r, statColE).Value)

        ' Skip a genuinely empty row (no location, id or status)
        If loc = "" And idv = "" And statusv = "" Then GoTo NextEMRRow
        cTotal = cTotal + 1

        Application.StatusBar = "Checking EMR row " & r & " of " & lastRowEMR & " ..."

        locCat = LocationCategory(loc)

        fIPM = (locCat = "iPM")
        fUnkLoc = (locCat = "Unknown")
        fMiss = (idv = "")
        fDup = (idv <> "" And idCount(idv) > 1)
        fStat = (StrComp(statusv, "Booked(Confirmed)", vbTextCompare) <> 0)

        ' Build combined label (consistent order)
        flags = ""
        If fIPM Then flags = AppendFlag(flags, "iPM Location - Manual Rv")
        If fMiss Then flags = AppendFlag(flags, "Missing ID - Manual Rv")
        If fDup Then flags = AppendFlag(flags, "Duplicate ID - Manual Rv")
        If fStat Then flags = AppendFlag(flags, "Status - Manual Rv")
        If fUnkLoc Then flags = AppendFlag(flags, "Unkn Location - Manual Rv")

        ' Apply manual-review cell highlights (EMR extract)
        If fMiss Or fDup Then wsEMR.Cells(r, idColE).Interior.Color = CLR_ORANGE
        If fUnkLoc Then wsEMR.Cells(r, locColE).Interior.Color = CLR_ORANGE
        ' (iPM and Status: column flag only, no highlight)

        hasFlag = (flags <> "")
        If hasFlag Then
            wsEMR.Cells(r, mrCol).Value = flags
            cMrTotal = cMrTotal + 1
            If fMiss Then cMissingID = cMissingID + 1
            If fDup Then cDupID = cDupID + 1
            If fUnkLoc Then cUnknLoc = cUnknLoc + 1
            If fIPM Then cIPM = cIPM + 1
            If fStat Then cStatus = cStatus + 1
            GoTo NextEMRRow      ' flagged rows are NOT compared
        End If

        ' ----- COMPARED ROW: Bookwise location, confirmed, single ID -----
        cCompared = cCompared + 1

        If Not bookDict.Exists(idv) Then
            ' Unknown ID - not in Bookwise master
            wsEMR.Cells(r, idColE).Interior.Color = CLR_ORANGE
            Call CopyEMRRowToUnknown(wsEMR, wsUnk, r, srcCols, idColE, runTime)
            cUnknownID = cUnknownID + 1
        Else
            Dim bkRow As Long
            Dim diffMRN As Boolean, diffDate As Boolean, diffTime As Boolean, diffSite As Boolean
            Dim expSite As String, bkSite As String
            bkRow = CLng(bookDict(idv))

            diffMRN = (TrimText(wsEMR.Cells(r, mrnColE).Value) <> _
                       TrimText(wsBook.Cells(bkRow, bkUrCol).Value))
            diffDate = Not DatesEqual(wsEMR.Cells(r, dateColE).Value, _
                                      wsBook.Cells(bkRow, bkDateCol).Value)
            diffTime = Not TimesEqual(wsEMR.Cells(r, timeColE).Value, _
                                      wsBook.Cells(bkRow, bkTimeCol).Value)

            ' Site check: EMR LOCATION implies an expected site,
            ' verified against the Bookwise Location column.
            expSite = ExpectedSite(loc)
            bkSite = TrimText(wsBook.Cells(bkRow, bkLocCol).Value)
            diffSite = (StrComp(expSite, bkSite, vbTextCompare) <> 0)

            If diffMRN Then wsEMR.Cells(r, mrnColE).Interior.Color = CLR_YELLOW
            If diffDate Then wsEMR.Cells(r, dateColE).Interior.Color = CLR_YELLOW
            If diffTime Then wsEMR.Cells(r, timeColE).Interior.Color = CLR_YELLOW
            If diffSite Then wsEMR.Cells(r, locColE).Interior.Color = CLR_YELLOW

            If diffMRN Or diffDate Or diffTime Or diffSite Then
                Call CopyBookwiseRowToMismatch(wsBook, wsMis, bkRow, lastColBook, _
                        bkUrCol, bkDateCol, bkTimeCol, bkLocCol, _
                        diffMRN, diffDate, diffTime, diffSite, runTime)
                cMismatch = cMismatch + 1
            Else
                cClean = cClean + 1
            End If
        End If

NextEMRRow:
    Next r

    ' --------------------------------------------------------
    ' 9. AUDIT LOG (visible, protected)
    ' --------------------------------------------------------
    Set wsAudit = GetOrCreateAuditSheet(wbEMR)
    wsAudit.Unprotect
    Dim aRow As Long
    aRow = wsAudit.Cells(wsAudit.Rows.Count, 1).End(xlUp).Row + 1
    With wsAudit
        .Cells(aRow, 1).Value = runTime
        .Cells(aRow, 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        .Cells(aRow, 2).Value = Environ("USERNAME")
        .Cells(aRow, 3).Value = bookPath
        .Cells(aRow, 4).Value = cTotal
        .Cells(aRow, 5).Value = cCompared
        .Cells(aRow, 6).Value = cClean
        .Cells(aRow, 7).Value = cMismatch
        .Cells(aRow, 8).Value = cUnknownID
        .Cells(aRow, 9).Value = cMrTotal
        .Cells(aRow, 10).Value = cMissingID
        .Cells(aRow, 11).Value = cDupID
        .Cells(aRow, 12).Value = cUnknLoc
        .Cells(aRow, 13).Value = cIPM
        .Cells(aRow, 14).Value = cStatus
    End With
    wsAudit.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True

    ' --------------------------------------------------------
    ' 10. CLOSE BOOKWISE (never saved) + SUMMARY
    ' --------------------------------------------------------
    wbBook.Close SaveChanges:=False
    Set wbBook = Nothing

    MsgBox "EMR -> Bookwise Migration Check Complete" & vbCrLf & vbCrLf & _
           "Total EMR rows processed:     " & cTotal & vbCrLf & _
           "Bookwise-location rows compared: " & cCompared & vbCrLf & _
           "   Clean matches:             " & cClean & vbCrLf & _
           "   Mismatches (field/site):   " & cMismatch & vbCrLf & _
           "   Unknown IDs (not in Bookwise): " & cUnknownID & vbCrLf & vbCrLf & _
           "Manual review (set aside):    " & cMrTotal & vbCrLf & _
           "   Missing ID:        " & cMissingID & vbCrLf & _
           "   Duplicate ID:      " & cDupID & vbCrLf & _
           "   Unknown Location:  " & cUnknLoc & vbCrLf & _
           "   iPM Location:      " & cIPM & vbCrLf & _
           "   Status:            " & cStatus, _
           vbInformation, "Migration Check Summary"

    GoTo CleanExit

' ============================================================
'  ERROR / EXIT HANDLING
' ============================================================
CleanFail:
    If Not wbBook Is Nothing Then
        On Error Resume Next
        wbBook.Close SaveChanges:=False
        On Error GoTo 0
        Set wbBook = Nothing
    End If
    On Error Resume Next
    If Not wsAudit Is Nothing Then
        wsAudit.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    End If
    On Error GoTo 0
    If Err.Number <> 0 Then
        MsgBox "An unexpected error occurred during the migration check." & vbCrLf & vbCrLf & _
               "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
               "The Bookwise file was not saved or modified.", _
               vbCritical, "Macro Error"
    End If

CleanExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

End Sub


' ============================================================
'  SafeText - "" for Null / Empty / error; CStr otherwise.
' ============================================================
Public Function SafeText(v As Variant) As String
    On Error Resume Next
    If IsNull(v) Or IsEmpty(v) Or IsError(v) Then
        SafeText = ""
    Else
        SafeText = CStr(v)
    End If
    On Error GoTo 0
End Function


' ============================================================
'  TrimText - SafeText + Trim (used before every comparison).
' ============================================================
Public Function TrimText(v As Variant) As String
    TrimText = Trim$(SafeText(v))
End Function


' ============================================================
'  AppendFlag - join manual-review labels with "; ".
' ============================================================
Private Function AppendFlag(existing As String, lbl As String) As String
    If existing = "" Then
        AppendFlag = lbl
    Else
        AppendFlag = existing & "; " & lbl
    End If
End Function


' ============================================================
'  LocationCategory - "iPM" / "Bookwise" / "Unknown".
'  Trimmed, case-insensitive. loc is assumed already trimmed.
'  Phase 2: keep iPM list here; the iPM branch in the main loop
'  is the only place that needs to change to automate it.
' ============================================================
Public Function LocationCategory(loc As String) As String
    Dim ipm As Variant, bw As Variant, x As Variant
    ipm = Array("A 4.2 BHH", "ADH Clinic", "BCC Ground Floor MAR", "OP YRS", "Outpatients HEA")
    bw = Array("4.3 BHH", "Alexandra Day Oncology BHH", "DAYCHEMO MAR", "YR ONC YRS")

    LocationCategory = "Unknown"
    For Each x In ipm
        If StrComp(loc, CStr(x), vbTextCompare) = 0 Then
            LocationCategory = "iPM"
            Exit Function
        End If
    Next x
    For Each x In bw
        If StrComp(loc, CStr(x), vbTextCompare) = 0 Then
            LocationCategory = "Bookwise"
            Exit Function
        End If
    Next x
End Function


' ============================================================
'  ExpectedSite - maps a Bookwise EMR LOCATION value to the
'  site that should appear in the Bookwise 'Location' column.
'  Trimmed, case-insensitive. "" if no mapping (defensive).
' ============================================================
Public Function ExpectedSite(emrLoc As String) As String
    Select Case True
        Case StrComp(emrLoc, "4.3 BHH", vbTextCompare) = 0
            ExpectedSite = "Box Hill"
        Case StrComp(emrLoc, "Alexandra Day Oncology BHH", vbTextCompare) = 0
            ExpectedSite = "Box Hill"
        Case StrComp(emrLoc, "DAYCHEMO MAR", vbTextCompare) = 0
            ExpectedSite = "Maroondah"
        Case StrComp(emrLoc, "YR ONC YRS", vbTextCompare) = 0
            ExpectedSite = "Yarra"
        Case Else
            ExpectedSite = ""
    End Select
End Function


' ============================================================
'  IsValidSite - True if s is Box Hill / Maroondah / Yarra
'  (trimmed by caller, compared case-insensitive).
' ============================================================
Public Function IsValidSite(s As String) As Boolean
    IsValidSite = (StrComp(s, "Box Hill", vbTextCompare) = 0 Or _
                   StrComp(s, "Maroondah", vbTextCompare) = 0 Or _
                   StrComp(s, "Yarra", vbTextCompare) = 0)
End Function


' ============================================================
'  FindHeaderColumn - locate headerName in headerRow up to
'  colLimit (case-insensitive, trimmed). 0 if not found.
' ============================================================
Public Function FindHeaderColumn(ws As Worksheet, headerName As String, _
                                  headerRow As Long, colLimit As Long) As Long
    Dim c As Long
    FindHeaderColumn = 0
    For c = 1 To colLimit
        If StrComp(TrimText(ws.Cells(headerRow, c).Value), headerName, vbTextCompare) = 0 Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
End Function


' ============================================================
'  LastUsedRow / LastUsedCol - robust last-used finders.
' ============================================================
Public Function LastUsedRow(ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If f Is Nothing Then LastUsedRow = 1 Else LastUsedRow = f.Row
End Function

Public Function LastUsedCol(ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                          SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If f Is Nothing Then LastUsedCol = 1 Else LastUsedCol = f.Column
End Function


' ============================================================
'  DatesEqual - parse BOTH sides to a real date (Australian
'  dd/mm; 2-digit years -> 2000s) and compare the DATE value
'  only. Unparseable on either side -> NOT equal (surfaces it
'  for manual review rather than silently passing).
' ============================================================
Public Function DatesEqual(v1 As Variant, v2 As Variant) As Boolean
    Dim d1 As Date, d2 As Date
    If Not TryParseDate(v1, d1) Then DatesEqual = False: Exit Function
    If Not TryParseDate(v2, d2) Then DatesEqual = False: Exit Function
    DatesEqual = (d1 = d2)
End Function

Public Function TryParseDate(v As Variant, ByRef outDate As Date) As Boolean
    Dim s As String, parts() As String
    Dim d As Long, m As Long, y As Long
    TryParseDate = False

    ' True date / numeric serial (Bookwise) - take date portion only
    If VarType(v) = vbDate Then
        outDate = Int(CDate(v)): TryParseDate = True: Exit Function
    End If
    If IsNumeric(v) And VarType(v) <> vbString Then
        outDate = Int(CDbl(v)): TryParseDate = True: Exit Function
    End If

    ' Text path: dd/mm/yy or dd/mm/yyyy (accept / . - separators)
    s = TrimText(v)
    If s = "" Then Exit Function
    s = Replace(Replace(s, ".", "/"), "-", "/")
    parts = Split(s, "/")
    If UBound(parts) <> 2 Then Exit Function
    If Not (IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2))) Then Exit Function

    d = CLng(Val(parts(0)))
    m = CLng(Val(parts(1)))
    y = CLng(Val(parts(2)))
    If y < 100 Then y = 2000 + y            ' all 2-digit years -> 2000s
    If m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function

    On Error Resume Next
    outDate = DateSerial(y, m, d)
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Function
    On Error GoTo 0
    TryParseDate = True
End Function


' ============================================================
'  TimesEqual - normalise both to hh:mm and compare.
'  Unparseable on either side -> NOT equal.
' ============================================================
Public Function TimesEqual(v1 As Variant, v2 As Variant) As Boolean
    Dim t1 As String, t2 As String
    t1 = NormaliseTime(v1)
    t2 = NormaliseTime(v2)
    If t1 = "" Or t2 = "" Then TimesEqual = False Else TimesEqual = (t1 = t2)
End Function

Public Function NormaliseTime(v As Variant) As String
    Dim s As String, parts() As String
    Dim hh As Long, mm As Long
    NormaliseTime = ""

    If VarType(v) = vbDate Then
        NormaliseTime = Format$(CDate(v), "hh:mm"): Exit Function
    End If
    If IsNumeric(v) And VarType(v) <> vbString Then
        NormaliseTime = Format$(CDate(CDbl(v)), "hh:mm"): Exit Function
    End If

    s = TrimText(v)
    If s = "" Then Exit Function
    parts = Split(s, ":")
    If UBound(parts) >= 1 Then
        If IsNumeric(Trim$(parts(0))) And IsNumeric(Trim$(parts(1))) Then
            hh = CLng(Val(Trim$(parts(0))))
            mm = CLng(Val(Trim$(parts(1))))
            If hh >= 0 And hh < 24 And mm >= 0 And mm < 60 Then
                NormaliseTime = Format$(TimeSerial(hh, mm, 0), "hh:mm")
            End If
        End If
    End If
End Function


' ============================================================
'  SetupSheetClear - return named sheet in wb, fully cleared,
'  created at the end if it does not exist.
' ============================================================
Public Function SetupSheetClear(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet, s As Worksheet
    For Each s In wb.Worksheets
        If s.Name = sheetName Then
            s.Cells.Clear
            Set SetupSheetClear = s
            Exit Function
        End If
    Next s
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheetName
    Set SetupSheetClear = ws
End Function


' ============================================================
'  CopyEMRRowToUnknown - copy an EMR row (values only) to the
'  Unknown-ID sheet using srcCols (excludes Manual Review col).
'  ID cell highlighted orange; timestamp appended.
' ============================================================
Public Sub CopyEMRRowToUnknown(wsEMR As Worksheet, wsUnk As Worksheet, _
                                emrRow As Long, srcCols() As Long, _
                                idColE As Long, ts As Date)
    Dim nextRow As Long, j As Long, n As Long, cv As Variant
    nextRow = wsUnk.Cells(wsUnk.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    n = UBound(srcCols)
    For j = 1 To n
        cv = wsEMR.Cells(emrRow, srcCols(j)).Value
        If IsError(cv) Then
            wsUnk.Cells(nextRow, j).Value = ""
        Else
            wsUnk.Cells(nextRow, j).Value = cv
        End If
        If srcCols(j) = idColE Then
            wsUnk.Cells(nextRow, j).Interior.Color = CLR_ORANGE
        End If
    Next j

    wsUnk.Cells(nextRow, n + 1).Value = ts
    wsUnk.Cells(nextRow, n + 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
End Sub


' ============================================================
'  CopyBookwiseRowToMismatch - copy the matching Bookwise row
'  (values only, all Bookwise columns) to the review sheet.
'  Differing UR Number / Date / Start Time / Location highlighted
'  yellow; timestamp appended.
' ============================================================
Public Sub CopyBookwiseRowToMismatch(wsBook As Worksheet, wsMis As Worksheet, _
                                      bkRow As Long, lastColBook As Long, _
                                      bkUrCol As Long, bkDateCol As Long, bkTimeCol As Long, _
                                      bkLocCol As Long, _
                                      diffMRN As Boolean, diffDate As Boolean, _
                                      diffTime As Boolean, diffSite As Boolean, _
                                      ts As Date)
    Dim nextRow As Long, j As Long, cv As Variant
    nextRow = wsMis.Cells(wsMis.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    For j = 1 To lastColBook
        cv = wsBook.Cells(bkRow, j).Value
        If IsError(cv) Then
            wsMis.Cells(nextRow, j).Value = ""
        Else
            wsMis.Cells(nextRow, j).Value = cv
        End If
    Next j

    If diffMRN Then wsMis.Cells(nextRow, bkUrCol).Interior.Color = CLR_YELLOW
    If diffDate Then wsMis.Cells(nextRow, bkDateCol).Interior.Color = CLR_YELLOW
    If diffTime Then wsMis.Cells(nextRow, bkTimeCol).Interior.Color = CLR_YELLOW
    If diffSite Then wsMis.Cells(nextRow, bkLocCol).Interior.Color = CLR_YELLOW

    wsMis.Cells(nextRow, lastColBook + 1).Value = ts
    wsMis.Cells(nextRow, lastColBook + 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
End Sub


' ============================================================
'  GetOrCreateAuditSheet - visible, protected audit log.
'  Columns: Run Date/Time | Run By | Bookwise File | Processed |
'  Compared | Clean | Mismatches | Unknown IDs | Manual Total |
'  Missing ID | Duplicate ID | Unknown Loc | iPM Loc | Status
' ============================================================
Public Function GetOrCreateAuditSheet(wb As Workbook) As Worksheet
    Const AUDIT_NAME As String = "Reconcile Audit Log"
    Dim ws As Worksheet, s As Worksheet
    For Each s In wb.Worksheets
        If s.Name = AUDIT_NAME Then
            Set GetOrCreateAuditSheet = s
            Exit Function
        End If
    Next s

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = AUDIT_NAME
    ws.Visible = xlSheetVisible
    With ws
        .Cells(1, 1).Value = "Run Date/Time"
        .Cells(1, 2).Value = "Run By (Username)"
        .Cells(1, 3).Value = "Bookwise File Path"
        .Cells(1, 4).Value = "Total Processed"
        .Cells(1, 5).Value = "Compared"
        .Cells(1, 6).Value = "Clean Matches"
        .Cells(1, 7).Value = "Mismatches"
        .Cells(1, 8).Value = "Unknown IDs"
        .Cells(1, 9).Value = "Manual Review Total"
        .Cells(1, 10).Value = "Missing ID"
        .Cells(1, 11).Value = "Duplicate ID"
        .Cells(1, 12).Value = "Unknown Location"
        .Cells(1, 13).Value = "iPM Location"
        .Cells(1, 14).Value = "Status"
        .Rows(1).Font.Bold = True
        .Columns(1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        .Columns(1).ColumnWidth = 20
        .Columns(2).ColumnWidth = 22
        .Columns(3).ColumnWidth = 80
    End With
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Set GetOrCreateAuditSheet = ws
End Function