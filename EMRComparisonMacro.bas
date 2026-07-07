Option Explicit

' ============================================================
'  ACDC Data Migration - EMR (Oracle Millennium) Comparison
'  Macro  v4
'
'  Sub: EMRComparisonMacro
'  One-way check: EMR extract validated AGAINST the consolidated
'  Bookwise master AND the consolidated iPM report. Bookings in
'  Bookwise/iPM missing from the EMR are NOT flagged.
'
'  Deploy as an Excel Add-In (.xlam); run from a QAT button while
'  the EMR extract is the active workbook. ActiveSheet.Parent
'  used throughout (NOT ThisWorkbook).
'
'  ROUTING (per EMR row):
'    * ACDC_BOOKING_ID populated -> compare vs Bookwise master
'    * IPM_SCHEDULE_ID populated  -> compare vs iPM report
'    * BOTH populated             -> Manual Review, compared vs
'                                    neither
'
'  ----- v4 CHANGES -------------------------------------------
'  * 'Manual Review Status' column renamed to 'Match Status'
'    (re-run safe: reuses/renames a legacy-named column).
'  * Clean matches labelled 'Ok - appt match bookwise' /
'    'Ok - appt match iPM'. Clean iPM row with blank Session
'    Code shows the note only (no Ok).
'  * Non-compared columns shaded light grey (header + all data
'    rows) on the EMR extract and all output sheets. Compared
'    fields and the ID columns are never greyed.
'  ----- v3 (retained) ----------------------------------------
'  * Second file picker + iPM comparison path; iPM hard stops
'    (missing cols, Cancelled Attend Status, duplicate schedule
'    ID); 'iPM Mismatches to Review'; Both-IDs flag.
'  ----- v2 (retained) ----------------------------------------
'  * Date/time copy-out written as canonical TEXT into Text-
'    formatted cells; locale-safe NormaliseTime.
'
'  Year handling: EMR 2-digit years map to 2000s (26 -> 2026).
'  STATUS gate (EMR): case-insensitive == "Booked(Confirmed)".
' ============================================================

' --- Colours -----------------------------------------------
Private Const CLR_YELLOW As Long = 65535          ' field mismatch
Private Const CLR_ORANGE As Long = 39423          ' RGB(255,153,0) - ID / location
Private Const CLR_GREY   As Long = 14277081       ' RGB(217,217,217) - not compared

Public Sub EMRComparisonMacro()

    Dim wbEMR   As Workbook, wsEMR As Worksheet
    Dim wbBook  As Workbook, wsBook As Worksheet
    Dim wbIPM   As Workbook, wsIPM As Worksheet
    Dim wsMis   As Worksheet          ' Bookwise Mismatches to Review
    Dim wsIMis  As Worksheet          ' iPM Mismatches to Review
    Dim wsUnk   As Worksheet          ' EMR Appt Unkn ID (shared)
    Dim wsAudit As Worksheet

    Dim bookDict As Object, ipmDict As Object
    Dim idCountAcdc As Object, idCountIpm As Object

    Dim bookPath As String, ipmPath As String
    Dim lastRowEMR As Long, lastColEMR As Long
    Dim lastRowBook As Long, lastColBook As Long
    Dim lastRowIPM As Long, lastColIPM As Long

    ' EMR columns
    Dim mrCol As Long, acdcColE As Long, ipmColE As Long, mrnColE As Long
    Dim locColE As Long, dateColE As Long, timeColE As Long, statColE As Long
    ' Bookwise columns
    Dim bkBookCol As Long, bkUrCol As Long, bkDateCol As Long, bkTimeCol As Long, bkLocCol As Long
    ' iPM columns
    Dim ipSchedCol As Long, ipPatCol As Long, ipLocCol As Long, ipDateCol As Long
    Dim ipTimeCol As Long, ipSessCol As Long, ipAttCol As Long

    Dim srcCols() As Long, nSrc As Long
    Dim r As Long, i As Long
    Dim runTime As Date

    ' Counters
    Dim cTotal As Long, cCompBook As Long, cCompIpm As Long
    Dim cClean As Long, cMismatch As Long, cUnknownID As Long
    Dim cMrTotal As Long, cMissingID As Long, cDupID As Long
    Dim cBothID As Long, cUnknLoc As Long, cStatus As Long

    Dim dupMsg As String, dk As Variant

    runTime = Now()

    ' 0. SHEET NAME CHECK ------------------------------------
    If ActiveSheet.Name <> "EMR Extract" Then
        MsgBox "This macro must be run on the EMR extract." & vbCrLf & vbCrLf & _
               "Please rename the active sheet to exactly  EMR Extract  and try again.", _
               vbCritical, "Wrong Sheet"
        Exit Sub
    End If
    Set wsEMR = ActiveSheet
    Set wbEMR = ActiveSheet.Parent

    ' 1. FILE PICKERS (both before any change) ---------------
    bookPath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx;*.xlsm;*.xls;*.xlsb),*.xlsx;*.xlsm;*.xls;*.xlsb", _
        Title:="Step 1 of 2: Select the CONSOLIDATED BOOKWISE MASTER sheet")
    If bookPath = "False" Then
        MsgBox "No Bookwise master selected. Macro cancelled - nothing was changed.", vbExclamation, "Cancelled"
        Exit Sub
    End If
    ipmPath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx;*.xlsm;*.xls;*.xlsb),*.xlsx;*.xlsm;*.xls;*.xlsb", _
        Title:="Step 2 of 2: Select the CONSOLIDATED iPM REPORT")
    If ipmPath = "False" Then
        MsgBox "No iPM report selected. Macro cancelled - nothing was changed.", vbExclamation, "Cancelled"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo CleanFail

    ' 2. OPEN BOOKWISE ---------------------------------------
    Set wbBook = Workbooks.Open(Filename:=bookPath, ReadOnly:=True)
    Set wsBook = wbBook.ActiveSheet
    lastColBook = wsBook.Cells(1, wsBook.Columns.Count).End(xlToLeft).Column
    If lastColBook < 1 Then
        MsgBox "Could not read the Bookwise header row (row 1).", vbCritical, "Bookwise Error": GoTo CleanFail
    End If
    bkBookCol = FindHeaderColumn(wsBook, "Book No.", 1, lastColBook)
    bkUrCol = FindHeaderColumn(wsBook, "UR Number", 1, lastColBook)
    bkDateCol = FindHeaderColumn(wsBook, "Date", 1, lastColBook)
    bkTimeCol = FindHeaderColumn(wsBook, "Start Time", 1, lastColBook)
    bkLocCol = FindHeaderColumn(wsBook, "Location", 1, lastColBook)
    If bkBookCol = 0 Or bkUrCol = 0 Or bkDateCol = 0 Or bkTimeCol = 0 Then
        MsgBox "The Bookwise master is missing required columns in row 1:" & vbCrLf & _
               "  Book No., UR Number, Date, Start Time.", vbCritical, "Missing Bookwise Column": GoTo CleanFail
    End If
    If bkLocCol = 0 Then
        MsgBox "The consolidated Bookwise master is missing a 'Location' column in row 1." & vbCrLf & vbCrLf & _
               "Add a 'Location' column populated with the source site (Box Hill / Maroondah / Yarra).", _
               vbCritical, "Missing Bookwise Location Column": GoTo CleanFail
    End If
    lastRowBook = wsBook.Cells(wsBook.Rows.Count, bkBookCol).End(xlUp).Row
    If lastRowBook < 2 Then
        MsgBox "No booking data found in the Bookwise master 'Book No.' column.", vbExclamation, "No Bookwise Data": GoTo CleanFail
    End If

    ' 3. OPEN iPM --------------------------------------------
    Set wbIPM = Workbooks.Open(Filename:=ipmPath, ReadOnly:=True)
    Set wsIPM = wbIPM.ActiveSheet
    lastColIPM = wsIPM.Cells(1, wsIPM.Columns.Count).End(xlToLeft).Column
    If lastColIPM < 1 Then
        MsgBox "Could not read the iPM report header row (row 1).", vbCritical, "iPM Error": GoTo CleanFail
    End If
    ipSchedCol = FindHeaderColumn(wsIPM, "i.PM Schedules ID", 1, lastColIPM)
    ipPatCol = FindHeaderColumn(wsIPM, "Patient Id", 1, lastColIPM)
    ipLocCol = FindHeaderColumn(wsIPM, "Clinic Location", 1, lastColIPM)
    ipDateCol = FindHeaderColumn(wsIPM, "Appointment Date", 1, lastColIPM)
    ipTimeCol = FindHeaderColumn(wsIPM, "Appointment Time", 1, lastColIPM)
    ipSessCol = FindHeaderColumn(wsIPM, "Session Code", 1, lastColIPM)
    ipAttCol = FindHeaderColumn(wsIPM, "Attend Status", 1, lastColIPM)
    If ipSchedCol = 0 Or ipPatCol = 0 Or ipLocCol = 0 Or ipDateCol = 0 Or _
       ipTimeCol = 0 Or ipSessCol = 0 Or ipAttCol = 0 Then
        MsgBox "The iPM report is missing required columns in row 1:" & vbCrLf & _
               "  i.PM Schedules ID, Patient Id, Clinic Location, Appointment Date," & vbCrLf & _
               "  Appointment Time, Session Code, Attend Status.", vbCritical, "Missing iPM Column": GoTo CleanFail
    End If
    lastRowIPM = wsIPM.Cells(wsIPM.Rows.Count, ipSchedCol).End(xlUp).Row
    If lastRowIPM < 2 Then
        MsgBox "No appointment data found in the iPM report 'i.PM Schedules ID' column.", vbExclamation, "No iPM Data": GoTo CleanFail
    End If

    ' 4. HARD STOPS on source files (BEFORE any EMR change) --
    ' 4a. Duplicate Book No.
    Set bookDict = BuildKeyRowDict(wsBook, bkBookCol, 2, lastRowBook, dupMsg)
    If Len(dupMsg) > 0 Then
        MsgBox "HARD STOP - Duplicate 'Book No.' values in the Bookwise master:" & vbCrLf & vbCrLf & _
               dupMsg & vbCrLf & "Resolve these before running. No changes made to any workbook.", _
               vbCritical, "Duplicate Book No. in Bookwise": GoTo CleanFail
    End If
    ' 4b. Cancelled iPM appointment
    Dim cancMsg As String: cancMsg = ""
    For i = 2 To lastRowIPM
        If StrComp(TrimText(wsIPM.Cells(i, ipAttCol).Value), "Cancelled", vbTextCompare) = 0 Then cancMsg = cancMsg & i & ", "
    Next i
    If Len(cancMsg) > 0 Then
        cancMsg = Left(cancMsg, Len(cancMsg) - 2)
        MsgBox "HARD STOP - Cancelled appointment(s) found in the iPM report." & vbCrLf & vbCrLf & _
               "Rows with Attend Status = Cancelled:" & vbCrLf & "  " & cancMsg & vbCrLf & vbCrLf & _
               "Remove the cancelled appointment(s) before running." & vbCrLf & _
               "No changes have been made to any workbook.", vbCritical, "Cancelled Appointment in iPM": GoTo CleanFail
    End If
    ' 4c. Duplicate i.PM Schedules ID
    dupMsg = ""
    Set ipmDict = BuildKeyRowDict(wsIPM, ipSchedCol, 2, lastRowIPM, dupMsg)
    If Len(dupMsg) > 0 Then
        MsgBox "HARD STOP - Duplicate 'i.PM Schedules ID' values in the iPM report:" & vbCrLf & vbCrLf & _
               dupMsg & vbCrLf & "Resolve these before running. No changes made to any workbook.", _
               vbCritical, "Duplicate i.PM Schedules ID in iPM": GoTo CleanFail
    End If
    ' 4d. Bookwise Location populated + valid
    Dim blankLocMsg As String, badLocMsg As String, bn As String, lv As String
    blankLocMsg = "": badLocMsg = ""
    For i = 2 To lastRowBook
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
            locMsg = locMsg & "Unexpected Location values (must be Box Hill, Maroondah or Yarra):" & vbCrLf & badLocMsg & vbCrLf
        End If
        locMsg = locMsg & "Resolve these before running. No changes made to any workbook."
        MsgBox locMsg, vbCritical, "Bookwise Location Validation": GoTo CleanFail
    End If

    ' 5. PREP EMR EXTRACT (first modification) ---------------
    '    Ensure a single 'Match Status' column. Reuse/rename a
    '    legacy 'Manual Review Status' column if present.
    Dim preCol As Long
    preCol = wsEMR.Cells(1, wsEMR.Columns.Count).End(xlToLeft).Column
    mrCol = FindHeaderColumn(wsEMR, "Match Status", 1, preCol)
    If mrCol = 0 Then
        mrCol = FindHeaderColumn(wsEMR, "Manual Review Status", 1, preCol)
        If mrCol > 0 Then wsEMR.Cells(1, mrCol).Value = "Match Status"   ' rename in place
    End If
    If mrCol = 0 Then
        wsEMR.Columns(1).Insert Shift:=xlToRight
        wsEMR.Cells(1, 1).Value = "Match Status"
        mrCol = 1
    End If

    lastRowEMR = LastUsedRow(wsEMR)
    lastColEMR = LastUsedCol(wsEMR)
    If lastRowEMR < 2 Then
        MsgBox "No data rows found below the header in the EMR extract.", vbExclamation, "No Data": GoTo CleanFail
    End If

    acdcColE = FindHeaderColumn(wsEMR, "ACDC_BOOKING_ID", 1, lastColEMR)
    ipmColE = FindHeaderColumn(wsEMR, "IPM_SCHEDULE_ID", 1, lastColEMR)
    mrnColE = FindHeaderColumn(wsEMR, "MRN", 1, lastColEMR)
    locColE = FindHeaderColumn(wsEMR, "LOCATION", 1, lastColEMR)
    dateColE = FindHeaderColumn(wsEMR, "APPOINTMENT_DATE", 1, lastColEMR)
    timeColE = FindHeaderColumn(wsEMR, "APPOINTMENT_TIME", 1, lastColEMR)
    statColE = FindHeaderColumn(wsEMR, "STATUS", 1, lastColEMR)
    If acdcColE = 0 Or ipmColE = 0 Or mrnColE = 0 Or locColE = 0 Or _
       dateColE = 0 Or timeColE = 0 Or statColE = 0 Then
        MsgBox "The EMR extract is missing required headers in row 1:" & vbCrLf & _
               "  ACDC_BOOKING_ID, IPM_SCHEDULE_ID, MRN, LOCATION," & vbCrLf & _
               "  APPOINTMENT_DATE, APPOINTMENT_TIME, STATUS.", vbCritical, "Missing EMR Column": GoTo CleanFail
    End If

    ' 6. FRESH SLATE (clear colour incl. header row) + grey --
    wsEMR.Range(wsEMR.Cells(1, 1), wsEMR.Cells(lastRowEMR, lastColEMR)).Interior.ColorIndex = xlNone
    wsEMR.Range(wsEMR.Cells(2, mrCol), wsEMR.Cells(lastRowEMR, mrCol)).ClearContents
    ' Grey non-compared EMR columns (never grey: Match Status, IDs, compared fields, STATUS)
    GreyColumns wsEMR, 1, lastRowEMR, 1, lastColEMR, _
        Array(mrCol, acdcColE, ipmColE, mrnColE, locColE, dateColE, timeColE, statColE)

    ' 7. (Re)build output sheets -----------------------------
    ReDim srcCols(1 To lastColEMR): nSrc = 0
    For i = 1 To lastColEMR
        If i <> mrCol Then nSrc = nSrc + 1: srcCols(nSrc) = i
    Next i
    ReDim Preserve srcCols(1 To nSrc)

    Set wsMis = SetupSheetClear(wbEMR, "Bookwise Mismatches to Review")
    For i = 1 To lastColBook: wsMis.Cells(1, i).Value = SafeText(wsBook.Cells(1, i).Value): Next i
    wsMis.Cells(1, lastColBook + 1).Value = "Dt/Tm Added by Macro"
    wsMis.Rows(1).Font.Bold = True

    Set wsIMis = SetupSheetClear(wbEMR, "iPM Mismatches to Review")
    For i = 1 To lastColIPM: wsIMis.Cells(1, i).Value = SafeText(wsIPM.Cells(1, i).Value): Next i
    wsIMis.Cells(1, lastColIPM + 1).Value = "Dt/Tm Added by Macro"
    wsIMis.Rows(1).Font.Bold = True

    Set wsUnk = SetupSheetClear(wbEMR, "EMR Appt Unkn ID")
    For i = 1 To nSrc: wsUnk.Cells(1, i).Value = SafeText(wsEMR.Cells(1, srcCols(i)).Value): Next i
    wsUnk.Cells(1, nSrc + 1).Value = "Dt/Tm Added by Macro"
    wsUnk.Rows(1).Font.Bold = True

    ' 8. Count EMR IDs for duplicate flags -------------------
    Set idCountAcdc = CountKeys(wsEMR, acdcColE, 2, lastRowEMR)
    Set idCountIpm = CountKeys(wsEMR, ipmColE, 2, lastRowEMR)

    ' 9. MAIN LOOP -------------------------------------------
    For r = 2 To lastRowEMR
        Dim loc As String, acdc As String, ipmid As String, statusv As String, locCat As String
        Dim flags As String
        Dim fBoth As Boolean, fMissID As Boolean, fDupID As Boolean, fStat As Boolean, fUnkLoc As Boolean

        loc = TrimText(wsEMR.Cells(r, locColE).Value)
        acdc = TrimText(wsEMR.Cells(r, acdcColE).Value)
        ipmid = TrimText(wsEMR.Cells(r, ipmColE).Value)
        statusv = TrimText(wsEMR.Cells(r, statColE).Value)

        If loc = "" And acdc = "" And ipmid = "" And statusv = "" Then GoTo NextEMRRow
        cTotal = cTotal + 1
        Application.StatusBar = "Checking EMR row " & r & " of " & lastRowEMR & " ..."

        locCat = LocationCategory(loc)
        fUnkLoc = (locCat = "Unknown")
        fStat = (StrComp(statusv, "Booked(Confirmed)", vbTextCompare) <> 0)
        fBoth = (acdc <> "" And ipmid <> "")

        fMissID = False: fDupID = False
        If Not fBoth Then
            If locCat = "Bookwise" Then
                fMissID = (acdc = "")
                fDupID = (acdc <> "" And idCountAcdc(acdc) > 1)
            ElseIf locCat = "iPM" Then
                fMissID = (ipmid = "")
                fDupID = (ipmid <> "" And idCountIpm(ipmid) > 1)
            End If
        End If

        flags = ""
        If fBoth Then flags = AppendFlag(flags, "Both IDs Present - Manual Rv")
        If fMissID Then flags = AppendFlag(flags, "Missing ID - Manual Rv")
        If fDupID Then flags = AppendFlag(flags, "Duplicate ID - Manual Rv")
        If fStat Then flags = AppendFlag(flags, "Status - Manual Rv")
        If fUnkLoc Then flags = AppendFlag(flags, "Unkn Location - Manual Rv")

        If fBoth Then
            wsEMR.Cells(r, acdcColE).Interior.Color = CLR_ORANGE
            wsEMR.Cells(r, ipmColE).Interior.Color = CLR_ORANGE
        ElseIf fMissID Or fDupID Then
            If locCat = "Bookwise" Then wsEMR.Cells(r, acdcColE).Interior.Color = CLR_ORANGE
            If locCat = "iPM" Then wsEMR.Cells(r, ipmColE).Interior.Color = CLR_ORANGE
        End If
        If fUnkLoc Then wsEMR.Cells(r, locColE).Interior.Color = CLR_ORANGE

        If flags <> "" Then
            wsEMR.Cells(r, mrCol).Value = flags
            cMrTotal = cMrTotal + 1
            If fBoth Then cBothID = cBothID + 1
            If fMissID Then cMissingID = cMissingID + 1
            If fDupID Then cDupID = cDupID + 1
            If fStat Then cStatus = cStatus + 1
            If fUnkLoc Then cUnknLoc = cUnknLoc + 1
            GoTo NextEMRRow
        End If

        ' ----- COMPARED ROW -----
        If locCat = "Bookwise" Then
            cCompBook = cCompBook + 1
            CompareBookwiseRow wsEMR, wsBook, wsMis, wsUnk, r, acdc, loc, mrCol, bookDict, srcCols, _
                mrnColE, locColE, dateColE, timeColE, acdcColE, ipmColE, _
                bkUrCol, bkDateCol, bkTimeCol, bkLocCol, lastColBook, runTime, cClean, cMismatch, cUnknownID
        ElseIf locCat = "iPM" Then
            cCompIpm = cCompIpm + 1
            CompareIPMRow wsEMR, wsIPM, wsIMis, wsUnk, r, ipmid, loc, mrCol, ipmDict, srcCols, _
                mrnColE, locColE, dateColE, timeColE, acdcColE, ipmColE, _
                ipPatCol, ipDateCol, ipTimeCol, ipLocCol, ipSessCol, lastColIPM, runTime, cClean, cMismatch, cUnknownID
        End If

NextEMRRow:
    Next r

    ' 9b. Grey non-compared columns on output sheets ---------
    Dim lastRowMis As Long, lastRowIMis As Long, lastRowUnk As Long
    lastRowMis = wsMis.Cells(wsMis.Rows.Count, 1).End(xlUp).Row
    GreyColumns wsMis, 1, lastRowMis, 1, lastColBook, Array(bkBookCol, bkUrCol, bkDateCol, bkTimeCol, bkLocCol)

    lastRowIMis = wsIMis.Cells(wsIMis.Rows.Count, 1).End(xlUp).Row
    GreyColumns wsIMis, 1, lastRowIMis, 1, lastColIPM, Array(ipSchedCol, ipPatCol, ipDateCol, ipTimeCol, ipLocCol, ipSessCol)

    lastRowUnk = wsUnk.Cells(wsUnk.Rows.Count, 1).End(xlUp).Row
    GreyUnknownSheet wsUnk, srcCols, lastRowUnk, _
        Array(mrnColE, locColE, dateColE, timeColE, statColE, acdcColE, ipmColE)

    ' 10. AUDIT LOG ------------------------------------------
    Set wsAudit = GetOrCreateAuditSheet(wbEMR)
    wsAudit.Unprotect
    Dim aRow As Long
    aRow = wsAudit.Cells(wsAudit.Rows.Count, 1).End(xlUp).Row + 1
    With wsAudit
        .Cells(aRow, 1).Value = runTime: .Cells(aRow, 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        .Cells(aRow, 2).Value = Environ("USERNAME")
        .Cells(aRow, 3).Value = bookPath
        .Cells(aRow, 4).Value = ipmPath
        .Cells(aRow, 5).Value = cTotal
        .Cells(aRow, 6).Value = cCompBook
        .Cells(aRow, 7).Value = cCompIpm
        .Cells(aRow, 8).Value = cClean
        .Cells(aRow, 9).Value = cMismatch
        .Cells(aRow, 10).Value = cUnknownID
        .Cells(aRow, 11).Value = cMrTotal
        .Cells(aRow, 12).Value = cMissingID
        .Cells(aRow, 13).Value = cDupID
        .Cells(aRow, 14).Value = cBothID
        .Cells(aRow, 15).Value = cUnknLoc
        .Cells(aRow, 16).Value = cStatus
    End With
    wsAudit.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True

    ' 11. CLOSE SOURCES + SUMMARY ----------------------------
    wbBook.Close SaveChanges:=False: Set wbBook = Nothing
    wbIPM.Close SaveChanges:=False: Set wbIPM = Nothing

    MsgBox "EMR Comparison Complete (v4)" & vbCrLf & vbCrLf & _
           "Total EMR rows processed:  " & cTotal & vbCrLf & vbCrLf & _
           "Bookwise rows compared:    " & cCompBook & vbCrLf & _
           "iPM rows compared:         " & cCompIpm & vbCrLf & _
           "   Clean matches:          " & cClean & vbCrLf & _
           "   Mismatches:             " & cMismatch & vbCrLf & _
           "   Unknown IDs:            " & cUnknownID & vbCrLf & vbCrLf & _
           "Manual review (set aside): " & cMrTotal & vbCrLf & _
           "   Missing ID:      " & cMissingID & vbCrLf & _
           "   Duplicate ID:    " & cDupID & vbCrLf & _
           "   Both IDs Present: " & cBothID & vbCrLf & _
           "   Unknown Location: " & cUnknLoc & vbCrLf & _
           "   Status:          " & cStatus, vbInformation, "Comparison Summary - v4"

    GoTo CleanExit

CleanFail:
    If Not wbBook Is Nothing Then
        On Error Resume Next: wbBook.Close SaveChanges:=False: On Error GoTo 0: Set wbBook = Nothing
    End If
    If Not wbIPM Is Nothing Then
        On Error Resume Next: wbIPM.Close SaveChanges:=False: On Error GoTo 0: Set wbIPM = Nothing
    End If
    On Error Resume Next
    If Not wsAudit Is Nothing Then wsAudit.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    On Error GoTo 0
    If Err.Number <> 0 Then
        MsgBox "An unexpected error occurred during the comparison." & vbCrLf & vbCrLf & _
               "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
               "Source files were not saved or modified.", vbCritical, "Macro Error"
    End If

CleanExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub


' ============================================================
'  CompareBookwiseRow
' ============================================================
Private Sub CompareBookwiseRow(wsEMR As Worksheet, wsBook As Worksheet, wsMis As Worksheet, wsUnk As Worksheet, _
        r As Long, acdc As String, loc As String, mrCol As Long, bookDict As Object, srcCols() As Long, _
        mrnColE As Long, locColE As Long, dateColE As Long, timeColE As Long, acdcColE As Long, ipmColE As Long, _
        bkUrCol As Long, bkDateCol As Long, bkTimeCol As Long, bkLocCol As Long, lastColBook As Long, _
        ts As Date, ByRef cClean As Long, ByRef cMismatch As Long, ByRef cUnknownID As Long)

    Dim bkRow As Long, diffMRN As Boolean, diffDate As Boolean, diffTime As Boolean, diffSite As Boolean

    If Not bookDict.Exists(acdc) Then
        wsEMR.Cells(r, acdcColE).Interior.Color = CLR_ORANGE
        CopyEMRRowToUnknown wsEMR, wsUnk, r, srcCols, acdcColE, dateColE, timeColE, ts
        cUnknownID = cUnknownID + 1
        Exit Sub
    End If

    bkRow = CLng(bookDict(acdc))
    diffMRN = (TrimText(wsEMR.Cells(r, mrnColE).Value) <> TrimText(wsBook.Cells(bkRow, bkUrCol).Value))
    diffDate = Not DatesEqual(wsEMR.Cells(r, dateColE).Value, wsBook.Cells(bkRow, bkDateCol).Value)
    diffTime = Not TimesEqual(wsEMR.Cells(r, timeColE).Value, wsBook.Cells(bkRow, bkTimeCol).Value)
    diffSite = (StrComp(ExpectedSite(loc), TrimText(wsBook.Cells(bkRow, bkLocCol).Value), vbTextCompare) <> 0)

    If diffMRN Then wsEMR.Cells(r, mrnColE).Interior.Color = CLR_YELLOW
    If diffDate Then wsEMR.Cells(r, dateColE).Interior.Color = CLR_YELLOW
    If diffTime Then wsEMR.Cells(r, timeColE).Interior.Color = CLR_YELLOW
    If diffSite Then wsEMR.Cells(r, locColE).Interior.Color = CLR_YELLOW

    If diffMRN Or diffDate Or diffTime Or diffSite Then
        CopyBookwiseRowToMismatch wsBook, wsMis, bkRow, lastColBook, bkUrCol, bkDateCol, bkTimeCol, bkLocCol, _
            diffMRN, diffDate, diffTime, diffSite, ts
        cMismatch = cMismatch + 1
    Else
        wsEMR.Cells(r, mrCol).Value = "Ok - appt match bookwise"
        cClean = cClean + 1
    End If
End Sub


' ============================================================
'  CompareIPMRow - blank Session Code note on every compared
'  iPM row; suppresses the 'Ok' label when present.
' ============================================================
Private Sub CompareIPMRow(wsEMR As Worksheet, wsIPM As Worksheet, wsIMis As Worksheet, wsUnk As Worksheet, _
        r As Long, ipmid As String, loc As String, mrCol As Long, ipmDict As Object, srcCols() As Long, _
        mrnColE As Long, locColE As Long, dateColE As Long, timeColE As Long, acdcColE As Long, ipmColE As Long, _
        ipPatCol As Long, ipDateCol As Long, ipTimeCol As Long, ipLocCol As Long, ipSessCol As Long, lastColIPM As Long, _
        ts As Date, ByRef cClean As Long, ByRef cMismatch As Long, ByRef cUnknownID As Long)

    Dim ipRow As Long, diffMRN As Boolean, diffDate As Boolean, diffTime As Boolean, diffSite As Boolean
    Dim sessBlank As Boolean

    If Not ipmDict.Exists(ipmid) Then
        wsEMR.Cells(r, ipmColE).Interior.Color = CLR_ORANGE
        CopyEMRRowToUnknown wsEMR, wsUnk, r, srcCols, ipmColE, dateColE, timeColE, ts
        cUnknownID = cUnknownID + 1
        Exit Sub
    End If

    ipRow = CLng(ipmDict(ipmid))
    sessBlank = (TrimText(wsIPM.Cells(ipRow, ipSessCol).Value) = "")

    diffMRN = (TrimText(wsEMR.Cells(r, mrnColE).Value) <> TrimText(wsIPM.Cells(ipRow, ipPatCol).Value))
    diffDate = Not DatesEqual(wsEMR.Cells(r, dateColE).Value, wsIPM.Cells(ipRow, ipDateCol).Value)
    diffTime = Not TimesEqual(wsEMR.Cells(r, timeColE).Value, wsIPM.Cells(ipRow, ipTimeCol).Value)
    diffSite = (StrComp(ExpectedIPMLocation(loc), TrimText(wsIPM.Cells(ipRow, ipLocCol).Value), vbTextCompare) <> 0)

    If diffMRN Then wsEMR.Cells(r, mrnColE).Interior.Color = CLR_YELLOW
    If diffDate Then wsEMR.Cells(r, dateColE).Interior.Color = CLR_YELLOW
    If diffTime Then wsEMR.Cells(r, timeColE).Interior.Color = CLR_YELLOW
    If diffSite Then wsEMR.Cells(r, locColE).Interior.Color = CLR_YELLOW

    If sessBlank Then _
        wsEMR.Cells(r, mrCol).Value = AppendFlag(SafeText(wsEMR.Cells(r, mrCol).Value), "iPM Session Code blank - Manual Rv")

    If diffMRN Or diffDate Or diffTime Or diffSite Then
        CopyIPMRowToMismatch wsIPM, wsIMis, ipRow, lastColIPM, ipPatCol, ipDateCol, ipTimeCol, ipLocCol, _
            diffMRN, diffDate, diffTime, diffSite, ts
        cMismatch = cMismatch + 1
    Else
        cClean = cClean + 1
        If Not sessBlank Then wsEMR.Cells(r, mrCol).Value = "Ok - appt match iPM"
    End If
End Sub


' ============================================================
'  Grey helpers
' ============================================================
Private Sub GreyColumns(ws As Worksheet, r1 As Long, r2 As Long, c1 As Long, c2 As Long, nonGrey As Variant)
    Dim c As Long
    If r2 < r1 Then Exit Sub
    For c = c1 To c2
        If Not InLongArray(nonGrey, c) Then ws.Range(ws.Cells(r1, c), ws.Cells(r2, c)).Interior.Color = CLR_GREY
    Next c
End Sub

Private Sub GreyUnknownSheet(wsUnk As Worksheet, srcCols() As Long, lastRow As Long, emrNonGreySrc As Variant)
    Dim j As Long
    If lastRow < 1 Then Exit Sub
    For j = 1 To UBound(srcCols)
        If Not InLongArray(emrNonGreySrc, srcCols(j)) Then _
            wsUnk.Range(wsUnk.Cells(1, j), wsUnk.Cells(lastRow, j)).Interior.Color = CLR_GREY
    Next j
End Sub

Private Function InLongArray(arr As Variant, v As Long) As Boolean
    Dim x As Variant
    For Each x In arr
        If CLng(x) = v Then InLongArray = True: Exit Function
    Next x
End Function


' ============================================================
'  SafeText / TrimText / TwoDig / AppendFlag
' ============================================================
Public Function SafeText(v As Variant) As String
    On Error Resume Next
    If IsNull(v) Or IsEmpty(v) Or IsError(v) Then SafeText = "" Else SafeText = CStr(v)
    On Error GoTo 0
End Function
Public Function TrimText(v As Variant) As String
    TrimText = Trim$(SafeText(v))
End Function
Public Function TwoDig(n As Long) As String
    TwoDig = Format$(n, "00")
End Function
Public Function AppendFlag(existing As String, lbl As String) As String
    Dim e As String: e = Trim$(existing)
    If e = "" Then AppendFlag = lbl Else AppendFlag = e & "; " & lbl
End Function


' ============================================================
'  LocationCategory / ExpectedSite / ExpectedIPMLocation / IsValidSite
' ============================================================
Public Function LocationCategory(loc As String) As String
    Dim ipm As Variant, bw As Variant, x As Variant
    ipm = Array("A 4.2 BHH", "ADH Clinic", "BCC Ground Floor MAR", "OP YRS", "Outpatients HEA")
    bw = Array("4.3 BHH", "Alexandra Day Oncology BHH", "DAYCHEMO MAR", "YR ONC YRS")
    LocationCategory = "Unknown"
    For Each x In ipm
        If StrComp(loc, CStr(x), vbTextCompare) = 0 Then LocationCategory = "iPM": Exit Function
    Next x
    For Each x In bw
        If StrComp(loc, CStr(x), vbTextCompare) = 0 Then LocationCategory = "Bookwise": Exit Function
    Next x
End Function

Public Function ExpectedSite(emrLoc As String) As String
    Select Case True
        Case StrComp(emrLoc, "4.3 BHH", vbTextCompare) = 0: ExpectedSite = "Box Hill"
        Case StrComp(emrLoc, "Alexandra Day Oncology BHH", vbTextCompare) = 0: ExpectedSite = "Box Hill"
        Case StrComp(emrLoc, "DAYCHEMO MAR", vbTextCompare) = 0: ExpectedSite = "Maroondah"
        Case StrComp(emrLoc, "YR ONC YRS", vbTextCompare) = 0: ExpectedSite = "Yarra"
        Case Else: ExpectedSite = ""
    End Select
End Function

Public Function ExpectedIPMLocation(emrLoc As String) As String
    Select Case True
        Case StrComp(emrLoc, "A 4.2 BHH", vbTextCompare) = 0: ExpectedIPMLocation = "BHH Bldg A 4.2 Oncology"
        Case StrComp(emrLoc, "ADH Clinic", vbTextCompare) = 0: ExpectedIPMLocation = "Alexandra District Health Oncology"
        Case StrComp(emrLoc, "BCC Ground Floor MAR", vbTextCompare) = 0: ExpectedIPMLocation = "EH Breast and Cancer Centre"
        Case StrComp(emrLoc, "OP YRS", vbTextCompare) = 0: ExpectedIPMLocation = "YRH OPD Clinic"
        Case StrComp(emrLoc, "Outpatients HEA", vbTextCompare) = 0: ExpectedIPMLocation = "HDH OPD Clinics"
        Case Else: ExpectedIPMLocation = ""
    End Select
End Function

Public Function IsValidSite(s As String) As Boolean
    IsValidSite = (StrComp(s, "Box Hill", vbTextCompare) = 0 Or _
                   StrComp(s, "Maroondah", vbTextCompare) = 0 Or _
                   StrComp(s, "Yarra", vbTextCompare) = 0)
End Function


' ============================================================
'  BuildKeyRowDict / CountKeys / FindHeaderColumn / LastUsed*
' ============================================================
Public Function BuildKeyRowDict(ws As Worksheet, keyCol As Long, startRow As Long, endRow As Long, ByRef dupMsg As String) As Object
    Dim d As Object, track As Object, i As Long, k As String, dk As Variant
    Set track = CreateObject("Scripting.Dictionary"): track.CompareMode = vbTextCompare
    For i = startRow To endRow
        k = TrimText(ws.Cells(i, keyCol).Value)
        If k <> "" Then
            If track.Exists(k) Then track(k) = track(k) & ", " & i Else track.Add k, CStr(i)
        End If
    Next i
    dupMsg = ""
    For Each dk In track.Keys
        If InStr(CStr(track(dk)), ",") > 0 Then dupMsg = dupMsg & "  " & dk & "  ->  Rows: " & track(dk) & vbCrLf
    Next dk
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    For Each dk In track.Keys: d.Add dk, CLng(Split(CStr(track(dk)), ",")(0)): Next dk
    Set BuildKeyRowDict = d
End Function

Public Function CountKeys(ws As Worksheet, keyCol As Long, startRow As Long, endRow As Long) As Object
    Dim d As Object, i As Long, k As String
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    For i = startRow To endRow
        k = TrimText(ws.Cells(i, keyCol).Value)
        If k <> "" Then
            If d.Exists(k) Then d(k) = d(k) + 1 Else d.Add k, 1
        End If
    Next i
    Set CountKeys = d
End Function

Public Function FindHeaderColumn(ws As Worksheet, headerName As String, headerRow As Long, colLimit As Long) As Long
    Dim c As Long
    FindHeaderColumn = 0
    For c = 1 To colLimit
        If StrComp(TrimText(ws.Cells(headerRow, c).Value), headerName, vbTextCompare) = 0 Then FindHeaderColumn = c: Exit Function
    Next c
End Function

Public Function LastUsedRow(ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If f Is Nothing Then LastUsedRow = 1 Else LastUsedRow = f.Row
End Function
Public Function LastUsedCol(ws As Worksheet) As Long
    Dim f As Range
    Set f = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If f Is Nothing Then LastUsedCol = 1 Else LastUsedCol = f.Column
End Function


' ============================================================
'  Date / time comparison + normalisation
' ============================================================
Public Function DatesEqual(v1 As Variant, v2 As Variant) As Boolean
    Dim d1 As Date, d2 As Date
    If Not TryParseDate(v1, d1) Then DatesEqual = False: Exit Function
    If Not TryParseDate(v2, d2) Then DatesEqual = False: Exit Function
    DatesEqual = (d1 = d2)
End Function

Public Function TryParseDate(v As Variant, ByRef outDate As Date) As Boolean
    Dim s As String, parts() As String, d As Long, m As Long, y As Long
    TryParseDate = False
    If VarType(v) = vbDate Then outDate = Int(CDate(v)): TryParseDate = True: Exit Function
    If IsNumeric(v) And VarType(v) <> vbString Then outDate = Int(CDbl(v)): TryParseDate = True: Exit Function
    s = TrimText(v)
    If s = "" Then Exit Function
    s = Replace(Replace(s, ".", "/"), "-", "/")
    parts = Split(s, "/")
    If UBound(parts) <> 2 Then Exit Function
    If Not (IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2))) Then Exit Function
    d = CLng(Val(parts(0))): m = CLng(Val(parts(1))): y = CLng(Val(parts(2)))
    If y < 100 Then y = 2000 + y
    If m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
    On Error Resume Next
    outDate = DateSerial(y, m, d)
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Function
    On Error GoTo 0
    TryParseDate = True
End Function

Public Function NormaliseDateText(v As Variant) As String
    Dim d As Date
    If TryParseDate(v, d) Then
        NormaliseDateText = TwoDig(Day(d)) & "/" & TwoDig(Month(d)) & "/" & Format$(Year(d), "0000")
    Else
        NormaliseDateText = TrimText(v)
    End If
End Function

Public Function NormaliseTime(v As Variant) As String
    Dim s As String, parts() As String, hh As Long, mm As Long, dt As Date
    NormaliseTime = ""
    If VarType(v) = vbDate Then NormaliseTime = TwoDig(Hour(v)) & ":" & TwoDig(Minute(v)): Exit Function
    If IsNumeric(v) And VarType(v) <> vbString Then
        dt = CDate(CDbl(v)): NormaliseTime = TwoDig(Hour(dt)) & ":" & TwoDig(Minute(dt)): Exit Function
    End If
    s = TrimText(v)
    If s = "" Then Exit Function
    parts = Split(s, ":")
    If UBound(parts) >= 1 Then
        If IsNumeric(Trim$(parts(0))) And IsNumeric(Trim$(parts(1))) Then
            hh = CLng(Val(Trim$(parts(0)))): mm = CLng(Val(Trim$(parts(1))))
            If hh >= 0 And hh < 24 And mm >= 0 And mm < 60 Then NormaliseTime = TwoDig(hh) & ":" & TwoDig(mm)
        End If
    End If
End Function

Public Function NormaliseTimeText(v As Variant) As String
    Dim t As String: t = NormaliseTime(v)
    If t = "" Then NormaliseTimeText = TrimText(v) Else NormaliseTimeText = t
End Function

Public Function TimesEqual(v1 As Variant, v2 As Variant) As Boolean
    Dim t1 As String, t2 As String
    t1 = NormaliseTime(v1): t2 = NormaliseTime(v2)
    If t1 = "" Or t2 = "" Then TimesEqual = False Else TimesEqual = (t1 = t2)
End Function


' ============================================================
'  SetupSheetClear
' ============================================================
Public Function SetupSheetClear(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet, s As Worksheet
    For Each s In wb.Worksheets
        If s.Name = sheetName Then s.Cells.Clear: Set SetupSheetClear = s: Exit Function
    Next s
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheetName
    Set SetupSheetClear = ws
End Function


' ============================================================
'  Copy helpers (date/time written as canonical TEXT)
' ============================================================
Public Sub CopyEMRRowToUnknown(wsEMR As Worksheet, wsUnk As Worksheet, emrRow As Long, srcCols() As Long, _
                                idColHi As Long, dateColE As Long, timeColE As Long, ts As Date)
    Dim nextRow As Long, j As Long, n As Long, sc As Long, cv As Variant
    nextRow = wsUnk.Cells(wsUnk.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2
    n = UBound(srcCols)
    For j = 1 To n
        sc = srcCols(j)
        If sc = dateColE Or sc = timeColE Then
            wsUnk.Cells(nextRow, j).NumberFormat = "@"
            wsUnk.Cells(nextRow, j).Value = TrimText(wsEMR.Cells(emrRow, sc).Value)
        Else
            cv = wsEMR.Cells(emrRow, sc).Value
            If IsError(cv) Then wsUnk.Cells(nextRow, j).Value = "" Else wsUnk.Cells(nextRow, j).Value = cv
        End If
        If sc = idColHi Then wsUnk.Cells(nextRow, j).Interior.Color = CLR_ORANGE
    Next j
    wsUnk.Cells(nextRow, n + 1).Value = ts
    wsUnk.Cells(nextRow, n + 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
End Sub

Public Sub CopyBookwiseRowToMismatch(wsBook As Worksheet, wsMis As Worksheet, bkRow As Long, lastColBook As Long, _
                                      bkUrCol As Long, bkDateCol As Long, bkTimeCol As Long, bkLocCol As Long, _
                                      diffMRN As Boolean, diffDate As Boolean, diffTime As Boolean, diffSite As Boolean, ts As Date)
    Dim nextRow As Long, j As Long, cv As Variant
    nextRow = wsMis.Cells(wsMis.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2
    For j = 1 To lastColBook
        If j = bkDateCol Then
            wsMis.Cells(nextRow, j).NumberFormat = "@": wsMis.Cells(nextRow, j).Value = NormaliseDateText(wsBook.Cells(bkRow, j).Value)
        ElseIf j = bkTimeCol Then
            wsMis.Cells(nextRow, j).NumberFormat = "@": wsMis.Cells(nextRow, j).Value = NormaliseTimeText(wsBook.Cells(bkRow, j).Value)
        Else
            cv = wsBook.Cells(bkRow, j).Value
            If IsError(cv) Then wsMis.Cells(nextRow, j).Value = "" Else wsMis.Cells(nextRow, j).Value = cv
        End If
    Next j
    If diffMRN Then wsMis.Cells(nextRow, bkUrCol).Interior.Color = CLR_YELLOW
    If diffDate Then wsMis.Cells(nextRow, bkDateCol).Interior.Color = CLR_YELLOW
    If diffTime Then wsMis.Cells(nextRow, bkTimeCol).Interior.Color = CLR_YELLOW
    If diffSite Then wsMis.Cells(nextRow, bkLocCol).Interior.Color = CLR_YELLOW
    wsMis.Cells(nextRow, lastColBook + 1).Value = ts
    wsMis.Cells(nextRow, lastColBook + 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
End Sub

Public Sub CopyIPMRowToMismatch(wsIPM As Worksheet, wsIMis As Worksheet, ipRow As Long, lastColIPM As Long, _
                                 ipPatCol As Long, ipDateCol As Long, ipTimeCol As Long, ipLocCol As Long, _
                                 diffMRN As Boolean, diffDate As Boolean, diffTime As Boolean, diffSite As Boolean, ts As Date)
    Dim nextRow As Long, j As Long, cv As Variant
    nextRow = wsIMis.Cells(wsIMis.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2
    For j = 1 To lastColIPM
        If j = ipDateCol Then
            wsIMis.Cells(nextRow, j).NumberFormat = "@": wsIMis.Cells(nextRow, j).Value = NormaliseDateText(wsIPM.Cells(ipRow, j).Value)
        ElseIf j = ipTimeCol Then
            wsIMis.Cells(nextRow, j).NumberFormat = "@": wsIMis.Cells(nextRow, j).Value = NormaliseTimeText(wsIPM.Cells(ipRow, j).Value)
        Else
            cv = wsIPM.Cells(ipRow, j).Value
            If IsError(cv) Then wsIMis.Cells(nextRow, j).Value = "" Else wsIMis.Cells(nextRow, j).Value = cv
        End If
    Next j
    If diffMRN Then wsIMis.Cells(nextRow, ipPatCol).Interior.Color = CLR_YELLOW
    If diffDate Then wsIMis.Cells(nextRow, ipDateCol).Interior.Color = CLR_YELLOW
    If diffTime Then wsIMis.Cells(nextRow, ipTimeCol).Interior.Color = CLR_YELLOW
    If diffSite Then wsIMis.Cells(nextRow, ipLocCol).Interior.Color = CLR_YELLOW
    wsIMis.Cells(nextRow, lastColIPM + 1).Value = ts
    wsIMis.Cells(nextRow, lastColIPM + 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
End Sub


' ============================================================
'  GetOrCreateAuditSheet
' ============================================================
Public Function GetOrCreateAuditSheet(wb As Workbook) As Worksheet
    Const AUDIT_NAME As String = "Reconcile Audit Log"
    Dim ws As Worksheet, s As Worksheet
    For Each s In wb.Worksheets
        If s.Name = AUDIT_NAME Then Set GetOrCreateAuditSheet = s: Exit Function
    Next s
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = AUDIT_NAME: ws.Visible = xlSheetVisible
    With ws
        .Cells(1, 1).Value = "Run Date/Time": .Cells(1, 2).Value = "Run By (Username)"
        .Cells(1, 3).Value = "Bookwise File Path": .Cells(1, 4).Value = "iPM File Path"
        .Cells(1, 5).Value = "Total Processed": .Cells(1, 6).Value = "Bookwise Compared"
        .Cells(1, 7).Value = "iPM Compared": .Cells(1, 8).Value = "Clean Matches"
        .Cells(1, 9).Value = "Mismatches": .Cells(1, 10).Value = "Unknown IDs"
        .Cells(1, 11).Value = "Manual Review Total": .Cells(1, 12).Value = "Missing ID"
        .Cells(1, 13).Value = "Duplicate ID": .Cells(1, 14).Value = "Both IDs Present"
        .Cells(1, 15).Value = "Unknown Location": .Cells(1, 16).Value = "Status"
        .Rows(1).Font.Bold = True
        .Columns(1).NumberFormat = "dd/mm/yyyy hh:mm:ss": .Columns(1).ColumnWidth = 20
        .Columns(2).ColumnWidth = 22: .Columns(3).ColumnWidth = 60: .Columns(4).ColumnWidth = 60
    End With
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Set GetOrCreateAuditSheet = ws
End Function

