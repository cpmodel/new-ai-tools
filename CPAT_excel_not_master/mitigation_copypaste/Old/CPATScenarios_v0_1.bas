Attribute VB_Name = "CPATScenarios"
Option Explicit
'==========================================================================
' CPAT-AI-Mitigation-MVP: batch runs of scenario definitions into StoredResults (values).
'
' MTInputs holds one column per scenario: J = scenario 1 (baseline), K = scenario 2 (the live policy
' scenario that Mitigation group 2 calculates), L onwards = scenario definitions (row 4 "Run?" = Yes/No,
' row 5 number, row 6 name, rows 8-415 the inputs).
'
' Public macros (Alt+F8):
'   RunAllScenarios       stores the baseline (only if not stored yet), then for every definition with
'                         Run? = Yes: copies it into the live column K, recalculates, runs the ETS goal seek
'                         when the definition applies a new ETS (and does not set "Override ETS price" =
'                         Yes), and stores the MTOutputs block of scenario 2 as values under the definition's
'                         number and name. Column K and the ETS override row are restored at the end.
'   StoreBaseline         (re)stores scenario 1 under ID 1 (replaces the stored baseline).
'   StoreLiveScenario     stores the live scenario 2 as it stands, under an ID you choose.
'   SolveETSLive          ETS goal seek for the live scenario only (fills its override row ets.ovr).
'   ClearStoredScenarios  deletes every stored scenario except the baseline.
'
' ETS goal seek: legacy damped log-space iteration (OverrideETSFast / SolveFast) on the workbook's proposal
' row ets.next: p <- p x (proposal / p) ^ alpha, mixing 0.3 with the previous iterate, 5% centred 3-year
' smoothing, adaptive alpha; stops when the worst |covered / cap - 1| is below ETS_TOL. Same steps as
' ets_goalseek_v0_2.py.
'==========================================================================

Private Const SH_MT As String = "MTInputs"
Private Const SH_MIT As String = "Mitigation"
Private Const SH_OUT As String = "MTOutputs"
Private Const SH_SR As String = "StoredResults"
Private Const MT_ROW_RUN As Long = 4
Private Const MT_ROW_NUM As Long = 5
Private Const MT_ROW_NAME As Long = 6
Private Const MT_FIRST As Long = 8
Private Const MT_LAST As Long = 415
Private Const MT_ROW_ID As Long = 10          ' formula rows of every scenario column (kept in column K)
Private Const MT_ROW_IDNAME As Long = 12
Private Const MT_COL_PARAM As Long = 8        ' H: NameOfParameter
Private Const MT_COL_FIRST As Long = 10       ' J: scenario 1
Private Const MIT_ROW_SCEN As Long = 5
Private Const MIT_COL_FIRST As Long = 11      ' K: code column of scenario group 1
Private Const MIT_MAX_ROW As Long = 5000
Private Const OUT_COL_Y0 As Long = 6          ' MTOutputs: first year column (F)
Private Const SR_FIRST As Long = 6            ' StoredResults: first data row
Private Const SR_COL_Y0 As Long = 8           ' StoredResults: first year column (H)
Private Const NYEARS As Long = 19             ' 2022-2040
Private Const LIVE_SCEN As Long = 2
Private Const BASE_SCEN As Long = 1
Private Const ETS_TOL As Double = 0.005
Private Const ETS_MAXITER As Long = 20

'---------------------------------------------------------------- public macros
Public Sub RunAllScenarios()
    Dim n As Long
    n = RunAll()
    MsgBox n & " scenario definition(s) run and stored in " & SH_SR & ".", vbInformation, "CPAT scenarios"
End Sub

Public Sub StoreBaseline()
    If MsgBox("Replace the stored baseline (ID 1) by the current scenario 1 results?", vbYesNo + vbQuestion, _
              "CPAT scenarios") <> vbYes Then Exit Sub
    Application.Calculate
    StoreScenario BASE_SCEN, BASE_SCEN, ""
End Sub

Public Sub StoreLiveScenario()
    Dim v As Variant
    v = InputBox("Store the live scenario 2 under which ID? (an existing ID is replaced)", "CPAT scenarios", _
                 CStr(NextFreeId()))
    If Len(CStr(v)) = 0 Or Not IsNumeric(v) Then Exit Sub
    Application.Calculate
    StoreScenario CLng(v), LIVE_SCEN, CStr(ThisWorkbook.Worksheets(SH_MT).Cells(MT_ROW_NAME, MTColumn(LIVE_SCEN)).Value)
End Sub

Public Sub SolveETSLive()
    Dim w As Double
    w = SolveETS()
    MsgBox "ETS goal seek done: worst |covered / cap - 1| = " & Format(w, "0.0000"), vbInformation, "CPAT scenarios"
End Sub

Public Sub ClearStoredScenarios()
    If MsgBox("Delete every stored scenario except the baseline (ID 1)?", vbYesNo + vbQuestion, _
              "CPAT scenarios") <> vbYes Then Exit Sub
    KeepOnlyBaseline
End Sub

'---------------------------------------------------------------- batch run (no dialogs: also used by tests)
Public Function RunAll() As Long
    Dim mt As Worksheet, liveCol As Long, c As Long, r As Long, n As Long
    Dim savedVal(1 To MT_LAST) As Variant, savedIsF(1 To MT_LAST) As Boolean
    Dim ovr As Long, ovrBase As Long, savedOvr(1 To NYEARS) As Variant, j As Long
    Set mt = ThisWorkbook.Worksheets(SH_MT)
    liveCol = MTColumn(LIVE_SCEN)
    For r = 1 To MT_LAST                                  ' save the live column (formulas and values)
        savedIsF(r) = mt.Cells(r, liveCol).HasFormula
        If savedIsF(r) Then savedVal(r) = mt.Cells(r, liveCol).Formula Else savedVal(r) = mt.Cells(r, liveCol).Value
    Next r
    ovr = MitRow("ets.ovr")
    ovrBase = MitBaseCol(LIVE_SCEN)
    For j = 1 To NYEARS
        savedOvr(j) = ThisWorkbook.Worksheets(SH_MIT).Cells(ovr, ovrBase + j - 1).Value
    Next j
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo 0
    Application.Calculate
    If FindStored(BASE_SCEN) = 0 Then StoreScenario BASE_SCEN, BASE_SCEN, ""
    c = liveCol + 1
    Do While Not IsEmpty(mt.Cells(MT_ROW_NUM, c).Value)
        If UCase(Left(Trim(CStr(mt.Cells(MT_ROW_RUN, c).Value)), 1)) = "Y" Then
            For r = MT_FIRST To MT_LAST
                If r <> MT_ROW_ID And r <> MT_ROW_IDNAME Then mt.Cells(r, liveCol).Value = mt.Cells(r, c).Value
            Next r
            mt.Cells(MT_ROW_NAME, liveCol).Value = mt.Cells(MT_ROW_NAME, c).Value
            For j = 1 To NYEARS
                ThisWorkbook.Worksheets(SH_MIT).Cells(ovr, ovrBase + j - 1).ClearContents
            Next j
            Application.Calculate
            If NeedsETS(liveCol) Then SolveETS
            StoreScenario CLng(mt.Cells(MT_ROW_NUM, c).Value), LIVE_SCEN, CStr(mt.Cells(MT_ROW_NAME, c).Value)
            n = n + 1
        End If
        c = c + 1
    Loop
    For r = 1 To MT_LAST                                  ' restore the live column and the override row
        If savedIsF(r) Then
            mt.Cells(r, liveCol).Formula = savedVal(r)
        ElseIf IsEmpty(savedVal(r)) Then
            mt.Cells(r, liveCol).ClearContents
        Else
            mt.Cells(r, liveCol).Value = savedVal(r)
        End If
    Next r
    For j = 1 To NYEARS
        If IsEmpty(savedOvr(j)) Then
            ThisWorkbook.Worksheets(SH_MIT).Cells(ovr, ovrBase + j - 1).ClearContents
        Else
            ThisWorkbook.Worksheets(SH_MIT).Cells(ovr, ovrBase + j - 1).Value = savedOvr(j)
        End If
    Next j
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    On Error GoTo 0
    Application.Calculate
    RunAll = n
End Function

'---------------------------------------------------------------- ETS goal seek on the live scenario
Public Function SolveETS() As Double
    Dim mit As Worksheet, mt As Worksheet, liveCol As Long, rO As Long, rE As Long, rN As Long, rG As Long, rC As Long
    Dim b As Long, j As Long, it As Long, hasOvr As Boolean, hasPrev As Boolean
    Dim p(1 To NYEARS) As Double, pPrev(1 To NYEARS) As Double, pNew(1 To NYEARS) As Double
    Dim resp(1 To NYEARS) As Double, gap As Double, cap As Double, sm As Double
    Dim alpha As Double, worst As Double, prevWorst As Double, converged As Boolean
    Set mit = ThisWorkbook.Worksheets(SH_MIT)
    Set mt = ThisWorkbook.Worksheets(SH_MT)
    liveCol = MTColumn(LIVE_SCEN)
    mt.Cells(MTParamRow("D_ETSPriceOverride"), liveCol).Value = "Yes"
    rO = MitRow("ets.ovr"): rE = MitRow("ets.est"): rN = MitRow("ets.next")
    rG = MitRow("co2.gap"): rC = MitRow("co2.cap")
    b = MitBaseCol(LIVE_SCEN)
    Application.Calculate
    For j = 1 To NYEARS
        p(j) = Num(mit.Cells(rO, b + j - 1).Value)
        If p(j) <> 0 Then hasOvr = True
    Next j
    If Not hasOvr Then
        For j = 1 To NYEARS
            p(j) = Num(mit.Cells(rE, b + j - 1).Value)
        Next j
    End If
    alpha = 1#
    prevWorst = 1E+300
    For it = 1 To ETS_MAXITER
        For j = 1 To NYEARS
            mit.Cells(rO, b + j - 1).Value = p(j)
        Next j
        Application.Calculate
        worst = 0#
        For j = 1 To NYEARS
            resp(j) = Num(mit.Cells(rN, b + j - 1).Value)
            gap = Num(mit.Cells(rG, b + j - 1).Value)
            cap = Num(mit.Cells(rC, b + j - 1).Value)
            If cap > 0 And Abs(gap) > worst Then worst = Abs(gap)
        Next j
        If worst < ETS_TOL Then
            converged = True
            Exit For
        End If
        For j = 1 To NYEARS
            pNew(j) = LogStep(p(j), resp(j), alpha)
            If hasPrev Then pNew(j) = 0.7 * pNew(j) + 0.3 * pPrev(j)
        Next j
        For j = 1 To NYEARS                               ' 5% weight on a centred 3-year moving average
            If j = 1 Then
                sm = (pNew(1) + pNew(2)) / 2
            ElseIf j = NYEARS Then
                sm = (pNew(NYEARS - 1) + pNew(NYEARS)) / 2
            Else
                sm = (pNew(j - 1) + pNew(j) + pNew(j + 1)) / 3
            End If
            resp(j) = 0.95 * pNew(j) + 0.05 * sm          ' resp reused as the smoothed vector
        Next j
        For j = 1 To NYEARS
            If resp(j) < 0 Then resp(j) = p(j)
        Next j
        If worst > prevWorst Then
            alpha = alpha * 0.5
            If alpha < 0.0625 Then alpha = 0.0625
        ElseIf worst < 0.5 * prevWorst Then
            alpha = alpha * 1.2
            If alpha > 1# Then alpha = 1#
        End If
        prevWorst = worst
        For j = 1 To NYEARS
            pPrev(j) = p(j)
            p(j) = resp(j)
        Next j
        hasPrev = True
    Next it
    If Not converged Then                                 ' leave the last step's prices in place
        For j = 1 To NYEARS
            mit.Cells(rO, b + j - 1).Value = p(j)
        Next j
        Application.Calculate
    End If
    SolveETS = worst
End Function

Private Function LogStep(ByVal pc As Double, ByVal pr As Double, ByVal alpha As Double) As Double
    If pc <= 0 Or pr <= 0 Then
        LogStep = pc + alpha * (pr - pc)
        If LogStep < 0 Then LogStep = 0
    Else
        LogStep = pc * (pr / pc) ^ alpha
    End If
End Function

'---------------------------------------------------------------- storing results as values
Private Sub StoreScenario(ByVal id As Long, ByVal srcScen As Long, ByVal nm As String)
    Dim mo As Worksheet, sr As Worksheet, h As Long, r As Long, n As Long, d As Long, k As Long, j As Long
    Dim code As String, country As String, stamp As Date
    Set mo = ThisWorkbook.Worksheets(SH_OUT)
    Set sr = ThisWorkbook.Worksheets(SH_SR)
    For r = 5 To 2000                                     ' header row of the MTOutputs block
        If CStr(mo.Cells(r, 1).Value) = "Scenario block" Then
            If Num(mo.Cells(r, 5).Value) = srcScen Then
                h = r
                Exit For
            End If
        End If
    Next r
    If h = 0 Then Err.Raise vbObjectError + 1, , "No MTOutputs block for scenario " & srcScen
    If Len(nm) = 0 Then nm = CStr(mo.Cells(h, 3).Value)
    Do While Len(CStr(mo.Cells(h + 1 + n, 2).Value)) > 0
        n = n + 1
    Loop
    d = FindStored(id)
    If d = 0 Then d = FirstFreeRow()
    code = CStr(mo.Cells(h + 1, 1).Value)
    country = Left(code, InStr(code, ".") - 1)
    stamp = Now
    For k = 0 To n - 1
        sr.Cells(d + k, 1).Value = id
        sr.Cells(d + k, 2).Value = nm
        sr.Cells(d + k, 3).Value = mo.Cells(h + 1 + k, 2).Value
        sr.Cells(d + k, 4).Value = country & "." & CStr(mo.Cells(h + 1 + k, 2).Value) & "." & id
        sr.Cells(d + k, 5).Value = mo.Cells(h + 1 + k, 3).Value
        sr.Cells(d + k, 6).Value = mo.Cells(h + 1 + k, 4).Value
        sr.Cells(d + k, 7).Value = stamp
        For j = 1 To NYEARS
            sr.Cells(d + k, SR_COL_Y0 + j - 1).Value = mo.Cells(h + 1 + k, OUT_COL_Y0 + j - 1).Value
        Next j
    Next k
End Sub

Private Sub KeepOnlyBaseline()
    Dim sr As Worksheet, r As Long, w As Long, c As Long, last As Long
    Set sr = ThisWorkbook.Worksheets(SH_SR)
    last = FirstFreeRow() - 1
    w = SR_FIRST
    For r = SR_FIRST To last
        If Num(sr.Cells(r, 1).Value) = BASE_SCEN Then
            If w <> r Then
                For c = 1 To SR_COL_Y0 + NYEARS - 1
                    sr.Cells(w, c).Value = sr.Cells(r, c).Value
                Next c
            End If
            w = w + 1
        End If
    Next r
    If w <= last Then sr.Range(sr.Cells(w, 1), sr.Cells(last, SR_COL_Y0 + NYEARS - 1)).ClearContents
End Sub

'---------------------------------------------------------------- lookups
Private Function MTColumn(ByVal scen As Long) As Long
    Dim c As Long
    For c = MT_COL_FIRST To MT_COL_FIRST + 200
        If Num(ThisWorkbook.Worksheets(SH_MT).Cells(MT_ROW_NUM, c).Value) = scen Then
            MTColumn = c
            Exit Function
        End If
    Next c
    Err.Raise vbObjectError + 2, , "Scenario " & scen & " not found in MTInputs row 5"
End Function

Private Function MTParamRow(ByVal nm As String) As Long
    Dim r As Long
    For r = MT_FIRST To MT_LAST
        If CStr(ThisWorkbook.Worksheets(SH_MT).Cells(r, MT_COL_PARAM).Value) = nm Then
            MTParamRow = r
            Exit Function
        End If
    Next r
    Err.Raise vbObjectError + 3, , "Parameter " & nm & " not found in MTInputs column H"
End Function

Private Function NeedsETS(ByVal liveCol As Long) As Boolean
    Dim mt As Worksheet
    Set mt = ThisWorkbook.Worksheets(SH_MT)
    NeedsETS = UCase(Left(CStr(mt.Cells(MTParamRow("D_NewETS"), liveCol).Value), 3)) = "YES" And _
               UCase(Left(CStr(mt.Cells(MTParamRow("D_ETSPriceOverride"), liveCol).Value), 3)) <> "YES"
End Function

Private Function MitRow(ByVal code As String) As Long
    Dim r As Long
    For r = 1 To MIT_MAX_ROW
        If CStr(ThisWorkbook.Worksheets(SH_MIT).Cells(r, 1).Value) = code Then
            MitRow = r
            Exit Function
        End If
    Next r
    Err.Raise vbObjectError + 4, , "Row " & code & " not found in Mitigation column A"
End Function

Private Function MitBaseCol(ByVal scen As Long) As Long
    Dim c As Long
    For c = MIT_COL_FIRST To MIT_COL_FIRST + 2000
        If Num(ThisWorkbook.Worksheets(SH_MIT).Cells(MIT_ROW_SCEN, c).Value) = scen Then
            MitBaseCol = c + 1                            ' code column + 1 = base year
            Exit Function
        End If
    Next c
    Err.Raise vbObjectError + 5, , "Scenario group " & scen & " not found in Mitigation row 5"
End Function

Private Function FindStored(ByVal id As Long) As Long
    Dim sr As Worksheet, r As Long
    Set sr = ThisWorkbook.Worksheets(SH_SR)
    r = SR_FIRST
    Do While Len(CStr(sr.Cells(r, 3).Value)) > 0
        If Num(sr.Cells(r, 1).Value) = id Then
            FindStored = r
            Exit Function
        End If
        r = r + 1
    Loop
End Function

Private Function FirstFreeRow() As Long
    Dim sr As Worksheet, r As Long
    Set sr = ThisWorkbook.Worksheets(SH_SR)
    r = SR_FIRST
    Do While Len(CStr(sr.Cells(r, 3).Value)) > 0
        r = r + 1
    Loop
    FirstFreeRow = r
End Function

Private Function NextFreeId() As Long
    Dim sr As Worksheet, r As Long, m As Long
    Set sr = ThisWorkbook.Worksheets(SH_SR)
    m = 99
    r = SR_FIRST
    Do While Len(CStr(sr.Cells(r, 3).Value)) > 0
        If Num(sr.Cells(r, 1).Value) > m Then m = Num(sr.Cells(r, 1).Value)
        r = r + 1
    Loop
    NextFreeId = m + 1
End Function

Private Function Num(ByVal v As Variant) As Double
    If IsError(v) Then
        Num = 0
    ElseIf IsNumeric(v) And Not IsEmpty(v) Then
        Num = CDbl(v)
    Else
        Num = 0
    End If
End Function
