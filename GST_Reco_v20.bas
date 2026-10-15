Option Explicit

'----------------------------------------------------------------
' GST RECONCILIATION v20
' Created by Chinmay Bhat
'
' v20 FIXES (on top of v19):
'   1) Invoice normaliser no longer treats a bare 2-digit / 4-digit
'      tail as a financial year ("ABC/12" and "ABC/13" stay distinct).
'      Tails are stripped only when clearly a FY (INV/123/25, INV/24-25).
'   2) Matching is GSTIN-scoped (GSTIN|InvoiceKey). A cross-GSTIN
'      fallback runs LAST, on the exact invoice key only, and only if
'      taxable value is within tolerance (catches vendor GSTIN typos).
'   3) Duplicates are detected BEFORE matching - every source row
'      appears exactly once, so counts and totals are not inflated.
'   4) Amount tolerance constant GST_TOL (set to 0 = exact to the paisa;
'      raise to e.g. 1 to ignore small rounding). Round-half-away-from-zero
'      instead of banker's rounding.
'   5) Colour constants corrected (VBA colours are BGR, not RGB).
'   6) O(n) counting sort replaces the O(n^2) bubble sort.
'   7) Dates parsed as dd/mm/yyyy explicitly (no PC-locale surprises).
'   8) Invoice keys use digit-only checks (IsNumeric accepted "1E5").
'   9) Existing result sheets are only deleted after input validation;
'      calculation mode is restored to what the user had.
'  10) New pass: same GSTIN + same date + amount within tolerance, invoice
'      numbers written differently (Portal 4546 vs Books 26/4546).
'      Reported as Matched with note "Inv No format differs".
'
' BEFORE IMPORT: Delete ALL old GST modules in VBA Editor!
' IMPORT : File > Import File > this .bas
' RUN    : Alt+F8 > GST_Run > Run
' NOTE   : Windows Excel only (uses Scripting.Dictionary).
'----------------------------------------------------------------

'----------------------------------------------------------------
' COLUMN LAYOUT  (1-based)
'  1=Status  2=MismatchDetail  3=GSTIN
'  4=PartyName(Portal)    5=PartyName(Books)
'  6=InvDate(Portal)      7=InvDate(Books)
'  8=InvNo(Portal)        9=InvNo(Books)
' 10=TaxableVal(Portal)  11=TaxableVal(Books)  12=DiffTaxable
' 13=IGST(Portal)        14=IGST(Books)        15=DiffIGST
' 16=CGST(Portal)        17=CGST(Books)        18=DiffCGST
' 19=SGST(Portal)        20=SGST(Books)        21=DiffSGST
' 22=CESS(Portal)        23=CESS(Books)        24=DiffCESS
'----------------------------------------------------------------

'==================================================================
' COLOUR PALETTE
'==================================================================
Public Const GST_CH As Long = 10053171    ' #336699 navy-blue header
Public Const GST_CM As Long = 13037255    ' #C7EEC6 fresh green   (Matched)
Public Const GST_CA As Long = 12904703    ' #FFE8C4 amber-cream   (Mismatch-Amt)
Public Const GST_CO As Long = 13889011    ' #F3EDD3 golden wheat  (Mismatch-Other)
Public Const GST_CP As Long = 15068411    ' #FBECE5 peach-rose    (Missing Portal)
Public Const GST_CT As Long = 15918811    ' #DBE6F2 cornflower    (Missing Books)
Public Const GST_CD As Long = 15392229    ' #E5DDEA soft mauve    (Duplicate)
Public Const GST_CS As Long = 12090980    ' #647EB8 periwinkle    (Summary col hdrs)

Public Const GST_NC  As Long = 24
Public Const GST_SNC As Long = 18
Public Const GST_THR As Double = 0.82
Public Const GST_TOL As Double = 0#       ' amount tolerance in Rs (0 = exact to the paisa)
Public GST_BnrRow(0 To 7) As Long

'==================================================================
' MAIN ENTRY POINT
'==================================================================
Sub GST_Run()
    Dim GST_rnWsP  As Worksheet
    Dim GST_rnWsT  As Worksheet
    Dim GST_rnWsR  As Worksheet
    Dim GST_rnWsS  As Worksheet
    Dim GST_rnNP   As Long
    Dim GST_rnNT   As Long
    Dim GST_rnNR   As Long
    Dim GST_rnAP() As Variant
    Dim GST_rnAT() As Variant
    Dim GST_rnAR() As Variant
    Dim GST_rnIdx  As Long
    Dim GST_rnErr  As Boolean
    Dim GST_rnCalc As XlCalculation

    On Error GoTo GST_rnEH
    GST_rnCalc = Application.Calculation

    If Not GST_HasSheet("As per Portal") Then
        MsgBox "Sheet [As per Portal] not found!", vbCritical, "GST Reco"
        Exit Sub
    End If
    If Not GST_HasSheet("As per Books") Then
        MsgBox "Sheet [As per Books] not found!", vbCritical, "GST Reco"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    For GST_rnIdx = 0 To 7
        GST_BnrRow(GST_rnIdx) = 0
    Next GST_rnIdx

    Set GST_rnWsP = ThisWorkbook.Sheets("As per Portal")
    Set GST_rnWsT = ThisWorkbook.Sheets("As per Books")

    GST_rnNP = GST_LastRow(GST_rnWsP)
    GST_rnNT = GST_LastRow(GST_rnWsT)

    ' Validate BEFORE touching existing result sheets
    If GST_rnNP < 2 Then
        MsgBox "No data in As per Portal!", vbExclamation, "GST Reco"
        GoTo GST_rnCL
    End If
    If GST_rnNT < 2 Then
        MsgBox "No data in As per Books!", vbExclamation, "GST Reco"
        GoTo GST_rnCL
    End If

    Set GST_rnWsR = GST_MakeSheet()
    Set GST_rnWsS = GST_MakeSumSheet()

    GST_rnAP = GST_rnWsP.Range("A1:I" & CStr(GST_rnNP)).Value
    GST_rnAT = GST_rnWsT.Range("A1:I" & CStr(GST_rnNT)).Value

    Call GST_DoReco(GST_rnAP, GST_rnNP, GST_rnAT, GST_rnNT, GST_rnAR, GST_rnNR)
    Call GST_WriteRows(GST_rnWsR, GST_rnAR, GST_rnNR)
    Call GST_WriteSumSheet(GST_rnWsS, GST_rnAR, GST_rnNR)

    GST_rnWsR.Columns("A:X").AutoFit
    GST_rnWsR.Columns("A").ColumnWidth = 22
    GST_rnWsR.Columns("B").ColumnWidth = 30
    GST_rnWsR.Columns("C").ColumnWidth = 18
    GST_rnWsR.Columns("D").ColumnWidth = 22
    GST_rnWsR.Columns("E").ColumnWidth = 22
    GST_rnWsR.Columns("L").ColumnWidth = 16
    GST_rnWsR.Columns("O").ColumnWidth = 14
    GST_rnWsR.Columns("R").ColumnWidth = 14
    GST_rnWsR.Columns("U").ColumnWidth = 14
    GST_rnWsR.Columns("X").ColumnWidth = 14

    GST_rnWsS.Activate
    GST_rnWsS.Range("A1").Select

GST_rnCL:
    Application.ScreenUpdating = True
    Application.Calculation = GST_rnCalc
    Application.EnableEvents = True
    If GST_rnNR > 0 And Not GST_rnErr Then
        MsgBox "Done! " & CStr(GST_rnNR) & " records." & vbCrLf & "See GST_Summary sheet.", _
               vbInformation, "GST Reco"
    End If
    Exit Sub

GST_rnEH:
    GST_rnErr = True
    Application.ScreenUpdating = True
    Application.Calculation = GST_rnCalc
    Application.EnableEvents = True
    MsgBox "Error " & CStr(Err.Number) & ": " & Err.Description, vbCritical, "GST Reco"
    Resume GST_rnCL
End Sub

'==================================================================
' RECONCILIATION CORE
'==================================================================
Sub GST_DoReco(aP() As Variant, nP As Long, aT() As Variant, nT As Long, _
               aR() As Variant, nR As Long)

    Dim GST_drBuf    As Long
    Dim GST_drRow    As Long
    Dim GST_drKey    As String
    Dim GST_drKey2   As String
    Dim GST_drKey3   As String
    Dim GST_drKey4   As String
    Dim GST_drBest   As Long
    Dim GST_drBDiff  As Double
    Dim GST_drPG     As String
    Dim GST_drPP     As String
    Dim GST_drPD     As String
    Dim GST_drPI     As String
    Dim GST_drPISfx  As String
    Dim GST_drPIPfx  As String
    Dim GST_drPINum  As String
    Dim GST_drPTV    As Double
    Dim GST_drPIG    As Double
    Dim GST_drPCG    As Double
    Dim GST_drPSG    As Double
    Dim GST_drPCS    As Double
    Dim GST_drTG     As String
    Dim GST_drTP     As String
    Dim GST_drTD     As String
    Dim GST_drTTV    As Double
    Dim GST_drTIG    As Double
    Dim GST_drTCG    As Double
    Dim GST_drTSG    As Double
    Dim GST_drTCS    As Double
    Dim GST_drDet    As String
    Dim GST_drAmt    As Boolean
    Dim GST_drOth    As Boolean
    Dim GST_drDicT   As Object
    Dim GST_drDicTX  As Object
    Dim GST_drDicPD  As Object
    Dim GST_drDicTD  As Object
    Dim GST_drDicTL  As Object
    Dim GST_drKey5   As String
    Dim GST_drPF()   As Boolean
    Dim GST_drUsed() As Boolean
    Dim GST_drPDup() As Boolean
    Dim GST_drTDup() As Boolean
    Dim GST_drPM()   As Long

    ' Every source row yields at most ONE result record
    GST_drBuf = CLng(nP) + CLng(nT) + 10
    ReDim aR(1 To GST_drBuf, 1 To 24)
    ReDim GST_drUsed(1 To nT)
    ReDim GST_drTDup(1 To nT)
    ReDim GST_drPDup(1 To nP)
    ReDim GST_drPM(1 To nP)
    ReDim GST_drPF(1 To nP)
    nR = 0

    '--------------------------------------------------------------
    ' STEP 1: Duplicate detection (BEFORE matching)
    '   Same GSTIN + same normalised invoice no. = duplicate.
    '   First occurrence is kept for matching; later ones are flagged.
    '   Blank invoice numbers are never treated as duplicates.
    '--------------------------------------------------------------
    Set GST_drDicPD = CreateObject("Scripting.Dictionary")
    For GST_drRow = 2 To nP
        GST_drKey = GST_NormInv(CStr(aP(GST_drRow, 4)))
        If Len(GST_drKey) > 0 Then
            GST_drKey = UCase(Trim(CStr(aP(GST_drRow, 1)))) & "||" & GST_drKey
            If GST_drDicPD.Exists(GST_drKey) Then
                GST_drPDup(GST_drRow) = True
            Else
                GST_drDicPD.Add GST_drKey, GST_drRow
            End If
        End If
    Next GST_drRow

    Set GST_drDicTD = CreateObject("Scripting.Dictionary")
    For GST_drRow = 2 To nT
        GST_drKey = GST_NormInv(CStr(aT(GST_drRow, 4)))
        If Len(GST_drKey) > 0 Then
            GST_drKey = UCase(Trim(CStr(aT(GST_drRow, 1)))) & "||" & GST_drKey
            If GST_drDicTD.Exists(GST_drKey) Then
                GST_drTDup(GST_drRow) = True
            Else
                GST_drDicTD.Add GST_drKey, GST_drRow
            End If
        End If
    Next GST_drRow

    '--------------------------------------------------------------
    ' STEP 2: Build Books dictionaries (non-duplicate rows only)
    '   DicT  : GSTIN-scoped, 4 keys per invoice
    '             GSTIN|exact   GSTIN|trailing-num   GSTIN|leading-alpha-stripped
    '             GSTIN|pure-numeric
    '   DicTX : exact normalised key only (no GSTIN) - last-resort fallback
    '--------------------------------------------------------------
    Set GST_drDicT = CreateObject("Scripting.Dictionary")
    Set GST_drDicTX = CreateObject("Scripting.Dictionary")
    Set GST_drDicTL = CreateObject("Scripting.Dictionary")
    For GST_drRow = 2 To nT
        If Not GST_drTDup(GST_drRow) Then
            GST_drKey = GST_NormInv(CStr(aT(GST_drRow, 4)))
            If Len(GST_drKey) > 0 Then
                GST_drTG = UCase(Trim(CStr(aT(GST_drRow, 1)))) & "|"

                Call GST_DicAdd(GST_drDicTX, GST_drKey, CStr(GST_drRow))
                GST_drKey5 = GST_LastSeg(CStr(aT(GST_drRow, 4)))
                If Len(GST_drKey5) > 0 Then Call GST_DicAdd(GST_drDicTL, GST_drTG & GST_drKey5, CStr(GST_drRow))
                Call GST_DicAdd(GST_drDicT, GST_drTG & GST_drKey, CStr(GST_drRow))

                GST_drKey2 = GST_InvNumSuffix(GST_drKey)
                If Len(GST_drKey2) > 0 And GST_drKey2 <> GST_drKey Then
                    Call GST_DicAdd(GST_drDicT, GST_drTG & GST_drKey2, CStr(GST_drRow))
                End If

                GST_drKey3 = GST_InvNumPrefix(GST_drKey)
                If Len(GST_drKey3) > 0 And GST_drKey3 <> GST_drKey And GST_drKey3 <> GST_drKey2 Then
                    Call GST_DicAdd(GST_drDicT, GST_drTG & GST_drKey3, CStr(GST_drRow))
                End If

                GST_drKey4 = GST_ExtractPureNumeric(CStr(aT(GST_drRow, 4)))
                If Len(GST_drKey4) > 0 _
                   And GST_drKey4 <> GST_drKey _
                   And GST_drKey4 <> GST_drKey2 _
                   And GST_drKey4 <> GST_drKey3 Then
                    Call GST_DicAdd(GST_drDicT, GST_drTG & GST_drKey4, CStr(GST_drRow))
                End If
            End If
        End If
    Next GST_drRow

    '--------------------------------------------------------------
    ' STEP 3a: GSTIN-scoped matching, 4 passes, closest taxable wins
    '--------------------------------------------------------------
    For GST_drRow = 2 To nP
        If Not GST_drPDup(GST_drRow) Then
            GST_drPI = GST_NormInv(CStr(aP(GST_drRow, 4)))
            If Len(GST_drPI) > 0 Then
                GST_drPG = UCase(Trim(CStr(aP(GST_drRow, 1)))) & "|"
                GST_drPTV = GST_ToD(aP(GST_drRow, 5))
                GST_drPISfx = GST_InvNumSuffix(GST_drPI)
                GST_drPIPfx = GST_InvNumPrefix(GST_drPI)
                GST_drPINum = GST_ExtractPureNumeric(CStr(aP(GST_drRow, 4)))

                GST_drBest = 0
                GST_drBDiff = 99999999999#

                ' Pass 1: exact normalised key
                Call GST_FindBest(GST_drDicT, GST_drPG & GST_drPI, aT, GST_drUsed, _
                                  GST_drPTV, GST_drBest, GST_drBDiff)

                ' Pass 2: trailing-numeric suffix
                If GST_drBest = 0 And Len(GST_drPISfx) > 0 And GST_drPISfx <> GST_drPI Then
                    Call GST_FindBest(GST_drDicT, GST_drPG & GST_drPISfx, aT, GST_drUsed, _
                                      GST_drPTV, GST_drBest, GST_drBDiff)
                End If

                ' Pass 3: leading-alpha stripped number
                If GST_drBest = 0 And Len(GST_drPIPfx) > 0 _
                   And GST_drPIPfx <> GST_drPI And GST_drPIPfx <> GST_drPISfx Then
                    Call GST_FindBest(GST_drDicT, GST_drPG & GST_drPIPfx, aT, GST_drUsed, _
                                      GST_drPTV, GST_drBest, GST_drBDiff)
                End If

                ' Pass 4: pure numeric extraction (ABC/12 -> 12)
                If GST_drBest = 0 And Len(GST_drPINum) > 0 _
                   And GST_drPINum <> GST_drPI _
                   And GST_drPINum <> GST_drPISfx _
                   And GST_drPINum <> GST_drPIPfx Then
                    Call GST_FindBest(GST_drDicT, GST_drPG & GST_drPINum, aT, GST_drUsed, _
                                      GST_drPTV, GST_drBest, GST_drBDiff)
                End If

                If GST_drBest > 0 Then
                    GST_drPM(GST_drRow) = GST_drBest
                    GST_drUsed(GST_drBest) = True
                End If
            End If
        End If
    Next GST_drRow

    '--------------------------------------------------------------
    ' STEP 3b: Same GSTIN, invoice numbers written differently
    '   e.g. Portal "4546" vs Books "26/4546" (series/FY prefix).
    '   Compares the trailing digit run (4546). Because this is looser,
    '   it also requires SAME invoice date AND taxable within tolerance.
    '--------------------------------------------------------------
    For GST_drRow = 2 To nP
        If Not GST_drPDup(GST_drRow) And GST_drPM(GST_drRow) = 0 Then
            GST_drKey = GST_LastSeg(CStr(aP(GST_drRow, 4)))
            If Len(GST_drKey) > 0 Then
                GST_drPG = UCase(Trim(CStr(aP(GST_drRow, 1)))) & "|"
                GST_drPTV = GST_ToD(aP(GST_drRow, 5))
                GST_drPD = GST_NormDate(aP(GST_drRow, 3))
                GST_drBest = 0
                GST_drBDiff = GST_TOL + 0.001
                Call GST_FindBestD(GST_drDicTL, GST_drPG & GST_drKey, aT, GST_drUsed, _
                                   GST_drPTV, GST_drPD, GST_drBest, GST_drBDiff)
                If GST_drBest > 0 Then
                    GST_drPM(GST_drRow) = GST_drBest
                    GST_drPF(GST_drRow) = True
                    GST_drUsed(GST_drBest) = True
                End If
            End If
        End If
    Next GST_drRow

    '--------------------------------------------------------------
    ' STEP 3c: Cross-GSTIN fallback for rows still unmatched.
    '   Exact invoice key only, and taxable value must be within
    '   tolerance. Result is reported as Mismatch-Other (GSTIN).
    '--------------------------------------------------------------
    For GST_drRow = 2 To nP
        If Not GST_drPDup(GST_drRow) And GST_drPM(GST_drRow) = 0 Then
            GST_drPI = GST_NormInv(CStr(aP(GST_drRow, 4)))
            If Len(GST_drPI) > 0 Then
                GST_drPTV = GST_ToD(aP(GST_drRow, 5))
                GST_drBest = 0
                GST_drBDiff = GST_TOL + 0.001
                Call GST_FindBest(GST_drDicTX, GST_drPI, aT, GST_drUsed, _
                                  GST_drPTV, GST_drBest, GST_drBDiff)
                If GST_drBest > 0 Then
                    GST_drPM(GST_drRow) = GST_drBest
                    GST_drUsed(GST_drBest) = True
                End If
            End If
        End If
    Next GST_drRow

    '--------------------------------------------------------------
    ' STEP 4: Write one result record per Portal row
    '--------------------------------------------------------------
    For GST_drRow = 2 To nP
        nR = nR + 1
        If nR > GST_drBuf Then Exit For

        GST_drPG = UCase(Trim(CStr(aP(GST_drRow, 1))))
        GST_drPP = CStr(aP(GST_drRow, 2))
        GST_drPD = GST_NormDate(aP(GST_drRow, 3))
        GST_drPTV = GST_ToD(aP(GST_drRow, 5))
        GST_drPIG = GST_ToD(aP(GST_drRow, 6))
        GST_drPCG = GST_ToD(aP(GST_drRow, 7))
        GST_drPSG = GST_ToD(aP(GST_drRow, 8))
        GST_drPCS = GST_ToD(aP(GST_drRow, 9))

        aR(nR, 3) = GST_drPG
        aR(nR, 4) = GST_drPP
        aR(nR, 6) = GST_FmtDate(aP(GST_drRow, 3))
        aR(nR, 8) = CStr(aP(GST_drRow, 4))
        aR(nR, 10) = GST_drPTV
        aR(nR, 13) = GST_drPIG
        aR(nR, 16) = GST_drPCG
        aR(nR, 19) = GST_drPSG
        aR(nR, 22) = GST_drPCS

        GST_drBest = GST_drPM(GST_drRow)

        If GST_drPDup(GST_drRow) Then
            aR(nR, 1) = "Duplicate-Portal": aR(nR, 2) = "Dup InvNo in Portal"

        ElseIf GST_drBest > 0 Then
            GST_drTG = UCase(Trim(CStr(aT(GST_drBest, 1))))
            GST_drTP = CStr(aT(GST_drBest, 2))
            GST_drTD = GST_NormDate(aT(GST_drBest, 3))
            GST_drTTV = GST_ToD(aT(GST_drBest, 5))
            GST_drTIG = GST_ToD(aT(GST_drBest, 6))
            GST_drTCG = GST_ToD(aT(GST_drBest, 7))
            GST_drTSG = GST_ToD(aT(GST_drBest, 8))
            GST_drTCS = GST_ToD(aT(GST_drBest, 9))

            aR(nR, 5) = GST_drTP
            aR(nR, 7) = GST_FmtDate(aT(GST_drBest, 3))
            aR(nR, 9) = CStr(aT(GST_drBest, 4))
            aR(nR, 11) = GST_drTTV
            aR(nR, 14) = GST_drTIG
            aR(nR, 17) = GST_drTCG
            aR(nR, 20) = GST_drTSG
            aR(nR, 23) = GST_drTCS
            aR(nR, 12) = GST_Rnd2(GST_drPTV - GST_drTTV)
            aR(nR, 15) = GST_Rnd2(GST_drPIG - GST_drTIG)
            aR(nR, 18) = GST_Rnd2(GST_drPCG - GST_drTCG)
            aR(nR, 21) = GST_Rnd2(GST_drPSG - GST_drTSG)
            aR(nR, 24) = GST_Rnd2(GST_drPCS - GST_drTCS)

            GST_drDet = ""
            GST_drAmt = False
            GST_drOth = False
            If GST_drPG <> GST_drTG Then GST_drDet = GST_drDet & "GSTIN | ": GST_drOth = True
            If Not GST_FuzzyParty(GST_drPP, GST_drTP) Then GST_drDet = GST_drDet & "Party Name | ": GST_drOth = True
            If GST_drPD <> GST_drTD Then GST_drDet = GST_drDet & "Inv Date | ": GST_drOth = True
            If GST_AmtDiff(GST_drPTV, GST_drTTV) Then GST_drDet = GST_drDet & "Taxable Val | ": GST_drAmt = True
            If GST_AmtDiff(GST_drPIG, GST_drTIG) Then GST_drDet = GST_drDet & "IGST | ": GST_drAmt = True
            If GST_AmtDiff(GST_drPCG, GST_drTCG) Then GST_drDet = GST_drDet & "CGST | ": GST_drAmt = True
            If GST_AmtDiff(GST_drPSG, GST_drTSG) Then GST_drDet = GST_drDet & "SGST | ": GST_drAmt = True
            If GST_AmtDiff(GST_drPCS, GST_drTCS) Then GST_drDet = GST_drDet & "CESS | ": GST_drAmt = True
            If Len(GST_drDet) > 3 Then GST_drDet = Left(GST_drDet, Len(GST_drDet) - 3)

            If GST_drDet = "" Then
                aR(nR, 1) = "Matched": aR(nR, 2) = ""
                If GST_drPF(GST_drRow) Then aR(nR, 2) = "Inv No format differs"
            ElseIf GST_drAmt And GST_drOth Then
                aR(nR, 1) = "Mismatch-Amt+Other": aR(nR, 2) = GST_drDet
            ElseIf GST_drAmt Then
                aR(nR, 1) = "Mismatch-Amount": aR(nR, 2) = GST_drDet
            Else
                aR(nR, 1) = "Mismatch-Other": aR(nR, 2) = GST_drDet
            End If

        Else
            aR(nR, 1) = "Missing in Books": aR(nR, 2) = ""
            aR(nR, 11) = 0: aR(nR, 14) = 0: aR(nR, 17) = 0
            aR(nR, 20) = 0: aR(nR, 23) = 0
            aR(nR, 12) = GST_Rnd2(GST_drPTV)
            aR(nR, 15) = GST_Rnd2(GST_drPIG)
            aR(nR, 18) = GST_Rnd2(GST_drPCG)
            aR(nR, 21) = GST_Rnd2(GST_drPSG)
            aR(nR, 24) = GST_Rnd2(GST_drPCS)
        End If
    Next GST_drRow

    '--------------------------------------------------------------
    ' STEP 5: Books rows not matched -> Missing in Portal
    '         Books duplicate rows   -> Duplicate-Books
    '--------------------------------------------------------------
    For GST_drRow = 2 To nT
        If Not GST_drUsed(GST_drRow) Then
            nR = nR + 1
            If nR > GST_drBuf Then Exit For
            If GST_drTDup(GST_drRow) Then
                aR(nR, 1) = "Duplicate-Books": aR(nR, 2) = "Dup InvNo in Books"
                aR(nR, 11) = GST_ToD(aT(GST_drRow, 5)): aR(nR, 14) = GST_ToD(aT(GST_drRow, 6))
                aR(nR, 17) = GST_ToD(aT(GST_drRow, 7)): aR(nR, 20) = GST_ToD(aT(GST_drRow, 8))
                aR(nR, 23) = GST_ToD(aT(GST_drRow, 9))
            Else
                aR(nR, 1) = "Missing in Portal": aR(nR, 2) = ""
                aR(nR, 10) = 0:                              aR(nR, 11) = GST_ToD(aT(GST_drRow, 5))
                aR(nR, 12) = GST_Rnd2(0 - GST_ToD(aT(GST_drRow, 5)))
                aR(nR, 13) = 0:                              aR(nR, 14) = GST_ToD(aT(GST_drRow, 6))
                aR(nR, 15) = GST_Rnd2(0 - GST_ToD(aT(GST_drRow, 6)))
                aR(nR, 16) = 0:                              aR(nR, 17) = GST_ToD(aT(GST_drRow, 7))
                aR(nR, 18) = GST_Rnd2(0 - GST_ToD(aT(GST_drRow, 7)))
                aR(nR, 19) = 0:                              aR(nR, 20) = GST_ToD(aT(GST_drRow, 8))
                aR(nR, 21) = GST_Rnd2(0 - GST_ToD(aT(GST_drRow, 8)))
                aR(nR, 22) = 0:                              aR(nR, 23) = GST_ToD(aT(GST_drRow, 9))
                aR(nR, 24) = GST_Rnd2(0 - GST_ToD(aT(GST_drRow, 9)))
            End If
            aR(nR, 3) = UCase(Trim(CStr(aT(GST_drRow, 1))))
            aR(nR, 5) = CStr(aT(GST_drRow, 2))
            aR(nR, 7) = GST_FmtDate(aT(GST_drRow, 3))
            aR(nR, 9) = CStr(aT(GST_drRow, 4))
        End If
    Next GST_drRow

    Call GST_SortOut(aR, nR)
End Sub

'==================================================================
' NEW v19: Extract pure numeric portion from raw invoice string
' Examples:
'   "ABC/12"  -> "12"     "AB/13"  -> "13"
'   "INV-007" -> "7"      "12"     -> "" (already numeric main key)
'   "ABCDEF"  -> ""       "RP/19"  -> "19"
'==================================================================
Function GST_ExtractPureNumeric(GST_epRaw As String) As String
    Dim GST_epNorm    As String
    Dim GST_epNoSep   As String
    Dim GST_epI       As Long
    Dim GST_epCh      As String
    Dim GST_epCurNum  As String
    Dim GST_epLastNum As String

    GST_epNorm = UCase(Trim(GST_epRaw))
    If Len(GST_epNorm) = 0 Then GST_ExtractPureNumeric = "": Exit Function

    GST_epNoSep = Replace(Replace(Replace(Replace(Replace( _
                  GST_epNorm, " ", ""), "/", ""), "-", ""), ".", ""), "_", "")
    If GST_AllDigits(GST_epNoSep) Then
        GST_ExtractPureNumeric = ""
        Exit Function
    End If

    GST_epLastNum = ""
    GST_epCurNum = ""

    For GST_epI = 1 To Len(GST_epNorm)
        GST_epCh = Mid(GST_epNorm, GST_epI, 1)
        If GST_epCh >= "0" And GST_epCh <= "9" Then
            GST_epCurNum = GST_epCurNum & GST_epCh
        Else
            If Len(GST_epCurNum) > 0 Then
                GST_epLastNum = GST_epCurNum
                GST_epCurNum = ""
            End If
        End If
    Next GST_epI
    If Len(GST_epCurNum) > 0 Then GST_epLastNum = GST_epCurNum

    If Len(GST_epLastNum) = 0 Then
        GST_ExtractPureNumeric = ""
    Else
        GST_ExtractPureNumeric = GST_StripLZ(GST_epLastNum)
    End If
End Function

'==================================================================
' HELPERS
'==================================================================
Sub GST_DicAdd(dic As Object, k As String, v As String)
    If dic.Exists(k) Then
        dic(k) = dic(k) & "," & v
    Else
        dic.Add k, v
    End If
End Sub

Function GST_LastSeg(GST_lsRaw As String) As String
    ' Trailing run of digits of an invoice no., leading zeros removed
    ' (26/4546 -> 4546 ; 4546 -> 4546 ; INV-0012 -> 12)
    Dim GST_lsS   As String
    Dim GST_lsI   As Long
    Dim GST_lsC   As String
    Dim GST_lsOut As String
    GST_lsS = UCase(Trim(GST_lsRaw))
    GST_lsOut = ""
    For GST_lsI = Len(GST_lsS) To 1 Step -1
        GST_lsC = Mid(GST_lsS, GST_lsI, 1)
        If GST_lsC >= "0" And GST_lsC <= "9" Then
            GST_lsOut = GST_lsC & GST_lsOut
        Else
            Exit For
        End If
    Next GST_lsI
    If Len(GST_lsOut) = 0 Then
        GST_LastSeg = ""
    Else
        GST_LastSeg = GST_StripLZ(GST_lsOut)
    End If
End Function

Sub GST_FindBestD(dicT As Object, srchKey As String, aT() As Variant, _
                  usedArr() As Boolean, pTV As Double, pDate As String, _
                  ByRef bestRow As Long, ByRef bestDiff As Double)
    ' Like GST_FindBest but the invoice date must also be identical
    Dim cands() As String
    Dim ci      As Long
    Dim tRow    As Long
    Dim diff    As Double

    If Not dicT.Exists(srchKey) Then Exit Sub
    cands = Split(CStr(dicT(srchKey)), ",")
    For ci = 0 To UBound(cands)
        tRow = CLng(Trim(cands(ci)))
        If Not usedArr(tRow) Then
            If GST_NormDate(aT(tRow, 3)) = pDate Then
                diff = Abs(GST_ToD(aT(tRow, 5)) - pTV)
                If diff < bestDiff Then
                    bestDiff = diff
                    bestRow = tRow
                End If
            End If
        End If
    Next ci
End Sub

Sub GST_FindBest(dicT As Object, srchKey As String, aT() As Variant, _
                 usedArr() As Boolean, pTV As Double, _
                 ByRef bestRow As Long, ByRef bestDiff As Double)
    Dim cands() As String
    Dim ci      As Long
    Dim tRow    As Long
    Dim diff    As Double

    If Not dicT.Exists(srchKey) Then Exit Sub
    cands = Split(CStr(dicT(srchKey)), ",")
    For ci = 0 To UBound(cands)
        tRow = CLng(Trim(cands(ci)))
        If Not usedArr(tRow) Then
            diff = Abs(GST_ToD(aT(tRow, 5)) - pTV)
            If diff < bestDiff Then
                bestDiff = diff
                bestRow = tRow
            End If
        End If
    Next ci
End Sub

'==================================================================
' WRITE RECO_RESULT SHEET
'==================================================================
Sub GST_WriteRows(wsR As Worksheet, aR() As Variant, nR As Long)
    Dim GST_wrHdr(23)  As String
    Dim GST_wrColN     As Long
    Dim GST_wrRecN     As Long
    Dim GST_wrRowN     As Long
    Dim GST_wrPrevGrp  As String
    Dim GST_wrClr      As Long
    Dim GST_wrCelVal   As Variant
    Dim GST_wrStat     As String
    Dim GST_wrBMap     As Object
    Dim GST_wrGrpSum(10 To 24) As Double
    Dim GST_wrGrpCnt   As Long
    Dim GST_wrGrpStart As Long
    Dim GST_wrC        As Long

    Set GST_wrBMap = CreateObject("Scripting.Dictionary")
    GST_wrBMap.Add "Matched", 0:              GST_wrBMap.Add "Mismatch-Amount", 1
    GST_wrBMap.Add "Mismatch-Amt+Other", 2:   GST_wrBMap.Add "Mismatch-Other", 3
    GST_wrBMap.Add "Missing in Portal", 4:    GST_wrBMap.Add "Missing in Books", 5
    GST_wrBMap.Add "Duplicate-Portal", 6:     GST_wrBMap.Add "Duplicate-Books", 7

    GST_wrHdr(0) = "Status":                GST_wrHdr(1) = "Mismatch Detail"
    GST_wrHdr(2) = "GSTIN":                 GST_wrHdr(3) = "Party Name (Portal)"
    GST_wrHdr(4) = "Party Name (Books)":    GST_wrHdr(5) = "Invoice Date (Portal)"
    GST_wrHdr(6) = "Invoice Date (Books)": GST_wrHdr(7) = "Invoice No (Portal)"
    GST_wrHdr(8) = "Invoice No (Books)":    GST_wrHdr(9) = "Taxable Val (Portal)"
    GST_wrHdr(10) = "Taxable Val (Books)":  GST_wrHdr(11) = "Diff Taxable Val"
    GST_wrHdr(12) = "IGST (Portal)":        GST_wrHdr(13) = "IGST (Books)"
    GST_wrHdr(14) = "Diff IGST":            GST_wrHdr(15) = "CGST (Portal)"
    GST_wrHdr(16) = "CGST (Books)":         GST_wrHdr(17) = "Diff CGST"
    GST_wrHdr(18) = "SGST (Portal)":        GST_wrHdr(19) = "SGST (Books)"
    GST_wrHdr(20) = "Diff SGST":            GST_wrHdr(21) = "CESS (Portal)"
    GST_wrHdr(22) = "CESS (Books)":         GST_wrHdr(23) = "Diff CESS"

    '--- Title ---
    With wsR.Range(wsR.Cells(1, 1), wsR.Cells(1, GST_NC))
        .Merge
        .Value = "GST RECONCILIATION REPORT v20  |  GSTIN-scoped Smart InvNo Matching  |  Full-Row Navigation"
        .Font.Bold = True: .Font.Size = 12: .Font.Color = vbWhite
        .Interior.Color = GST_CH: .HorizontalAlignment = xlLeft: .RowHeight = 28
    End With
    '--- Sub-title ---
    With wsR.Range(wsR.Cells(2, 1), wsR.Cells(2, GST_NC))
        .Merge
        .Value = "Run: " & Format(Now(), "DD-MMM-YYYY HH:MM:SS") & _
                 "  |  Party Threshold: 82%  |  Amount Tolerance: Rs " & Format(GST_TOL, "0.00") & _
                 "  |  Diff = Portal - Books" & _
                 "  |  Green=Zero  Red=Portal>Books  Blue=Portal<Books"
        .Font.Size = 9: .Font.Color = RGB(30, 50, 80)
        .Interior.Color = RGB(225, 235, 248)
        .HorizontalAlignment = xlLeft: .RowHeight = 15
    End With

    '--- Column headers ---
    For GST_wrColN = 0 To GST_NC - 1
        With wsR.Cells(3, GST_wrColN + 1)
            .Value = GST_wrHdr(GST_wrColN)
            .Font.Bold = True: .Font.Color = vbWhite: .Font.Size = 9
            If GST_wrColN = 11 Or GST_wrColN = 14 Or GST_wrColN = 17 _
               Or GST_wrColN = 20 Or GST_wrColN = 23 Then
                .Interior.Color = RGB(56, 104, 116)
            Else
                .Interior.Color = GST_CH
            End If
            .HorizontalAlignment = xlCenter: .WrapText = True
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Color = RGB(150, 170, 195)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(150, 170, 195)
        End With
    Next GST_wrColN
    wsR.Rows(3).RowHeight = 40

    wsR.Activate
    wsR.Range("A4").Select
    ActiveWindow.FreezePanes = True

    GST_wrRowN = 4
    GST_wrPrevGrp = ""
    GST_wrGrpCnt = 0
    GST_wrGrpStart = 0

    For GST_wrRecN = 1 To nR
        GST_wrStat = CStr(aR(GST_wrRecN, 1))

        '--- Group change ---
        If GST_wrStat <> GST_wrPrevGrp Then
            If GST_wrGrpCnt > 0 And GST_wrGrpStart > 0 Then
                Call GST_WriteGroupTotal(wsR, GST_wrRowN, GST_wrGrpSum, GST_wrGrpCnt, GST_wrPrevGrp)
                GST_wrRowN = GST_wrRowN + 1
            End If
            If GST_wrRowN > 4 Then GST_wrRowN = GST_wrRowN + 1

            With wsR.Range(wsR.Cells(GST_wrRowN, 1), wsR.Cells(GST_wrRowN, GST_NC))
                .Merge
                .Value = ">> " & UCase(GST_wrStat)
                .Font.Bold = True: .Font.Color = vbWhite: .Font.Size = 10
                .Interior.Color = GST_BnrClr(GST_wrStat)
                .HorizontalAlignment = xlLeft: .RowHeight = 20
            End With
            If GST_wrBMap.Exists(GST_wrStat) Then
                GST_BnrRow(GST_wrBMap(GST_wrStat)) = GST_wrRowN
            End If
            GST_wrRowN = GST_wrRowN + 1
            GST_wrPrevGrp = GST_wrStat
            GST_wrGrpCnt = 0
            GST_wrGrpStart = GST_wrRowN
            For GST_wrC = 10 To 24
                GST_wrGrpSum(GST_wrC) = 0
            Next GST_wrC
        End If

        GST_wrClr = GST_RowClr(GST_wrStat)

        '--- Write data row ---
        For GST_wrColN = 1 To GST_NC
            GST_wrCelVal = aR(GST_wrRecN, GST_wrColN)
            With wsR.Cells(GST_wrRowN, GST_wrColN)
                .Interior.Color = GST_wrClr
                .Font.Size = 9: .Font.Color = RGB(25, 35, 55)
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Color = RGB(195, 200, 215)
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Color = RGB(195, 200, 215)
                If GST_wrColN >= 10 Then
                    If IsNumeric(GST_wrCelVal) Then
                        .Value = CDbl(GST_wrCelVal)
                        .NumberFormat = "#,##0.00"
                        .HorizontalAlignment = xlRight
                        .Font.Bold = True
                    Else
                        .Value = "": .HorizontalAlignment = xlRight: .Font.Bold = False
                    End If
                Else
                    .Value = CStr(GST_wrCelVal): .HorizontalAlignment = xlLeft: .Font.Bold = False
                End If
            End With
        Next GST_wrColN

        '--- Accumulate group totals ---
        For GST_wrC = 10 To 24
            If IsNumeric(aR(GST_wrRecN, GST_wrC)) Then
                GST_wrGrpSum(GST_wrC) = GST_wrGrpSum(GST_wrC) + CDbl(aR(GST_wrRecN, GST_wrC))
            End If
        Next GST_wrC
        GST_wrGrpCnt = GST_wrGrpCnt + 1

        If InStr(1, GST_wrStat, "Mismatch") > 0 Then
            Call GST_HiAmt(wsR, GST_wrRowN, aR, GST_wrRecN)
        End If
        Call GST_HiDiff(wsR, GST_wrRowN, aR, GST_wrRecN)

        GST_wrRowN = GST_wrRowN + 1
    Next GST_wrRecN

    '--- Last group total ---
    If GST_wrGrpCnt > 0 And GST_wrGrpStart > 0 Then
        Call GST_WriteGroupTotal(wsR, GST_wrRowN, GST_wrGrpSum, GST_wrGrpCnt, GST_wrPrevGrp)
    End If
End Sub

'==================================================================
' GROUP TOTAL ROW
'==================================================================
Sub GST_WriteGroupTotal(wsR As Worksheet, GST_gtRN As Long, _
                        grpSum() As Double, grpCnt As Long, grpStat As String)
    Dim GST_gtC   As Long
    Dim GST_gtClr As Long
    Dim dv        As Double

    Select Case grpStat
        Case "Matched":            GST_gtClr = RGB(140, 195, 148)
        Case "Mismatch-Amount":    GST_gtClr = RGB(218, 175, 100)
        Case "Mismatch-Amt+Other": GST_gtClr = RGB(210, 163, 88)
        Case "Mismatch-Other":     GST_gtClr = RGB(196, 198, 122)
        Case "Missing in Portal":  GST_gtClr = RGB(210, 155, 148)
        Case "Missing in Books":   GST_gtClr = RGB(140, 182, 210)
        Case "Duplicate-Portal":   GST_gtClr = RGB(175, 158, 200)
        Case "Duplicate-Books":    GST_gtClr = RGB(162, 145, 188)
        Case Else:                 GST_gtClr = RGB(175, 178, 192)
    End Select

    With wsR.Range(wsR.Cells(GST_gtRN, 1), wsR.Cells(GST_gtRN, 9))
        .Merge: .Value = "TOTAL  (" & CStr(grpCnt) & " records)"
        .Font.Bold = True: .Font.Size = 9: .Font.Color = RGB(20, 30, 55)
        .Interior.Color = GST_gtClr: .HorizontalAlignment = xlRight: .RowHeight = 18
        .Borders(xlEdgeBottom).LineStyle = xlDouble
        .Borders(xlEdgeBottom).Color = RGB(80, 100, 140)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = RGB(120, 140, 170)
    End With

    For GST_gtC = 10 To 24
        dv = GST_Rnd2(grpSum(GST_gtC))
        With wsR.Cells(GST_gtRN, GST_gtC)
            .Value = dv: .NumberFormat = "#,##0.00"
            .Font.Bold = True: .Font.Size = 9
            .Interior.Color = GST_gtClr: .HorizontalAlignment = xlRight
            .Borders(xlEdgeBottom).LineStyle = xlDouble
            .Borders(xlEdgeBottom).Color = RGB(80, 100, 140)
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Color = RGB(120, 140, 170)
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Color = RGB(150, 160, 185)
            If GST_gtC = 12 Or GST_gtC = 15 Or GST_gtC = 18 _
               Or GST_gtC = 21 Or GST_gtC = 24 Then
                If Abs(dv) < 0.005 Then
                    .Font.Color = RGB(30, 110, 50)
                ElseIf dv > 0 Then
                    .Font.Color = RGB(150, 40, 30)
                Else
                    .Font.Color = RGB(30, 60, 155)
                End If
            Else
                .Font.Color = RGB(20, 30, 55)
            End If
        End With
    Next GST_gtC
End Sub

'==================================================================
' HIGHLIGHT MISMATCH AMOUNT PAIRS
'==================================================================
Sub GST_HiAmt(wsR As Worksheet, GST_haRN As Long, aR() As Variant, GST_haRC As Long)
    Dim GST_haPidx    As Long
    Dim GST_haPArr(4) As Long
    Dim GST_haTArr(4) As Long
    GST_haPArr(0) = 10: GST_haTArr(0) = 11
    GST_haPArr(1) = 13: GST_haTArr(1) = 14
    GST_haPArr(2) = 16: GST_haTArr(2) = 17
    GST_haPArr(3) = 19: GST_haTArr(3) = 20
    GST_haPArr(4) = 22: GST_haTArr(4) = 23
    For GST_haPidx = 0 To 4
        If GST_AmtDiff(GST_ToD(aR(GST_haRC, GST_haPArr(GST_haPidx))), _
                       GST_ToD(aR(GST_haRC, GST_haTArr(GST_haPidx)))) Then
            wsR.Cells(GST_haRN, GST_haPArr(GST_haPidx)).Interior.Color = RGB(255, 200, 140)
            wsR.Cells(GST_haRN, GST_haPArr(GST_haPidx)).Font.Bold = True
            wsR.Cells(GST_haRN, GST_haPArr(GST_haPidx)).Font.Color = RGB(130, 45, 10)
            wsR.Cells(GST_haRN, GST_haTArr(GST_haPidx)).Interior.Color = RGB(255, 200, 140)
            wsR.Cells(GST_haRN, GST_haTArr(GST_haPidx)).Font.Bold = True
            wsR.Cells(GST_haRN, GST_haTArr(GST_haPidx)).Font.Color = RGB(130, 45, 10)
        End If
    Next GST_haPidx
End Sub

'==================================================================
' HIGHLIGHT DIFF COLUMNS
'==================================================================
Sub GST_HiDiff(wsR As Worksheet, GST_hdRN As Long, aR() As Variant, GST_hdRC As Long)
    Dim GST_hdIdx         As Long
    Dim GST_hdVal         As Double
    Dim GST_hdDiffCols(4) As Long
    GST_hdDiffCols(0) = 12: GST_hdDiffCols(1) = 15
    GST_hdDiffCols(2) = 18: GST_hdDiffCols(3) = 21
    GST_hdDiffCols(4) = 24
    For GST_hdIdx = 0 To 4
        GST_hdVal = GST_Rnd2(GST_ToD(aR(GST_hdRC, GST_hdDiffCols(GST_hdIdx))))
        With wsR.Cells(GST_hdRN, GST_hdDiffCols(GST_hdIdx))
            .Font.Bold = True
            If Abs(GST_hdVal) < 0.005 Then
                .Interior.Color = RGB(200, 235, 200): .Font.Color = RGB(20, 100, 40)
            ElseIf GST_hdVal > 0 Then
                .Interior.Color = RGB(252, 215, 195): .Font.Color = RGB(140, 35, 20)
            Else
                .Interior.Color = RGB(200, 220, 250): .Font.Color = RGB(25, 55, 150)
            End If
        End With
    Next GST_hdIdx
End Sub

'==================================================================
' SUMMARY SHEET  v19
' - Category cell col 1 = arrow icon + hyperlink to Reco_Result
' - Cols 1-17 = hyperlinks on entire row
' - Col 18 = vivid navy button with orange border
'==================================================================
Sub GST_WriteSumSheet(wsS As Worksheet, aR() As Variant, nR As Long)
    Dim GST_ssRec    As Long
    Dim GST_ssStat   As String
    Dim GST_ssGrp    As Long
    Dim GST_ssCnt(0 To 7)   As Long
    Dim GST_ssPTax(0 To 7)  As Double
    Dim GST_ssTTax(0 To 7)  As Double
    Dim GST_ssPIGS(0 To 7)  As Double
    Dim GST_ssTIGS(0 To 7)  As Double
    Dim GST_ssPCGS(0 To 7)  As Double
    Dim GST_ssTCGS(0 To 7)  As Double
    Dim GST_ssPSGS(0 To 7)  As Double
    Dim GST_ssTSGS(0 To 7)  As Double
    Dim GST_ssPCES(0 To 7)  As Double
    Dim GST_ssTCES(0 To 7)  As Double
    Dim GST_ssRow      As Long
    Dim GST_ssCol      As Long
    Dim GST_ssIdx      As Long
    Dim GST_ssClr      As Long
    Dim GST_ssDiff     As Double
    Dim GST_ssTotCnt   As Long
    Dim GST_ssHdr(17)     As String
    Dim GST_ssGrpNm(7)    As String
    Dim GST_ssGrpClr(7)   As Long
    Dim GST_ssLgNm(5)     As String
    Dim GST_ssLgClr(5)    As Long
    Dim GST_ssDicM     As Object
    Dim GST_ssPV       As Double
    Dim GST_ssTV       As Double
    Dim GST_ssTotP(4)  As Double
    Dim GST_ssTotT(4)  As Double
    Dim GST_ssGTDiff   As Double
    Dim GST_ssTgtAddr  As String
    Dim GST_ssCellVal  As String

    Set GST_ssDicM = CreateObject("Scripting.Dictionary")
    GST_ssDicM.Add "Matched", 0:              GST_ssDicM.Add "Mismatch-Amount", 1
    GST_ssDicM.Add "Mismatch-Amt+Other", 2:   GST_ssDicM.Add "Mismatch-Other", 3
    GST_ssDicM.Add "Missing in Portal", 4:    GST_ssDicM.Add "Missing in Books", 5
    GST_ssDicM.Add "Duplicate-Portal", 6:     GST_ssDicM.Add "Duplicate-Books", 7

    For GST_ssRec = 1 To nR
        GST_ssStat = CStr(aR(GST_ssRec, 1))
        If Not GST_ssDicM.Exists(GST_ssStat) Then GoTo GST_ssNextRec
        GST_ssGrp = GST_ssDicM(GST_ssStat)
        GST_ssCnt(GST_ssGrp) = GST_ssCnt(GST_ssGrp) + 1
        GST_ssPTax(GST_ssGrp) = GST_ssPTax(GST_ssGrp) + GST_ToD(aR(GST_ssRec, 10))
        GST_ssTTax(GST_ssGrp) = GST_ssTTax(GST_ssGrp) + GST_ToD(aR(GST_ssRec, 11))
        GST_ssPIGS(GST_ssGrp) = GST_ssPIGS(GST_ssGrp) + GST_ToD(aR(GST_ssRec, 13))
        GST_ssTIGS(GST_ssGrp) = GST_ssTIGS(GST_ssGrp) + GST_ToD(aR(GST_ssRec, 14))
        GST_ssPCGS(GST_ssGrp) = GST_ssPCGS(GST_ssGrp) + GST_ToD(aR(GST_ssRec, 16))
        GST_ssTCGS(GST_ssGrp) = GST_ssTCGS(GST_ssGrp) + GST_ToD(aR(GST_ssRec, 17))
        GST_ssPSGS(GST_ssGrp) = GST_ssPSGS(GST_ssGrp) + GST_ToD(aR(GST_ssRec, 19))
        GST_ssTSGS(GST_ssGrp) = GST_ssTSGS(GST_ssGrp) + GST_ToD(aR(GST_ssRec, 20))
        GST_ssPCES(GST_ssGrp) = GST_ssPCES(GST_ssGrp) + GST_ToD(aR(GST_ssRec, 22))
        GST_ssTCES(GST_ssGrp) = GST_ssTCES(GST_ssGrp) + GST_ToD(aR(GST_ssRec, 23))
GST_ssNextRec:
    Next GST_ssRec

    For GST_ssIdx = 0 To 7
        GST_ssTotCnt = GST_ssTotCnt + GST_ssCnt(GST_ssIdx)
        GST_ssTotP(0) = GST_ssTotP(0) + GST_ssPTax(GST_ssIdx)
        GST_ssTotT(0) = GST_ssTotT(0) + GST_ssTTax(GST_ssIdx)
        GST_ssTotP(1) = GST_ssTotP(1) + GST_ssPIGS(GST_ssIdx)
        GST_ssTotT(1) = GST_ssTotT(1) + GST_ssTIGS(GST_ssIdx)
        GST_ssTotP(2) = GST_ssTotP(2) + GST_ssPCGS(GST_ssIdx)
        GST_ssTotT(2) = GST_ssTotT(2) + GST_ssTCGS(GST_ssIdx)
        GST_ssTotP(3) = GST_ssTotP(3) + GST_ssPSGS(GST_ssIdx)
        GST_ssTotT(3) = GST_ssTotT(3) + GST_ssTSGS(GST_ssIdx)
        GST_ssTotP(4) = GST_ssTotP(4) + GST_ssPCES(GST_ssIdx)
        GST_ssTotT(4) = GST_ssTotT(4) + GST_ssTCES(GST_ssIdx)
    Next GST_ssIdx

    '--- Title ---
    With wsS.Range(wsS.Cells(1, 1), wsS.Cells(1, GST_SNC))
        .Merge: .Value = "GST RECONCILIATION  -  SUMMARY DASHBOARD v20"
        .Font.Bold = True: .Font.Size = 14: .Font.Color = vbWhite
        .Interior.Color = GST_CH: .HorizontalAlignment = xlCenter: .RowHeight = 34
    End With
    With wsS.Range(wsS.Cells(2, 1), wsS.Cells(2, GST_SNC))
        .Merge
        .Value = "Run: " & Format(Now(), "DD-MMM-YYYY HH:MM:SS") & _
                 "  |  Total: " & CStr(GST_ssTotCnt) & " records" & _
                 "  |  Click anywhere on a category row to navigate to Reco_Result"
        .Font.Size = 9: .Font.Color = RGB(25, 40, 75)
        .Interior.Color = RGB(225, 235, 248): .HorizontalAlignment = xlLeft: .RowHeight = 16
    End With

    '--- Column headers ---
    GST_ssHdr(0) = "Category":         GST_ssHdr(1) = "Count"
    GST_ssHdr(2) = "Portal Taxable":   GST_ssHdr(3) = "Books Taxable"
    GST_ssHdr(4) = "Diff Taxable":     GST_ssHdr(5) = "Portal IGST"
    GST_ssHdr(6) = "Books IGST":       GST_ssHdr(7) = "Diff IGST"
    GST_ssHdr(8) = "Portal CGST":      GST_ssHdr(9) = "Books CGST"
    GST_ssHdr(10) = "Diff CGST":       GST_ssHdr(11) = "Portal SGST"
    GST_ssHdr(12) = "Books SGST":      GST_ssHdr(13) = "Diff SGST"
    GST_ssHdr(14) = "Portal CESS":     GST_ssHdr(15) = "Books CESS"
    GST_ssHdr(16) = "Diff CESS":       GST_ssHdr(17) = "NAVIGATE"

    For GST_ssCol = 0 To GST_SNC - 1
        With wsS.Cells(3, GST_ssCol + 1)
            .Font.Bold = True: .Font.Size = 9
            If GST_ssCol = 17 Then
                .Value = "NAVIGATE  >>": .Font.Color = vbWhite
                .Interior.Color = RGB(20, 80, 165)
            Else
                .Value = GST_ssHdr(GST_ssCol)
                .Font.Color = vbWhite: .Interior.Color = GST_CS
            End If
            .HorizontalAlignment = xlCenter: .WrapText = True
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Color = RGB(130, 155, 190)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(130, 155, 190)
        End With
    Next GST_ssCol
    wsS.Rows(3).RowHeight = 36

    GST_ssGrpNm(0) = "Matched":            GST_ssGrpNm(1) = "Mismatch-Amount"
    GST_ssGrpNm(2) = "Mismatch-Amt+Other": GST_ssGrpNm(3) = "Mismatch-Other"
    GST_ssGrpNm(4) = "Missing in Portal":  GST_ssGrpNm(5) = "Missing in Books"
    GST_ssGrpNm(6) = "Duplicate-Portal":   GST_ssGrpNm(7) = "Duplicate-Books"
    GST_ssGrpClr(0) = GST_CM: GST_ssGrpClr(1) = GST_CA: GST_ssGrpClr(2) = GST_CA
    GST_ssGrpClr(3) = GST_CO: GST_ssGrpClr(4) = GST_CP: GST_ssGrpClr(5) = GST_CT
    GST_ssGrpClr(6) = GST_CD: GST_ssGrpClr(7) = GST_CD

    GST_ssRow = 4
    For GST_ssIdx = 0 To 7
        GST_ssClr = GST_ssGrpClr(GST_ssIdx)

        '--- Write cell values first (needed for TextToDisplay in hyperlinks) ---
        wsS.Cells(GST_ssRow, 1).Value = "> " & GST_ssGrpNm(GST_ssIdx)
        wsS.Cells(GST_ssRow, 1).Font.Bold = True

        wsS.Cells(GST_ssRow, 2).Value = GST_ssCnt(GST_ssIdx)
        wsS.Cells(GST_ssRow, 2).HorizontalAlignment = xlCenter
        wsS.Cells(GST_ssRow, 2).Font.Bold = True

        GST_ssPV = GST_ssPTax(GST_ssIdx): GST_ssTV = GST_ssTTax(GST_ssIdx)
        GST_ssDiff = GST_Rnd2(GST_ssPV - GST_ssTV)
        wsS.Cells(GST_ssRow, 3).Value = GST_ssPV:    wsS.Cells(GST_ssRow, 3).Font.Bold = True
        wsS.Cells(GST_ssRow, 4).Value = GST_ssTV:    wsS.Cells(GST_ssRow, 4).Font.Bold = True
        wsS.Cells(GST_ssRow, 5).Value = GST_ssDiff:  wsS.Cells(GST_ssRow, 5).Font.Bold = True
        Call GST_SsDiffClr(wsS, GST_ssRow, 5, GST_ssDiff)

        GST_ssPV = GST_ssPIGS(GST_ssIdx): GST_ssTV = GST_ssTIGS(GST_ssIdx)
        GST_ssDiff = GST_Rnd2(GST_ssPV - GST_ssTV)
        wsS.Cells(GST_ssRow, 6).Value = GST_ssPV:    wsS.Cells(GST_ssRow, 6).Font.Bold = True
        wsS.Cells(GST_ssRow, 7).Value = GST_ssTV:    wsS.Cells(GST_ssRow, 7).Font.Bold = True
        wsS.Cells(GST_ssRow, 8).Value = GST_ssDiff:  wsS.Cells(GST_ssRow, 8).Font.Bold = True
        Call GST_SsDiffClr(wsS, GST_ssRow, 8, GST_ssDiff)

        GST_ssPV = GST_ssPCGS(GST_ssIdx): GST_ssTV = GST_ssTCGS(GST_ssIdx)
        GST_ssDiff = GST_Rnd2(GST_ssPV - GST_ssTV)
        wsS.Cells(GST_ssRow, 9).Value = GST_ssPV:    wsS.Cells(GST_ssRow, 9).Font.Bold = True
        wsS.Cells(GST_ssRow, 10).Value = GST_ssTV:   wsS.Cells(GST_ssRow, 10).Font.Bold = True
        wsS.Cells(GST_ssRow, 11).Value = GST_ssDiff: wsS.Cells(GST_ssRow, 11).Font.Bold = True
        Call GST_SsDiffClr(wsS, GST_ssRow, 11, GST_ssDiff)

        GST_ssPV = GST_ssPSGS(GST_ssIdx): GST_ssTV = GST_ssTSGS(GST_ssIdx)
        GST_ssDiff = GST_Rnd2(GST_ssPV - GST_ssTV)
        wsS.Cells(GST_ssRow, 12).Value = GST_ssPV:   wsS.Cells(GST_ssRow, 12).Font.Bold = True
        wsS.Cells(GST_ssRow, 13).Value = GST_ssTV:   wsS.Cells(GST_ssRow, 13).Font.Bold = True
        wsS.Cells(GST_ssRow, 14).Value = GST_ssDiff: wsS.Cells(GST_ssRow, 14).Font.Bold = True
        Call GST_SsDiffClr(wsS, GST_ssRow, 14, GST_ssDiff)

        GST_ssPV = GST_ssPCES(GST_ssIdx): GST_ssTV = GST_ssTCES(GST_ssIdx)
        GST_ssDiff = GST_Rnd2(GST_ssPV - GST_ssTV)
        wsS.Cells(GST_ssRow, 15).Value = GST_ssPV:   wsS.Cells(GST_ssRow, 15).Font.Bold = True
        wsS.Cells(GST_ssRow, 16).Value = GST_ssTV:   wsS.Cells(GST_ssRow, 16).Font.Bold = True
        wsS.Cells(GST_ssRow, 17).Value = GST_ssDiff: wsS.Cells(GST_ssRow, 17).Font.Bold = True
        Call GST_SsDiffClr(wsS, GST_ssRow, 17, GST_ssDiff)

        For GST_ssCol = 3 To 17
            wsS.Cells(GST_ssRow, GST_ssCol).NumberFormat = "#,##0.00"
            wsS.Cells(GST_ssRow, GST_ssCol).HorizontalAlignment = xlRight
            wsS.Cells(GST_ssRow, GST_ssCol).Font.Size = 9
        Next GST_ssCol

        '--- Apply row background and borders to all 18 cols ---
        For GST_ssCol = 1 To GST_SNC
            With wsS.Cells(GST_ssRow, GST_ssCol)
                .Interior.Color = GST_ssClr
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Color = RGB(170, 180, 205)
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Color = RGB(170, 180, 205)
                .Font.Size = 9
            End With
        Next GST_ssCol

        '================================================================
        ' FULL-ROW NAVIGATION - hyperlinks on cols 1-17, button on col 18
        '================================================================
        If GST_BnrRow(GST_ssIdx) > 0 Then
            GST_ssTgtAddr = "Reco_Result!A" & CStr(GST_BnrRow(GST_ssIdx))

            '--- Cols 1-17: individual cell hyperlinks ---
            For GST_ssCol = 1 To 17
                GST_ssCellVal = CStr(wsS.Cells(GST_ssRow, GST_ssCol).Value)
                If Len(GST_ssCellVal) = 0 Then GST_ssCellVal = " "
                wsS.Hyperlinks.Add _
                    Anchor:=wsS.Cells(GST_ssRow, GST_ssCol), _
                    Address:="", _
                    SubAddress:=GST_ssTgtAddr, _
                    TextToDisplay:=GST_ssCellVal
                With wsS.Cells(GST_ssRow, GST_ssCol)
                    .Font.Color = RGB(20, 30, 60)
                    .Font.Underline = xlUnderlineStyleNone
                    .Font.Bold = True
                    .Font.Size = 9
                    .Interior.Color = GST_ssClr
                End With
            Next GST_ssCol

            '--- Col 18: vivid navy button ---
            wsS.Hyperlinks.Add _
                Anchor:=wsS.Cells(GST_ssRow, 18), _
                Address:="", _
                SubAddress:=GST_ssTgtAddr, _
                TextToDisplay:=">> " & GST_ssGrpNm(GST_ssIdx) & " >>"
            With wsS.Cells(GST_ssRow, 18)
                .Interior.Color = RGB(20, 80, 165)
                .Font.Bold = True: .Font.Size = 10: .Font.Color = vbWhite
                .Font.Underline = xlUnderlineStyleNone
                .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlMedium
                .Borders(xlEdgeLeft).Color = RGB(220, 120, 20)
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlMedium
                .Borders(xlEdgeRight).Color = RGB(220, 120, 20)
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlMedium
                .Borders(xlEdgeTop).Color = RGB(220, 120, 20)
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlMedium
                .Borders(xlEdgeBottom).Color = RGB(220, 120, 20)
            End With
        Else
            '--- No records: muted button, no hyperlink ---
            With wsS.Cells(GST_ssRow, 18)
                .Value = "(No records)"
                .Interior.Color = RGB(170, 175, 185)
                .Font.Bold = True: .Font.Size = 9
                .Font.Color = RGB(230, 232, 238)
                .HorizontalAlignment = xlCenter
            End With
        End If

        wsS.Rows(GST_ssRow).RowHeight = 22
        GST_ssRow = GST_ssRow + 1
    Next GST_ssIdx

    '--- Grand Total row ---
    With wsS.Range(wsS.Cells(GST_ssRow, 1), wsS.Cells(GST_ssRow, 2))
        .Merge: .Value = "GRAND TOTAL"
        .Font.Bold = True: .Font.Size = 10: .Font.Color = vbWhite
        .Interior.Color = GST_CH: .HorizontalAlignment = xlCenter
    End With
    wsS.Cells(GST_ssRow, 2).Interior.Color = GST_CH: wsS.Cells(GST_ssRow, 2).Value = ""
    For GST_ssCol = 3 To 18
        With wsS.Cells(GST_ssRow, GST_ssCol)
            .Interior.Color = GST_CH: .Font.Color = vbWhite
            .Font.Bold = True: .Font.Size = 9
        End With
    Next GST_ssCol
    With wsS.Cells(GST_ssRow, 1)
        .Value = "GRAND TOTAL": .Font.Bold = True
        .Font.Color = vbWhite: .Interior.Color = GST_CH
    End With
    wsS.Cells(GST_ssRow, 2).Value = GST_ssTotCnt
    wsS.Cells(GST_ssRow, 2).Font.Bold = True
    wsS.Cells(GST_ssRow, 2).Font.Color = vbWhite
    wsS.Cells(GST_ssRow, 2).Interior.Color = GST_CH
    wsS.Cells(GST_ssRow, 2).HorizontalAlignment = xlCenter
    For GST_ssCol = 3 To 17
        wsS.Cells(GST_ssRow, GST_ssCol).NumberFormat = "#,##0.00"
        wsS.Cells(GST_ssRow, GST_ssCol).HorizontalAlignment = xlRight
    Next GST_ssCol

    GST_ssPV = GST_ssTotP(0): GST_ssTV = GST_ssTotT(0)
    GST_ssGTDiff = GST_Rnd2(GST_ssPV - GST_ssTV)
    wsS.Cells(GST_ssRow, 3).Value = GST_ssPV: wsS.Cells(GST_ssRow, 4).Value = GST_ssTV
    wsS.Cells(GST_ssRow, 5).Value = GST_ssGTDiff
    If Abs(GST_ssGTDiff) > 0.005 Then wsS.Cells(GST_ssRow, 5).Font.Color = RGB(255, 160, 130) _
                                  Else wsS.Cells(GST_ssRow, 5).Font.Color = RGB(160, 240, 170)

    GST_ssPV = GST_ssTotP(1): GST_ssTV = GST_ssTotT(1)
    GST_ssGTDiff = GST_Rnd2(GST_ssPV - GST_ssTV)
    wsS.Cells(GST_ssRow, 6).Value = GST_ssPV: wsS.Cells(GST_ssRow, 7).Value = GST_ssTV
    wsS.Cells(GST_ssRow, 8).Value = GST_ssGTDiff
    If Abs(GST_ssGTDiff) > 0.005 Then wsS.Cells(GST_ssRow, 8).Font.Color = RGB(255, 160, 130) _
                                  Else wsS.Cells(GST_ssRow, 8).Font.Color = RGB(160, 240, 170)

    GST_ssPV = GST_ssTotP(2): GST_ssTV = GST_ssTotT(2)
    GST_ssGTDiff = GST_Rnd2(GST_ssPV - GST_ssTV)
    wsS.Cells(GST_ssRow, 9).Value = GST_ssPV: wsS.Cells(GST_ssRow, 10).Value = GST_ssTV
    wsS.Cells(GST_ssRow, 11).Value = GST_ssGTDiff
    If Abs(GST_ssGTDiff) > 0.005 Then wsS.Cells(GST_ssRow, 11).Font.Color = RGB(255, 160, 130) _
                                  Else wsS.Cells(GST_ssRow, 11).Font.Color = RGB(160, 240, 170)

    GST_ssPV = GST_ssTotP(3): GST_ssTV = GST_ssTotT(3)
    GST_ssGTDiff = GST_Rnd2(GST_ssPV - GST_ssTV)
    wsS.Cells(GST_ssRow, 12).Value = GST_ssPV: wsS.Cells(GST_ssRow, 13).Value = GST_ssTV
    wsS.Cells(GST_ssRow, 14).Value = GST_ssGTDiff
    If Abs(GST_ssGTDiff) > 0.005 Then wsS.Cells(GST_ssRow, 14).Font.Color = RGB(255, 160, 130) _
                                  Else wsS.Cells(GST_ssRow, 14).Font.Color = RGB(160, 240, 170)

    GST_ssPV = GST_ssTotP(4): GST_ssTV = GST_ssTotT(4)
    GST_ssGTDiff = GST_Rnd2(GST_ssPV - GST_ssTV)
    wsS.Cells(GST_ssRow, 15).Value = GST_ssPV: wsS.Cells(GST_ssRow, 16).Value = GST_ssTV
    wsS.Cells(GST_ssRow, 17).Value = GST_ssGTDiff
    If Abs(GST_ssGTDiff) > 0.005 Then wsS.Cells(GST_ssRow, 17).Font.Color = RGB(255, 160, 130) _
                                  Else wsS.Cells(GST_ssRow, 17).Font.Color = RGB(160, 240, 170)

    wsS.Rows(GST_ssRow).RowHeight = 22
    GST_ssRow = GST_ssRow + 2

    '--- Colour legend ---
    GST_ssLgNm(0) = "  > Matched":           GST_ssLgNm(1) = "  > Mismatch (Amount)"
    GST_ssLgNm(2) = "  > Mismatch (Other)":  GST_ssLgNm(3) = "  > Missing in Portal"
    GST_ssLgNm(4) = "  > Missing in Books":  GST_ssLgNm(5) = "  > Duplicate Invoice"
    GST_ssLgClr(0) = GST_CM: GST_ssLgClr(1) = GST_CA: GST_ssLgClr(2) = GST_CO
    GST_ssLgClr(3) = GST_CP: GST_ssLgClr(4) = GST_CT: GST_ssLgClr(5) = GST_CD

    With wsS.Cells(GST_ssRow, 1)
        .Value = "COLOUR LEGEND": .Font.Bold = True
        .Font.Size = 9: .Font.Color = RGB(25, 40, 75)
    End With
    wsS.Rows(GST_ssRow).RowHeight = 16
    GST_ssRow = GST_ssRow + 1

    For GST_ssIdx = 0 To 5
        With wsS.Range(wsS.Cells(GST_ssRow, 1), wsS.Cells(GST_ssRow, 5))
            .Merge: .Value = GST_ssLgNm(GST_ssIdx)
            .Interior.Color = GST_ssLgClr(GST_ssIdx)
            .Font.Bold = True: .Font.Size = 9
            .Font.Color = RGB(20, 30, 55): .RowHeight = 15
        End With
        GST_ssRow = GST_ssRow + 1
    Next GST_ssIdx

    GST_ssRow = GST_ssRow + 1
    With wsS.Range(wsS.Cells(GST_ssRow, 1), wsS.Cells(GST_ssRow, GST_SNC))
        .Merge: .Value = "Created by Chinmay Bhat  |  GST Reconciliation Tool v20"
        .Font.Size = 7: .Font.Color = RGB(140, 150, 170)
        .HorizontalAlignment = xlRight: .RowHeight = 13
    End With

    wsS.Columns("A").ColumnWidth = 28: wsS.Columns("B").ColumnWidth = 8
    For GST_ssCol = 3 To 17
        wsS.Columns(GST_ssCol).ColumnWidth = 14
    Next GST_ssCol
    wsS.Columns(18).ColumnWidth = 30
    wsS.Rows(3).RowHeight = 36
    wsS.Activate: wsS.Range("A1").Select
End Sub

'==================================================================
' HELPER: diff colour for summary cells
'==================================================================
Sub GST_SsDiffClr(wsS As Worksheet, r As Long, c As Long, dv As Double)
    If Abs(dv) > 0.005 Then
        wsS.Cells(r, c).Font.Color = RGB(145, 35, 20)
    Else
        wsS.Cells(r, c).Font.Color = RGB(25, 100, 40)
    End If
End Sub

'==================================================================
' SORT
'==================================================================
Sub GST_SortOut(aR() As Variant, nR As Long)
    ' Stable O(n) counting sort on the 8 status groups
    Dim GST_soObj      As Object
    Dim GST_soPri()    As Long
    Dim GST_soCnt(0 To 100) As Long
    Dim GST_soPos(0 To 100) As Long
    Dim GST_soTmp()    As Variant
    Dim GST_soI        As Long
    Dim GST_soC        As Long
    Dim GST_soP        As Long
    Dim GST_soDst      As Long
    Dim GST_soRun      As Long

    If nR < 2 Then Exit Sub

    Set GST_soObj = CreateObject("Scripting.Dictionary")
    GST_soObj.Add "Matched", 1:              GST_soObj.Add "Mismatch-Amount", 2
    GST_soObj.Add "Mismatch-Amt+Other", 3:   GST_soObj.Add "Mismatch-Other", 4
    GST_soObj.Add "Missing in Portal", 5:    GST_soObj.Add "Missing in Books", 6
    GST_soObj.Add "Duplicate-Portal", 7:     GST_soObj.Add "Duplicate-Books", 8

    ReDim GST_soPri(1 To nR)
    For GST_soI = 1 To nR
        GST_soP = 99
        If GST_soObj.Exists(CStr(aR(GST_soI, 1))) Then GST_soP = GST_soObj(CStr(aR(GST_soI, 1)))
        GST_soPri(GST_soI) = GST_soP
        GST_soCnt(GST_soP) = GST_soCnt(GST_soP) + 1
    Next GST_soI

    GST_soRun = 1
    For GST_soP = 0 To 100
        GST_soPos(GST_soP) = GST_soRun
        GST_soRun = GST_soRun + GST_soCnt(GST_soP)
    Next GST_soP

    ReDim GST_soTmp(1 To nR, 1 To 24)
    For GST_soI = 1 To nR
        GST_soDst = GST_soPos(GST_soPri(GST_soI))
        GST_soPos(GST_soPri(GST_soI)) = GST_soDst + 1
        For GST_soC = 1 To 24
            GST_soTmp(GST_soDst, GST_soC) = aR(GST_soI, GST_soC)
        Next GST_soC
    Next GST_soI

    For GST_soI = 1 To nR
        For GST_soC = 1 To 24
            aR(GST_soI, GST_soC) = GST_soTmp(GST_soI, GST_soC)
        Next GST_soC
    Next GST_soI
End Sub

'==================================================================
' SHEET MANAGEMENT
'==================================================================
Function GST_MakeSheet() As Worksheet
    Application.DisplayAlerts = False
    If GST_HasSheet("Reco_Result") Then ThisWorkbook.Sheets("Reco_Result").Delete
    Application.DisplayAlerts = True
    Dim GST_mkWs As Worksheet
    Set GST_mkWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    GST_mkWs.Name = "Reco_Result": GST_mkWs.Tab.Color = RGB(60, 130, 75)
    Set GST_MakeSheet = GST_mkWs
End Function

Function GST_MakeSumSheet() As Worksheet
    Application.DisplayAlerts = False
    If GST_HasSheet("GST_Summary") Then ThisWorkbook.Sheets("GST_Summary").Delete
    Application.DisplayAlerts = True
    Dim GST_msWs As Worksheet
    Set GST_msWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    GST_msWs.Name = "GST_Summary": GST_msWs.Tab.Color = RGB(70, 115, 170)
    Set GST_MakeSumSheet = GST_msWs
End Function

'==================================================================
' FUZZY PARTY MATCHING
'==================================================================
Function GST_FuzzyParty(GST_fpA As String, GST_fpB As String) As Boolean
    Dim GST_fpNA As String
    Dim GST_fpNB As String
    GST_fpNA = GST_NormParty(GST_fpA)
    GST_fpNB = GST_NormParty(GST_fpB)
    If GST_fpNA = GST_fpNB Then GST_FuzzyParty = True: Exit Function
    If Len(GST_fpNA) > 4 And Len(GST_fpNB) > 4 Then
        If InStr(1, GST_fpNA, GST_fpNB, vbBinaryCompare) > 0 Then GST_FuzzyParty = True: Exit Function
        If InStr(1, GST_fpNB, GST_fpNA, vbBinaryCompare) > 0 Then GST_FuzzyParty = True: Exit Function
    End If
    If GST_BigramSim(GST_fpNA, GST_fpNB) >= GST_THR Then GST_FuzzyParty = True: Exit Function
    GST_FuzzyParty = False
End Function

Function GST_NormParty(GST_npS As String) As String
    Dim GST_npR As String
    GST_npR = UCase(Trim(GST_npS))
    If Left(GST_npR, 4) = "MRS." Then GST_npR = Trim(Mid(GST_npR, 5))
    If Left(GST_npR, 4) = "MRS " Then GST_npR = Trim(Mid(GST_npR, 5))
    If Left(GST_npR, 3) = "MRS" Then
        If Len(GST_npR) = 3 Or Mid(GST_npR, 4, 1) = " " Then GST_npR = Trim(Mid(GST_npR, 4))
    End If
    If Left(GST_npR, 3) = "MR." Then GST_npR = Trim(Mid(GST_npR, 4))
    If Left(GST_npR, 3) = "MR " Then GST_npR = Trim(Mid(GST_npR, 4))
    If Left(GST_npR, 2) = "MR" Then
        If Len(GST_npR) = 2 Or Mid(GST_npR, 3, 1) = " " Then GST_npR = Trim(Mid(GST_npR, 3))
    End If
    If Left(GST_npR, 3) = "MS." Then GST_npR = Trim(Mid(GST_npR, 4))
    If Left(GST_npR, 3) = "MS " Then GST_npR = Trim(Mid(GST_npR, 4))
    If Left(GST_npR, 2) = "MS" Then
        If Len(GST_npR) = 2 Or Mid(GST_npR, 3, 1) = " " Then GST_npR = Trim(Mid(GST_npR, 3))
    End If
    If Left(GST_npR, 3) = "DR." Then GST_npR = Trim(Mid(GST_npR, 4))
    If Left(GST_npR, 3) = "DR " Then GST_npR = Trim(Mid(GST_npR, 4))
    If Left(GST_npR, 2) = "DR" Then
        If Len(GST_npR) = 2 Or Mid(GST_npR, 3, 1) = " " Then GST_npR = Trim(Mid(GST_npR, 3))
    End If
    If Left(GST_npR, 5) = "SHRI " Then GST_npR = Trim(Mid(GST_npR, 6))
    If Left(GST_npR, 4) = "SMT." Then GST_npR = Trim(Mid(GST_npR, 5))
    If Left(GST_npR, 4) = "SMT " Then GST_npR = Trim(Mid(GST_npR, 5))
    If Left(GST_npR, 3) = "SMT" Then
        If Len(GST_npR) = 3 Or Mid(GST_npR, 4, 1) = " " Then GST_npR = Trim(Mid(GST_npR, 4))
    End If
    If Left(GST_npR, 4) = "M/S " Then GST_npR = Trim(Mid(GST_npR, 5))
    If Left(GST_npR, 4) = "M/S." Then GST_npR = Trim(Mid(GST_npR, 5))
    If Left(GST_npR, 3) = "M/S" Then GST_npR = Trim(Mid(GST_npR, 4))
    GST_npR = Replace(GST_npR, ".", ""):    GST_npR = Replace(GST_npR, ",", "")
    GST_npR = Replace(GST_npR, "-", ""):    GST_npR = Replace(GST_npR, "/", "")
    GST_npR = Replace(GST_npR, Chr(39), ""): GST_npR = Replace(GST_npR, "&", "AND")
    Do While InStr(1, GST_npR, "  ") > 0
        GST_npR = Replace(GST_npR, "  ", " ")
    Loop
    GST_npR = Replace(GST_npR, "PRIVATE LIMITED", "PVTLTD")
    GST_npR = Replace(GST_npR, "PVT LTD", "PVTLTD")
    GST_npR = Replace(GST_npR, "PVT LIMITED", "PVTLTD")
    GST_npR = Replace(GST_npR, "PRIVATE LTD", "PVTLTD")
    GST_npR = Replace(GST_npR, "(P) LTD", "PVTLTD")
    GST_npR = Replace(GST_npR, "(P)LTD", "PVTLTD")
    GST_npR = Replace(GST_npR, "LIMITED", "LTD")
    GST_npR = Replace(GST_npR, "PRIVATE", "PVT")
    GST_npR = Replace(GST_npR, "CORPORATION", "CORP")
    GST_npR = Replace(GST_npR, "COMPANY", "CO")
    GST_npR = Replace(GST_npR, "ENTERPRISES", "ENT")
    GST_npR = Replace(GST_npR, "ENTERPRISE", "ENT")
    GST_npR = Replace(GST_npR, "INDUSTRIES", "IND")
    GST_npR = Replace(GST_npR, "INDUSTRY", "IND")
    GST_npR = Replace(GST_npR, "TRADERS", "TRADING")
    GST_npR = Replace(GST_npR, " ", "")
    GST_NormParty = Trim(GST_npR)
End Function

Function GST_BigramSim(GST_bs1 As String, GST_bs2 As String) As Double
    Dim GST_bsL1    As Long
    Dim GST_bsL2    As Long
    Dim GST_bsHits  As Long
    Dim GST_bsI     As Long
    Dim GST_bsJ     As Long
    Dim GST_bsBG1() As String
    Dim GST_bsBG2() As String
    Dim GST_bsChk() As Boolean

    GST_bsL1 = Len(GST_bs1) - 1
    GST_bsL2 = Len(GST_bs2) - 1
    If GST_bsL1 <= 0 Or GST_bsL2 <= 0 Then
        If GST_bs1 = GST_bs2 Then GST_BigramSim = 1# Else GST_BigramSim = 0#
        Exit Function
    End If
    ReDim GST_bsBG1(1 To GST_bsL1)
    ReDim GST_bsBG2(1 To GST_bsL2)
    ReDim GST_bsChk(1 To GST_bsL2)
    For GST_bsI = 1 To GST_bsL1
        GST_bsBG1(GST_bsI) = Mid(GST_bs1, GST_bsI, 2)
    Next GST_bsI
    For GST_bsI = 1 To GST_bsL2
        GST_bsBG2(GST_bsI) = Mid(GST_bs2, GST_bsI, 2)
    Next GST_bsI
    GST_bsHits = 0
    For GST_bsI = 1 To GST_bsL1
        For GST_bsJ = 1 To GST_bsL2
            If Not GST_bsChk(GST_bsJ) Then
                If GST_bsBG1(GST_bsI) = GST_bsBG2(GST_bsJ) Then
                    GST_bsHits = GST_bsHits + 1
                    GST_bsChk(GST_bsJ) = True
                    Exit For
                End If
            End If
        Next GST_bsJ
    Next GST_bsI
    GST_BigramSim = (2# * CDbl(GST_bsHits)) / CDbl(GST_bsL1 + GST_bsL2)
End Function

'==================================================================
' INVOICE NUMBER NORMALISATION
'==================================================================
Function GST_NormInv(GST_niS As String) As String
    Dim GST_niR   As String
    Dim GST_niPos As Long
    Dim GST_niPfx As String
    Dim GST_niSfx As String

    GST_niR = UCase(Trim(GST_niS))
    If Len(GST_niR) = 0 Then GST_NormInv = "": Exit Function
    If Left(GST_niR, 2) = "FY" Then GST_niR = Mid(GST_niR, 3)
    GST_niR = GST_StripFYPfx(GST_niR)
    GST_niR = GST_StripFY4Pfx(GST_niR)
    GST_niR = GST_StripFYSfx(GST_niR)
    GST_niR = Replace(GST_niR, " ", ""):    GST_niR = Replace(GST_niR, ",", "")
    GST_niR = Replace(GST_niR, "-", ""):    GST_niR = Replace(GST_niR, "/", "")
    GST_niR = Replace(GST_niR, ".", ""):    GST_niR = Replace(GST_niR, Chr(9), "")
    GST_niR = Replace(GST_niR, Chr(92), ""): GST_niR = Replace(GST_niR, Chr(160), "")
    If Len(GST_niR) = 0 Then GST_NormInv = "": Exit Function
    If GST_AllDigits(GST_niR) Then GST_NormInv = GST_StripLZ(GST_niR): Exit Function
    GST_niPfx = "": GST_niSfx = ""
    For GST_niPos = Len(GST_niR) To 1 Step -1
        If Not IsNumeric(Mid(GST_niR, GST_niPos, 1)) Then
            GST_niPfx = Left(GST_niR, GST_niPos)
            GST_niSfx = Mid(GST_niR, GST_niPos + 1)
            Exit For
        End If
    Next GST_niPos
    If Len(GST_niSfx) > 0 And GST_AllDigits(GST_niSfx) Then
        GST_NormInv = GST_niPfx & GST_StripLZ(GST_niSfx)
    Else
        GST_NormInv = GST_niR
    End If
End Function

Function GST_InvNumSuffix(GST_isK As String) As String
    Dim GST_isI   As Long
    Dim GST_isPfx As String
    Dim GST_isSfx As String
    If Len(GST_isK) = 0 Then GST_InvNumSuffix = "": Exit Function
    If GST_AllDigits(GST_isK) Then GST_InvNumSuffix = "": Exit Function
    For GST_isI = Len(GST_isK) To 1 Step -1
        If Not IsNumeric(Mid(GST_isK, GST_isI, 1)) Then
            GST_isPfx = Left(GST_isK, GST_isI)
            GST_isSfx = Mid(GST_isK, GST_isI + 1)
            Exit For
        End If
    Next GST_isI
    If Len(GST_isPfx) = 0 Or Len(GST_isSfx) = 0 Then GST_InvNumSuffix = "": Exit Function
    If Not GST_AllDigits(GST_isSfx) Then GST_InvNumSuffix = "": Exit Function
    GST_InvNumSuffix = GST_StripLZ(GST_isSfx)
End Function

Function GST_InvNumPrefix(GST_ipK As String) As String
    Dim GST_ipI      As Long
    Dim GST_ipAlpEnd As Long
    Dim GST_ipNum    As String
    If Len(GST_ipK) = 0 Then GST_InvNumPrefix = "": Exit Function
    If GST_AllDigits(GST_ipK) Then GST_InvNumPrefix = "": Exit Function
    GST_ipAlpEnd = 0
    For GST_ipI = 1 To Len(GST_ipK)
        If IsNumeric(Mid(GST_ipK, GST_ipI, 1)) Then
            Exit For
        Else
            GST_ipAlpEnd = GST_ipI
        End If
    Next GST_ipI
    If GST_ipAlpEnd = 0 Then GST_InvNumPrefix = "": Exit Function
    GST_ipNum = Mid(GST_ipK, GST_ipAlpEnd + 1)
    If Len(GST_ipNum) = 0 Then GST_InvNumPrefix = "": Exit Function
    If Not GST_AllDigits(GST_ipNum) Then GST_InvNumPrefix = "": Exit Function
    GST_InvNumPrefix = GST_StripLZ(GST_ipNum)
End Function

Function GST_StripFY4Pfx(GST_s4S As String) As String
    Dim GST_s4Sl  As Long
    Dim GST_s4Tag As String
    Dim GST_s4Y1  As Long
    Dim GST_s4Y2  As Long
    Dim GST_s4Rst As String
    GST_s4Sl = InStr(1, GST_s4S, "/")
    If GST_s4Sl <> 5 Then GST_StripFY4Pfx = GST_s4S: Exit Function
    GST_s4Tag = Left(GST_s4S, 4)
    If Not GST_AllDigits(GST_s4Tag) Then GST_StripFY4Pfx = GST_s4S: Exit Function
    GST_s4Y1 = CInt(Left(GST_s4Tag, 2)): GST_s4Y2 = CInt(Right(GST_s4Tag, 2))
    If GST_s4Y2 = GST_s4Y1 + 1 Then
        GST_s4Rst = Mid(GST_s4S, 6)
        If Len(Trim(GST_s4Rst)) > 0 Then GST_StripFY4Pfx = Trim(GST_s4Rst): Exit Function
    End If
    GST_StripFY4Pfx = GST_s4S
End Function

Function GST_StripFYSfx(GST_ssxS As String) As String
    ' Strips a trailing financial-year tag after the LAST "/".
    ' A bare 2-digit / 4-digit tail is treated as a year ONLY when another
    ' "/" part precedes it (INV/123/25). "INV/12" keeps 12 as the invoice no.
    ' Explicit FY tails (24-25, 2024-25) are always stripped.
    Dim GST_ssxPos  As Long
    Dim GST_ssxBase As String
    Dim GST_ssxTail As String
    Dim GST_ssxDa   As Long
    Dim GST_ssxI    As Long
    Dim GST_ssxY1   As Long
    Dim GST_ssxY2   As Long
    Dim GST_ssxP1   As String
    Dim GST_ssxP2   As String
    Dim GST_ssxMore As Boolean
    For GST_ssxI = Len(GST_ssxS) To 1 Step -1
        If Mid(GST_ssxS, GST_ssxI, 1) = "/" Then GST_ssxPos = GST_ssxI: Exit For
    Next GST_ssxI
    If GST_ssxPos < 2 Then GST_StripFYSfx = GST_ssxS: Exit Function
    GST_ssxBase = Left(GST_ssxS, GST_ssxPos - 1)
    GST_ssxTail = Mid(GST_ssxS, GST_ssxPos + 1)
    If Len(GST_ssxBase) = 0 Or Len(GST_ssxTail) = 0 Then GST_StripFYSfx = GST_ssxS: Exit Function

    GST_ssxMore = (InStr(1, GST_ssxBase, "/") > 0)

    If GST_ssxMore And Len(GST_ssxTail) = 4 And GST_AllDigits(GST_ssxTail) Then
        GST_ssxY1 = CInt(Left(GST_ssxTail, 2)): GST_ssxY2 = CInt(Right(GST_ssxTail, 2))
        If GST_ssxY2 = GST_ssxY1 + 1 Then GST_StripFYSfx = GST_ssxBase: Exit Function
        If CInt(GST_ssxTail) >= 2000 And CInt(GST_ssxTail) <= 2099 Then
            GST_StripFYSfx = GST_ssxBase: Exit Function
        End If
    End If
    If GST_ssxMore And Len(GST_ssxTail) = 2 And GST_AllDigits(GST_ssxTail) Then
        GST_StripFYSfx = GST_ssxBase: Exit Function
    End If

    GST_ssxDa = InStr(1, GST_ssxTail, "-")
    If GST_ssxDa = 3 Or GST_ssxDa = 5 Then
        GST_ssxP1 = Left(GST_ssxTail, GST_ssxDa - 1)
        GST_ssxP2 = Mid(GST_ssxTail, GST_ssxDa + 1)
        If Len(GST_ssxP2) = 2 And GST_AllDigits(GST_ssxP1) And GST_AllDigits(GST_ssxP2) Then
            GST_ssxY1 = CInt(Right(GST_ssxP1, 2)): GST_ssxY2 = CInt(GST_ssxP2)
            If GST_ssxY2 = (GST_ssxY1 + 1) Mod 100 Then
                GST_StripFYSfx = GST_ssxBase: Exit Function
            End If
        End If
    End If
    GST_StripFYSfx = GST_ssxS
End Function

Function GST_StripFYPfx(GST_sfS As String) As String
    Dim GST_sfSl   As Long
    Dim GST_sfDa   As Long
    Dim GST_sfSl2  As Long
    Dim GST_sfPt1  As String
    Dim GST_sfPt2  As String
    Dim GST_sfRest As String
    GST_sfDa = InStr(1, GST_sfS, "-")
    If GST_sfDa >= 2 And GST_sfDa <= 5 Then
        GST_sfSl = InStr(GST_sfDa + 1, GST_sfS, "/")
        If GST_sfSl > GST_sfDa + 1 Then
            GST_sfPt1 = Left(GST_sfS, GST_sfDa - 1)
            GST_sfPt2 = Mid(GST_sfS, GST_sfDa + 1, GST_sfSl - GST_sfDa - 1)
            GST_sfRest = Mid(GST_sfS, GST_sfSl + 1)
            If Len(GST_sfPt1) >= 2 And Len(GST_sfPt1) <= 4 And Len(GST_sfPt2) = 2 Then
                If GST_AllDigits(GST_sfPt1) And GST_AllDigits(GST_sfPt2) And Len(Trim(GST_sfRest)) > 0 Then
                    GST_StripFYPfx = Trim(GST_sfRest): Exit Function
                End If
            End If
        End If
    End If
    GST_sfSl = InStr(1, GST_sfS, "/")
    If GST_sfSl >= 2 And GST_sfSl <= 5 Then
        GST_sfSl2 = InStr(GST_sfSl + 1, GST_sfS, "/")
        If GST_sfSl2 > GST_sfSl + 1 Then
            GST_sfPt1 = Left(GST_sfS, GST_sfSl - 1)
            GST_sfPt2 = Mid(GST_sfS, GST_sfSl + 1, GST_sfSl2 - GST_sfSl - 1)
            GST_sfRest = Mid(GST_sfS, GST_sfSl2 + 1)
            If Len(GST_sfPt1) = 2 And Len(GST_sfPt2) = 2 Then
                If GST_AllDigits(GST_sfPt1) And GST_AllDigits(GST_sfPt2) And Len(Trim(GST_sfRest)) > 0 Then
                    GST_StripFYPfx = Trim(GST_sfRest): Exit Function
                End If
            End If
        End If
    End If
    GST_StripFYPfx = GST_sfS
End Function

Function GST_AllDigits(GST_adS As String) As Boolean
    Dim GST_adI As Long
    If Len(GST_adS) = 0 Then GST_AllDigits = False: Exit Function
    For GST_adI = 1 To Len(GST_adS)
        If Mid(GST_adS, GST_adI, 1) < "0" Or Mid(GST_adS, GST_adI, 1) > "9" Then
            GST_AllDigits = False: Exit Function
        End If
    Next GST_adI
    GST_AllDigits = True
End Function

Function GST_StripLZ(GST_szS As String) As String
    Dim GST_szI As Long
    If Len(GST_szS) = 0 Then GST_StripLZ = "": Exit Function
    For GST_szI = 1 To Len(GST_szS) - 1
        If Mid(GST_szS, GST_szI, 1) <> "0" Then
            GST_StripLZ = Mid(GST_szS, GST_szI): Exit Function
        End If
    Next GST_szI
    GST_StripLZ = Right(GST_szS, 1)
End Function

'==================================================================
' DATE UTILITIES
'==================================================================
Function GST_ParseDate(GST_pdV As Variant, ByRef GST_pdOk As Boolean) As Date
    ' Locale-independent: numeric text is read as dd/mm/yyyy (or yyyy/mm/dd)
    Dim GST_pdS As String
    Dim GST_pdP() As String
    Dim GST_pd3   As String
    Dim GST_pdY   As Long
    Dim GST_pdM   As Long
    Dim GST_pdD   As Long
    Dim GST_pdT   As Date

    GST_pdOk = False
    If IsError(GST_pdV) Then Exit Function
    If IsEmpty(GST_pdV) Then Exit Function

    If VarType(GST_pdV) = vbDate Then
        GST_ParseDate = CDate(GST_pdV): GST_pdOk = True: Exit Function
    End If
    If VarType(GST_pdV) = vbDouble Then
        If GST_pdV >= 20000 And GST_pdV <= 80000 Then
            GST_ParseDate = CDate(GST_pdV): GST_pdOk = True: Exit Function
        End If
    End If

    GST_pdS = Trim(CStr(GST_pdV))
    If Len(GST_pdS) = 0 Then Exit Function
    GST_pdS = Replace(Replace(GST_pdS, "-", "/"), ".", "/")

    GST_pdP = Split(GST_pdS, "/")
    If UBound(GST_pdP) = 2 Then
        GST_pd3 = Trim(GST_pdP(2))
        If InStr(1, GST_pd3, " ") > 0 Then GST_pd3 = Left(GST_pd3, InStr(1, GST_pd3, " ") - 1)
        If GST_AllDigits(Trim(GST_pdP(0))) And GST_AllDigits(Trim(GST_pdP(1))) And GST_AllDigits(GST_pd3) Then
            If Len(Trim(GST_pdP(0))) = 4 Then
                GST_pdY = CLng(Trim(GST_pdP(0))): GST_pdM = CLng(Trim(GST_pdP(1))): GST_pdD = CLng(GST_pd3)
            Else
                GST_pdD = CLng(Trim(GST_pdP(0))): GST_pdM = CLng(Trim(GST_pdP(1))): GST_pdY = CLng(GST_pd3)
            End If
            If GST_pdY < 100 Then GST_pdY = GST_pdY + 2000
            If GST_pdM >= 1 And GST_pdM <= 12 And GST_pdD >= 1 And GST_pdD <= 31 Then
                GST_pdT = DateSerial(GST_pdY, GST_pdM, GST_pdD)
                If Day(GST_pdT) = GST_pdD Then
                    GST_ParseDate = GST_pdT: GST_pdOk = True: Exit Function
                End If
            End If
            Exit Function
        End If
    End If

    ' Text such as 05-Mar-2025 : let VBA parse it
    On Error Resume Next
    If IsDate(GST_pdS) Then
        GST_ParseDate = CDate(GST_pdS)
        GST_pdOk = (Err.Number = 0)
    End If
    On Error GoTo 0
End Function

Function GST_NormDate(GST_ndV As Variant) As String
    Dim GST_ndOk As Boolean
    Dim GST_ndD  As Date
    GST_ndD = GST_ParseDate(GST_ndV, GST_ndOk)
    If GST_ndOk Then
        GST_NormDate = Format(GST_ndD, "YYYYMMDD")
    Else
        GST_NormDate = Trim(CStr(GST_ndV))
    End If
End Function

Function GST_FmtDate(GST_fdV As Variant) As String
    Dim GST_fdOk As Boolean
    Dim GST_fdD  As Date
    GST_fdD = GST_ParseDate(GST_fdV, GST_fdOk)
    If GST_fdOk Then
        GST_FmtDate = Format(GST_fdD, "DD-MMM-YYYY")
    Else
        GST_FmtDate = Trim(CStr(GST_fdV))
    End If
End Function

'==================================================================
' NUMERIC & SHEET UTILITIES
'==================================================================
Function GST_Rnd2(GST_r2V As Double) As Double
    ' Round half away from zero (VBA Round() is banker's rounding)
    GST_Rnd2 = Sgn(GST_r2V) * Int(Abs(GST_r2V) * 100# + 0.5) / 100#
End Function

Function GST_AmtDiff(GST_adA As Double, GST_adB As Double) As Boolean
    ' True when two amounts differ by more than the tolerance
    GST_AmtDiff = (Abs(GST_adA - GST_adB) > GST_TOL + 0.001)
End Function

Function GST_ToD(GST_tdV As Variant) As Double
    On Error Resume Next
    If IsNumeric(GST_tdV) Then GST_ToD = CDbl(GST_tdV) Else GST_ToD = 0
    On Error GoTo 0
End Function

Function GST_HasSheet(GST_hsN As String) As Boolean
    Dim GST_hsWs As Worksheet
    On Error Resume Next
    Set GST_hsWs = ThisWorkbook.Sheets(GST_hsN)
    On Error GoTo 0
    GST_HasSheet = Not (GST_hsWs Is Nothing)
End Function

Function GST_LastRow(GST_lrWs As Worksheet) As Long
    GST_LastRow = GST_lrWs.Cells(GST_lrWs.Rows.Count, 1).End(xlUp).Row
End Function

'==================================================================
' COLOUR HELPERS
'==================================================================
Function GST_RowClr(GST_rcSt As String) As Long
    If GST_rcSt = "Matched" Then GST_RowClr = GST_CM
    If GST_rcSt = "Mismatch-Amount" Then GST_RowClr = GST_CA
    If GST_rcSt = "Mismatch-Amt+Other" Then GST_RowClr = GST_CA
    If GST_rcSt = "Mismatch-Other" Then GST_RowClr = GST_CO
    If GST_rcSt = "Missing in Portal" Then GST_RowClr = GST_CP
    If GST_rcSt = "Missing in Books" Then GST_RowClr = GST_CT
    If GST_rcSt = "Duplicate-Portal" Then GST_RowClr = GST_CD
    If GST_rcSt = "Duplicate-Books" Then GST_RowClr = GST_CD
    If GST_RowClr = 0 Then GST_RowClr = RGB(248, 249, 253)
End Function

Function GST_BnrClr(GST_bcSt As String) As Long
    If GST_bcSt = "Matched" Then GST_BnrClr = RGB(55, 130, 75)
    If GST_bcSt = "Mismatch-Amount" Then GST_BnrClr = RGB(190, 120, 35)
    If GST_bcSt = "Mismatch-Amt+Other" Then GST_BnrClr = RGB(175, 105, 25)
    If GST_bcSt = "Mismatch-Other" Then GST_BnrClr = RGB(130, 145, 35)
    If GST_bcSt = "Missing in Portal" Then GST_BnrClr = RGB(185, 70, 70)
    If GST_bcSt = "Missing in Books" Then GST_BnrClr = RGB(55, 120, 165)
    If GST_bcSt = "Duplicate-Portal" Then GST_BnrClr = RGB(110, 80, 155)
    If GST_bcSt = "Duplicate-Books" Then GST_BnrClr = RGB(95, 65, 140)
    If GST_BnrClr = 0 Then GST_BnrClr = RGB(90, 100, 120)
End Function

'==================================================================
' BUTTON CLICK HANDLER
'==================================================================
Sub GST_Button_Click()
    GST_Run
End Sub
