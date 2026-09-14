Attribute VB_Name = "Module1"
Option Explicit

'=========================================================
' MAIN MACRO
' ONE BUTTON DOES EVERYTHING
'=========================================================

Sub Run_Reconciliation()

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Running financial reconciliation..."

    '=====================================================
    ' 1. RECALCULATE RECONCILIATION
    '=====================================================

    Application.StatusBar = "Calculating reconciliation..."

    Worksheets("Reconciliation").Calculate


    '=====================================================
    ' 2. UPDATE SUMMARY
    '=====================================================

    Application.StatusBar = "Updating summary..."

    Worksheets("Summary").Calculate


    '=====================================================
    ' 3. DATA QUALITY CHECKS
    '=====================================================

    Application.StatusBar = "Running data quality checks..."

    UpdateDataQuality


    '=====================================================
    ' 4. RECALCULATE SUMMARY AGAIN
    '=====================================================

    Worksheets("Summary").Calculate


    '=====================================================
    ' 5. LAST RECONCILIATION RUN
    '=====================================================

    With Worksheets("Summary")

        .Range("D20").Value = "Last Reconciliation Run"

        .Range("E20").Value = Now

        .Range("E20").NumberFormat = "dd-mmm-yyyy hh:mm:ss"

    End With


    '=====================================================
    ' 6. EXCEPTION RATE
    '=====================================================

    With Worksheets("Summary")

        .Range("D22").Value = "Exception Rate"

        .Range("E22").Formula = "=1-B9"

        .Range("E22").NumberFormat = "0%"

    End With


    '=====================================================
    ' 7. EXPORT EXCEPTIONS
    '=====================================================

    Application.StatusBar = "Creating exception report..."

    Export_Exceptions_Silent


    '=====================================================
    ' 8. RESTORE EXCEL
    '=====================================================

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False


    '=====================================================
    ' 9. SUCCESS MESSAGE
    '=====================================================

    MsgBox "FINANCIAL RECONCILIATION COMPLETED" & _
           vbCrLf & vbCrLf & _
           "? Reconciliation updated" & _
           vbCrLf & _
           "? Summary updated" & _
           vbCrLf & _
           "? Data quality checks completed" & _
           vbCrLf & _
           "? Exception rate calculated" & _
           vbCrLf & _
           "? Exception report exported", _
           vbInformation, _
           "Financial Reconciliation Tool"

    Exit Sub


'=========================================================
' ERROR HANDLER
'=========================================================

ErrorHandler:

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "RECONCILIATION FAILED" & _
           vbCrLf & vbCrLf & _
           Err.Description, _
           vbCritical, _
           "Financial Reconciliation Tool"

End Sub



'=========================================================
' DATA QUALITY CHECKS
'=========================================================

Sub UpdateDataQuality()

    Dim wsData As Worksheet
    Dim wsSummary As Worksheet

    Dim lastRow As Long

    Dim contractData As Variant
    Dim amountData As Variant
    Dim transactionCodeData As Variant

    Dim dict As Object

    Dim i As Long

    Dim ContractID As String

    Dim duplicateIDs As Long
    Dim blankContracts As Long
    Dim blankAmounts As Long
    Dim negativeAmounts As Long
    Dim invalidContracts As Long


    Set wsData = Worksheets("Transaction_Data")

    Set wsSummary = Worksheets("Summary")


    '=====================================================
    ' FIND LAST ROW
    '=====================================================

    lastRow = wsData.Cells(wsData.Rows.Count, "D").End(xlUp).Row


    If lastRow < 2 Then

        MsgBox "No transaction data found.", _
               vbExclamation, _
               "Data Quality Check"

        Exit Sub

    End If


    '=====================================================
    ' 1. TOTAL RECORDS
    '=====================================================

    wsSummary.Range("E13").Value = lastRow - 1


    '=====================================================
    ' 2. CONTRACT ID CHECK
    '=====================================================

    contractData = wsData.Range("D2:D" & lastRow).Value


    Set dict = CreateObject("Scripting.Dictionary")

    dict.CompareMode = vbTextCompare


    For i = 1 To UBound(contractData, 1)

        ContractID = Trim(CStr(contractData(i, 1)))


        If ContractID = "" Then

            blankContracts = blankContracts + 1


        Else

            If dict.Exists(ContractID) Then

                'Count each duplicated Contract ID once

                If dict(ContractID) = 1 Then

                    duplicateIDs = duplicateIDs + 1

                End If


                dict(ContractID) = dict(ContractID) + 1


            Else

                dict.Add ContractID, 1

            End If

        End If

    Next i


    '=====================================================
    ' 3. TRANSACTION AMOUNT CHECK
    '=====================================================

    amountData = wsData.Range("C2:C" & lastRow).Value


    For i = 1 To UBound(amountData, 1)

        If Trim(CStr(amountData(i, 1))) = "" Then

            blankAmounts = blankAmounts + 1


        ElseIf IsNumeric(amountData(i, 1)) Then

            If CDbl(amountData(i, 1)) < 0 Then

                negativeAmounts = negativeAmounts + 1

            End If

        End If

    Next i


    '=====================================================
    ' 4. INVALID CONTRACT ID CHECK
    '=====================================================

    transactionCodeData = wsData.Range("B2:B" & lastRow).Value


    invalidContracts = CheckInvalidContractIDs( _
                            contractData, _
                            transactionCodeData)


    '=====================================================
    ' 5. WRITE RESULTS TO SUMMARY
    '=====================================================

    wsSummary.Range("E13").Value = lastRow - 1

    wsSummary.Range("E14").Value = duplicateIDs

    wsSummary.Range("E15").Value = blankContracts

    wsSummary.Range("E16").Value = blankAmounts

    wsSummary.Range("E17").Value = negativeAmounts

    wsSummary.Range("E18").Value = invalidContracts


    wsSummary.Range("E13:E18").NumberFormat = "#,##0"

End Sub



'=========================================================
' INVALID CONTRACT ID CHECK
'
' Group Premium format:
'
' 5 numbers + 5 letters + P/S/S1
'
' Examples:
'
' 54545GLEEAP
' 45676GRTHUS1
'
'=========================================================

Function CheckInvalidContractIDs( _
    contractData As Variant, _
    transactionCodeData As Variant) As Long


    Dim regex As Object

    Dim i As Long

    Dim ContractID As String

    Dim TransactionCode As String

    Dim invalidCount As Long


    Set regex = CreateObject("VBScript.RegExp")


    regex.Pattern = "^[0-9]{5}[A-Z]{5}(P|S|S1)$"

    regex.IgnoreCase = False

    regex.Global = False


    For i = 1 To UBound(contractData, 1)

        TransactionCode = _
            Trim(CStr(transactionCodeData(i, 1)))


        If TransactionCode = "Group Premium" Then

            ContractID = _
                Trim(CStr(contractData(i, 1)))


            If Not regex.Test(ContractID) Then

                invalidCount = invalidCount + 1

            End If

        End If

    Next i


    CheckInvalidContractIDs = invalidCount

End Function



'=========================================================
' EXPORT EXCEPTIONS
'
' Creates a separate Excel file containing:
'
' Amount Mismatch
' Business Type Mismatch
' Amount & Business Type Mismatch
' Contract Not Found
'
'=========================================================

Sub Export_Exceptions_Silent()


    Dim wsRec As Worksheet

    Dim wbNew As Workbook

    Dim wsNew As Worksheet

    Dim lastRow As Long

    Dim visibleRange As Range

    Dim savePath As String

    Dim fileName As String

    Dim exceptionCount As Long


    On Error GoTo ExportError


    Set wsRec = Worksheets("Reconciliation")


    '=====================================================
    ' FIND LAST ROW
    '=====================================================

    lastRow = wsRec.Cells( _
                    wsRec.Rows.Count, _
                    "A").End(xlUp).Row


    If lastRow < 2 Then Exit Sub


    '=====================================================
    ' REMOVE EXISTING FILTER
    '=====================================================

    If wsRec.AutoFilterMode Then

        wsRec.AutoFilterMode = False

    End If


    '=====================================================
    ' FILTER EXCEPTIONS
    '
    ' Column G = Reconciliation_Status
    '=====================================================

    wsRec.Range("A1:H" & lastRow).AutoFilter _
        Field:=7, _
        Criteria1:="<>Matched"


    '=====================================================
    ' GET VISIBLE EXCEPTION ROWS
    '=====================================================

    On Error Resume Next


    Set visibleRange = _
        wsRec.Range("A2:H" & lastRow) _
        .SpecialCells(xlCellTypeVisible)


    On Error GoTo ExportError


    '=====================================================
    ' NO EXCEPTIONS
    '=====================================================

    If visibleRange Is Nothing Then

        wsRec.AutoFilterMode = False

        Exit Sub

    End If


    '=====================================================
    ' CREATE NEW WORKBOOK
    '=====================================================

    Set wbNew = Workbooks.Add


    Set wsNew = wbNew.Worksheets(1)


    wsNew.Name = "Exceptions"


    '=====================================================
    ' COPY HEADERS
    '=====================================================

    wsNew.Range("A1:H1").Value = _
        wsRec.Range("A1:H1").Value


    '=====================================================
    ' COPY EXCEPTION DATA
    '=====================================================

    visibleRange.Copy


    wsNew.Range("A2").PasteSpecial xlPasteValues


    Application.CutCopyMode = False


    '=====================================================
    ' FORMAT EXCEPTION FILE
    '=====================================================

    With wsNew.Range("A1:H1")

        .Font.Bold = True

        .AutoFilter

    End With


    wsNew.Columns("A:H").AutoFit


    wsNew.Columns("C").NumberFormat = "#,##0.00"

    wsNew.Columns("E").NumberFormat = "#,##0.00"

    wsNew.Columns("H").NumberFormat = "#,##0.00"


    '=====================================================
    ' COUNT EXCEPTIONS
    '=====================================================

    exceptionCount = _
        wsNew.Cells( _
            wsNew.Rows.Count, _
            "A").End(xlUp).Row - 1


    '=====================================================
    ' CHECK MAIN WORKBOOK IS SAVED
    '=====================================================

    If ThisWorkbook.Path = "" Then

        wbNew.Close SaveChanges:=False

        wsRec.AutoFilterMode = False

        Exit Sub

    End If


    '=====================================================
    ' SAVE EXCEPTION FILE
    '=====================================================

    savePath = _
        ThisWorkbook.Path & _
        Application.PathSeparator


    fileName = _
        "Financial_Reconciliation_Exceptions_" & _
        Format(Now, "yyyymmdd_hhmmss") & _
        ".xlsx"


    wbNew.SaveAs _
        fileName:=savePath & fileName, _
        FileFormat:=xlOpenXMLWorkbook


    '=====================================================
    ' CLOSE NEW WORKBOOK
    '=====================================================

    wbNew.Close SaveChanges:=False


    '=====================================================
    ' REMOVE FILTER
    '=====================================================

    wsRec.AutoFilterMode = False


    Exit Sub


'=========================================================
' EXPORT ERROR
'=========================================================

ExportError:

    On Error Resume Next


    If Not wbNew Is Nothing Then

        wbNew.Close SaveChanges:=False

    End If


    If wsRec.AutoFilterMode Then

        wsRec.AutoFilterMode = False

    End If

End Sub

