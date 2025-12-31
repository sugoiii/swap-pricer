Option Explicit

' ====== DLL name/path ======
#If VBA7 Then
    Public Const SWAP_PRICER_DLL As String = "C:\path\to\swap_pricer.dll"
#Else
    Public Const SWAP_PRICER_DLL As String = "C:\path\to\swap_pricer.dll"
#End If

' ====== Struct definitions (must match C++ POD layout) ======

Public Type VBACurveInput
    id As LongPtr
    pillarSerials As LongPtr
    discountRates As LongPtr
    tenorStrings As LongPtr
    pillarCount As Long
    dayCountCode As Long
End Type

Public Type VBAFixingInput
    indexName As LongPtr
    fixingDateSerials As LongPtr
    fixingRates As LongPtr
    fixingCount As Long
End Type

Public Type VBABucketConfig
    curveId As LongPtr
    tenorStrings As LongPtr
    tenorCount As Long
    bumpSize As Double
End Type

Public Type VBALegSpec
    legType As Long
    payReceive As Long
    notional As Double
    startDateSerial As Double
    endDateSerial As Double
    frequencyCode As Long
    dayCountCode As Long
    bdcCode As Long
    fixedRate As Double
    indexName As LongPtr
    fixingDays As Long
    spread As Double
    isCompounded As Long
End Type

Public Type VBASwapSpec
    swapType As Long
    leg1 As VBALegSpec
    leg2 As VBALegSpec
    discountCurveId As LongPtr
    valuationCurveId As LongPtr
    valuationDateSerial As Double
End Type

' ====== Buffer holders to keep data alive ======

Public Type CurveBuffers
    curveId As String
    pillarSerials() As Double
    discountRates() As Double
    tenorStrings() As String
    tenorPtrs() As LongPtr
End Type

Public Type FixingBuffers
    indexName As String
    fixingDates() As Double
    fixingRates() As Double
End Type

Public Type BucketBuffers
    curveId As String
    tenorStrings() As String
    tenorPtrs() As LongPtr
End Type

Public Type LegBuffers
    indexName As String
End Type

Public Type SwapBuffers
    discountCurveId As String
    valuationCurveId As String
    leg1 As LegBuffers
    leg2 As LegBuffers
End Type

' ====== Module-level buffers to keep data alive across DLL calls ======

Private mCurves() As VBACurveInput
Private mCurveBuffers() As CurveBuffers
Private mFixings() As VBAFixingInput
Private mFixingBuffers() As FixingBuffers
Private mBuckets() As VBABucketConfig
Private mBucketBuffers() As BucketBuffers
Private mSwapBuffers As SwapBuffers
Private mHolidaySerials() As Double

' ====== DLL import ======

#If VBA7 Then
    Public Declare PtrSafe Function IRS_PRICE_AND_BUCKETS Lib SWAP_PRICER_DLL _
        (ByRef swapSpec As VBASwapSpec, _
         ByRef curveInputs As VBACurveInput, ByVal curveCount As Long, _
         ByRef fixingInputs As VBAFixingInput, ByVal fixingCount As Long, _
         ByRef bucketInputs As VBABucketConfig, ByVal bucketCount As Long, _
         ByRef holidaySerials As Double, ByVal holidayCount As Long, _
         ByRef outPillarSerials As Double, ByRef outDeltas As Double, _
         ByVal maxBuckets As Long, ByRef outUsedBuckets As Long) As Double
    Public Declare PtrSafe Function IRS_PRICE_AND_BUCKETS_PTR Lib SWAP_PRICER_DLL Alias "IRS_PRICE_AND_BUCKETS" _
        (ByRef swapSpec As VBASwapSpec, _
         ByRef curveInputs As VBACurveInput, ByVal curveCount As Long, _
         ByRef fixingInputs As VBAFixingInput, ByVal fixingCount As Long, _
         ByRef bucketInputs As VBABucketConfig, ByVal bucketCount As Long, _
         ByVal holidaySerials As LongPtr, ByVal holidayCount As Long, _
         ByRef outPillarSerials As Double, ByRef outDeltas As Double, _
         ByVal maxBuckets As Long, ByRef outUsedBuckets As Long) As Double
    Public Declare PtrSafe Function IRS_LAST_ERROR Lib SWAP_PRICER_DLL () As LongPtr
    Public Declare PtrSafe Function IRS_IS_NAN Lib SWAP_PRICER_DLL (ByVal value As Double) As Long
    Public Declare PtrSafe Sub IRS_SET_DEBUG_MODE Lib SWAP_PRICER_DLL (ByVal enabled As Long)
    Public Declare PtrSafe Sub IRS_SET_DEBUG_LOG_PATH Lib SWAP_PRICER_DLL (ByVal path As LongPtr)
    Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByVal dest As LongPtr, ByVal src As LongPtr, ByVal cb As LongPtr)
#Else
    Public Declare Function IRS_PRICE_AND_BUCKETS Lib SWAP_PRICER_DLL _
        (ByRef swapSpec As VBASwapSpec, _
         ByRef curveInputs As VBACurveInput, ByVal curveCount As Long, _
         ByRef fixingInputs As VBAFixingInput, ByVal fixingCount As Long, _
         ByRef bucketInputs As VBABucketConfig, ByVal bucketCount As Long, _
         ByRef holidaySerials As Double, ByVal holidayCount As Long, _
         ByRef outPillarSerials As Double, ByRef outDeltas As Double, _
         ByVal maxBuckets As Long, ByRef outUsedBuckets As Long) As Double
    Public Declare Function IRS_PRICE_AND_BUCKETS_PTR Lib SWAP_PRICER_DLL Alias "IRS_PRICE_AND_BUCKETS" _
        (ByRef swapSpec As VBASwapSpec, _
         ByRef curveInputs As VBACurveInput, ByVal curveCount As Long, _
         ByRef fixingInputs As VBAFixingInput, ByVal fixingCount As Long, _
         ByRef bucketInputs As VBABucketConfig, ByVal bucketCount As Long, _
         ByVal holidaySerials As LongPtr, ByVal holidayCount As Long, _
         ByRef outPillarSerials As Double, ByRef outDeltas As Double, _
         ByVal maxBuckets As Long, ByRef outUsedBuckets As Long) As Double
    Public Declare Function IRS_LAST_ERROR Lib SWAP_PRICER_DLL () As LongPtr
    Public Declare Function IRS_IS_NAN Lib SWAP_PRICER_DLL (ByVal value As Double) As Long
    Public Declare Sub IRS_SET_DEBUG_MODE Lib SWAP_PRICER_DLL (ByVal enabled As Long)
    Public Declare Sub IRS_SET_DEBUG_LOG_PATH Lib SWAP_PRICER_DLL (ByVal path As LongPtr)
    Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByVal dest As LongPtr, ByVal src As LongPtr, ByVal cb As LongPtr)
#End If

' ====== Helpers ======

Private Function StringToUtf16Buffer(ByVal value As String) As Byte()
    Dim raw() As Byte
    Dim result() As Byte
    Dim length As Long

    If Len(value) > 0 Then
        raw = StrConv(value, vbUnicode)
        length = UBound(raw) - LBound(raw) + 1
        ReDim result(0 To length + 1)
        CopyMemory VarPtr(result(0)), VarPtr(raw(0)), length
    Else
        ReDim result(0 To 1)
    End If

    result(UBound(result) - 1) = 0
    result(UBound(result)) = 0
    StringToUtf16Buffer = result
End Function

Private Function PtrToStringW(ByVal ptr As LongPtr) As String
    Dim length As Long

    If ptr = 0 Then
        PtrToStringW = vbNullString
        Exit Function
    End If

    length = lstrlenW(ptr)
    If length <= 0 Then
        PtrToStringW = vbNullString
        Exit Function
    End If

    PtrToStringW = String$(length, vbNullChar)
    CopyMemory ByVal StrPtr(PtrToStringW), ptr, CLng(length) * 2
End Function

Private Sub RangeToDoubleArray(ByVal source As Range, ByRef output() As Double)
    Dim values As Variant
    Dim i As Long
    Dim n As Long

    values = source.Value
    n = source.Count
    ReDim output(0 To n - 1)

    For i = 1 To n
        output(i - 1) = CDbl(values(i, 1))
    Next i
End Sub

Private Sub RangeToStringArray(ByVal source As Range, ByRef output() As String)
    Dim values As Variant
    Dim i As Long
    Dim n As Long

    values = source.Value
    n = source.Count
    ReDim output(0 To n - 1)

    For i = 1 To n
        output(i - 1) = CStr(values(i, 1))
    Next i
End Sub

Private Sub BuildUtf16StringPointers(ByRef strings() As String, _
                                     ByRef ptrs() As LongPtr)
    Dim i As Long
    Dim n As Long

    n = UBound(strings) - LBound(strings) + 1
    ReDim ptrs(0 To n - 1)

    For i = 0 To n - 1
        ptrs(i) = StrPtr(strings(i))
    Next i
End Sub

Private Function GetNamedRange(ByVal rangeName As String) As Range
    Dim target As Range

    On Error GoTo HandleError
    Set target = ThisWorkbook.Names(rangeName).RefersToRange
    Set GetNamedRange = target
    Exit Function

HandleError:
    Err.Raise vbObjectError + 1100, , "Missing named range: " & rangeName
End Function

Private Function TryGetNamedRange(ByVal rangeName As String) As Range
    If Len(rangeName) = 0 Then
        Exit Function
    End If

    On Error Resume Next
    Set TryGetNamedRange = ThisWorkbook.Names(rangeName).RefersToRange
    On Error GoTo 0
End Function

Private Sub LogMessage(ByVal message As String, Optional ByVal logSheet As Worksheet)
    Debug.Print message

    If Not logSheet Is Nothing Then
        Dim nextRow As Long

        If Application.WorksheetFunction.CountA(logSheet.Cells) = 0 Then
            nextRow = 1
        Else
            nextRow = logSheet.Cells(logSheet.Rows.Count, 1).End(xlUp).Row + 1
        End If

        logSheet.Cells(nextRow, 1).Value2 = Now
        logSheet.Cells(nextRow, 2).Value2 = message
    End If
End Sub

Private Function GetOrCreateLogSheet(ByVal logSheetName As String) As Worksheet
    Dim sheet As Worksheet

    On Error Resume Next
    Set sheet = ThisWorkbook.Worksheets(logSheetName)
    On Error GoTo 0

    If sheet Is Nothing Then
        Set sheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        sheet.Name = logSheetName
        sheet.Cells(1, 1).Value2 = "Timestamp"
        sheet.Cells(1, 2).Value2 = "Message"
    End If

    Set GetOrCreateLogSheet = sheet
End Function

Private Function HasHeaderRow(ByVal values As Variant, ByVal numericColumn As Long) As Boolean
    On Error GoTo NoHeader
    HasHeaderRow = Not IsNumeric(values(1, numericColumn))
    Exit Function

NoHeader:
    HasHeaderRow = False
End Function

Private Sub LoadCurvesFromTable(ByVal curveTable As Range, _
                                ByRef curves() As VBACurveInput, _
                                ByRef buffers() As CurveBuffers)
    Dim values As Variant
    Dim rowCount As Long
    Dim startRow As Long
    Dim r As Long
    Dim curveId As String
    Dim dict As Object
    Dim counts As Object
    Dim dayCounts As Object
    Dim key As Variant
    Dim idx As Long
    Dim positions() As Long

    If curveTable.Columns.Count < 5 Then
        Err.Raise vbObjectError + 1101, , "Curve table must have 5 columns: CurveId, PillarDate, Rate, Tenor, DayCount"
    End If

    values = curveTable.Value2
    rowCount = UBound(values, 1)
    startRow = IIf(HasHeaderRow(values, 2), 2, 1)

    Set dict = CreateObject("Scripting.Dictionary")
    Set counts = CreateObject("Scripting.Dictionary")
    Set dayCounts = CreateObject("Scripting.Dictionary")

    For r = startRow To rowCount
        curveId = Trim(CStr(values(r, 1)))
        If Len(curveId) = 0 Then
            Err.Raise vbObjectError + 1102, , "Curve table contains blank CurveId"
        End If

        If Not dict.Exists(curveId) Then
            dict.Add curveId, dict.Count
            counts.Add curveId, 0
            dayCounts.Add curveId, CLng(values(r, 5))
        ElseIf CLng(values(r, 5)) <> CLng(dayCounts(curveId)) Then
            Err.Raise vbObjectError + 1103, , "Curve " & curveId & " has inconsistent day count codes"
        End If

        counts(curveId) = counts(curveId) + 1
    Next r

    If dict.Count = 0 Then
        Err.Raise vbObjectError + 1104, , "Curve table is empty"
    End If

    ReDim curves(0 To dict.Count - 1)
    ReDim buffers(0 To dict.Count - 1)
    ReDim positions(0 To dict.Count - 1)

    For Each key In dict.Keys
        If CLng(counts(key)) <= 0 Then
            Err.Raise vbObjectError + 1116, , "Curve " & CStr(key) & " has no pillars"
        End If
        idx = dict(key)
        buffers(idx).curveId = CStr(key)
        ReDim buffers(idx).pillarSerials(0 To counts(key) - 1)
        ReDim buffers(idx).discountRates(0 To counts(key) - 1)
        ReDim buffers(idx).tenorStrings(0 To counts(key) - 1)
        ReDim buffers(idx).tenorPtrs(0 To counts(key) - 1)
    Next key

    For r = startRow To rowCount
        curveId = Trim(CStr(values(r, 1)))
        idx = dict(curveId)
        buffers(idx).pillarSerials(positions(idx)) = CDbl(values(r, 2))
        buffers(idx).discountRates(positions(idx)) = CDbl(values(r, 3))
        buffers(idx).tenorStrings(positions(idx)) = CStr(values(r, 4))
        positions(idx) = positions(idx) + 1
    Next r

    For Each key In dict.Keys
        idx = dict(key)
        BuildUtf16StringPointers buffers(idx).tenorStrings, buffers(idx).tenorPtrs
        curves(idx).id = StrPtr(buffers(idx).curveId)
        curves(idx).pillarSerials = VarPtr(buffers(idx).pillarSerials(0))
        curves(idx).discountRates = VarPtr(buffers(idx).discountRates(0))
        curves(idx).tenorStrings = VarPtr(buffers(idx).tenorPtrs(0))
        curves(idx).pillarCount = UBound(buffers(idx).pillarSerials) + 1
        curves(idx).dayCountCode = CLng(dayCounts(key))
    Next key
End Sub

Private Sub LoadFixingsFromTable(ByVal fixingTable As Range, _
                                 ByRef fixings() As VBAFixingInput, _
                                 ByRef buffers() As FixingBuffers)
    Dim values As Variant
    Dim rowCount As Long
    Dim startRow As Long
    Dim r As Long
    Dim indexName As String
    Dim dict As Object
    Dim counts As Object
    Dim key As Variant
    Dim idx As Long
    Dim positions() As Long

    If fixingTable.Columns.Count < 3 Then
        Err.Raise vbObjectError + 1105, , "Fixing table must have 3 columns: IndexName, FixingDate, FixingRate"
    End If

    values = fixingTable.Value2
    rowCount = UBound(values, 1)
    startRow = IIf(HasHeaderRow(values, 2), 2, 1)

    Set dict = CreateObject("Scripting.Dictionary")
    Set counts = CreateObject("Scripting.Dictionary")

    For r = startRow To rowCount
        indexName = Trim(CStr(values(r, 1)))
        If Len(indexName) = 0 Then
            Err.Raise vbObjectError + 1106, , "Fixing table contains blank IndexName"
        End If

        If Not dict.Exists(indexName) Then
            dict.Add indexName, dict.Count
            counts.Add indexName, 0
        End If

        counts(indexName) = counts(indexName) + 1
    Next r

    If dict.Count = 0 Then
        Err.Raise vbObjectError + 1114, , "Fixing table is empty"
    End If

    ReDim fixings(0 To dict.Count - 1)
    ReDim buffers(0 To dict.Count - 1)
    ReDim positions(0 To dict.Count - 1)

    For Each key In dict.Keys
        If CLng(counts(key)) <= 0 Then
            Err.Raise vbObjectError + 1117, , "Fixing index " & CStr(key) & " has no rows"
        End If
        idx = dict(key)
        buffers(idx).indexName = CStr(key)
        ReDim buffers(idx).fixingDates(0 To counts(key) - 1)
        ReDim buffers(idx).fixingRates(0 To counts(key) - 1)
    Next key

    For r = startRow To rowCount
        indexName = Trim(CStr(values(r, 1)))
        idx = dict(indexName)
        buffers(idx).fixingDates(positions(idx)) = CDbl(values(r, 2))
        buffers(idx).fixingRates(positions(idx)) = CDbl(values(r, 3))
        positions(idx) = positions(idx) + 1
    Next r

    For Each key In dict.Keys
        idx = dict(key)
        fixings(idx).indexName = StrPtr(buffers(idx).indexName)
        fixings(idx).fixingDateSerials = VarPtr(buffers(idx).fixingDates(0))
        fixings(idx).fixingRates = VarPtr(buffers(idx).fixingRates(0))
        fixings(idx).fixingCount = UBound(buffers(idx).fixingDates) + 1
    Next key
End Sub

Private Sub LoadBucketsFromTable(ByVal bucketTable As Range, _
                                 ByRef buckets() As VBABucketConfig, _
                                 ByRef buffers() As BucketBuffers)
    Dim values As Variant
    Dim rowCount As Long
    Dim startRow As Long
    Dim r As Long
    Dim curveId As String
    Dim dict As Object
    Dim counts As Object
    Dim bumpSizes As Object
    Dim key As Variant
    Dim idx As Long
    Dim positions() As Long

    If bucketTable.Columns.Count < 3 Then
        Err.Raise vbObjectError + 1107, , "Bucket table must have 3 columns: CurveId, Tenor, BumpSize"
    End If

    values = bucketTable.Value2
    rowCount = UBound(values, 1)
    startRow = IIf(HasHeaderRow(values, 3), 2, 1)

    Set dict = CreateObject("Scripting.Dictionary")
    Set counts = CreateObject("Scripting.Dictionary")
    Set bumpSizes = CreateObject("Scripting.Dictionary")

    For r = startRow To rowCount
        curveId = Trim(CStr(values(r, 1)))
        If Len(curveId) = 0 Then
            Err.Raise vbObjectError + 1108, , "Bucket table contains blank CurveId"
        End If

        If Not dict.Exists(curveId) Then
            dict.Add curveId, dict.Count
            counts.Add curveId, 0
            bumpSizes.Add curveId, CDbl(values(r, 3))
        ElseIf CDbl(values(r, 3)) <> CDbl(bumpSizes(curveId)) Then
            Err.Raise vbObjectError + 1109, , "Bucket curve " & curveId & " has inconsistent bump sizes"
        End If

        counts(curveId) = counts(curveId) + 1
    Next r

    If dict.Count = 0 Then
        Err.Raise vbObjectError + 1115, , "Bucket table is empty"
    End If

    ReDim buckets(0 To dict.Count - 1)
    ReDim buffers(0 To dict.Count - 1)
    ReDim positions(0 To dict.Count - 1)

    For Each key In dict.Keys
        If CLng(counts(key)) <= 0 Then
            Err.Raise vbObjectError + 1118, , "Bucket curve " & CStr(key) & " has no tenors"
        End If
        idx = dict(key)
        buffers(idx).curveId = CStr(key)
        ReDim buffers(idx).tenorStrings(0 To counts(key) - 1)
        ReDim buffers(idx).tenorPtrs(0 To counts(key) - 1)
    Next key

    For r = startRow To rowCount
        curveId = Trim(CStr(values(r, 1)))
        idx = dict(curveId)
        buffers(idx).tenorStrings(positions(idx)) = CStr(values(r, 2))
        positions(idx) = positions(idx) + 1
    Next r

    For Each key In dict.Keys
        idx = dict(key)
        BuildUtf16StringPointers buffers(idx).tenorStrings, buffers(idx).tenorPtrs
        buckets(idx).curveId = StrPtr(buffers(idx).curveId)
        buckets(idx).tenorStrings = VarPtr(buffers(idx).tenorPtrs(0))
        buckets(idx).tenorCount = UBound(buffers(idx).tenorStrings) + 1
        buckets(idx).bumpSize = CDbl(bumpSizes(key))
    Next key
End Sub

' ====== Loaders ======

Public Sub LoadCurveInputFromSheet(ByVal curveId As String, _
                                   ByVal pillarSerialsRange As Range, _
                                   ByVal discountRatesRange As Range, _
                                   ByVal tenorsRange As Range, _
                                   ByVal dayCountCode As Long, _
                                   ByRef curve As VBACurveInput, _
                                   ByRef buffers As CurveBuffers)
    If pillarSerialsRange.Count <> discountRatesRange.Count Then
        Err.Raise vbObjectError + 1000, , "Pillar and rate ranges must match"
    End If

    If pillarSerialsRange.Count <> tenorsRange.Count Then
        Err.Raise vbObjectError + 1001, , "Pillar and tenor ranges must match"
    End If

    RangeToDoubleArray pillarSerialsRange, buffers.pillarSerials
    RangeToDoubleArray discountRatesRange, buffers.discountRates
    RangeToStringArray tenorsRange, buffers.tenorStrings
    BuildUtf16StringPointers buffers.tenorStrings, buffers.tenorPtrs

    buffers.curveId = curveId
    curve.id = StrPtr(buffers.curveId)
    curve.pillarSerials = VarPtr(buffers.pillarSerials(0))
    curve.discountRates = VarPtr(buffers.discountRates(0))
    curve.tenorStrings = VarPtr(buffers.tenorPtrs(0))
    curve.pillarCount = pillarSerialsRange.Count
    curve.dayCountCode = dayCountCode
End Sub

Public Sub LoadFixingsFromSheet(ByVal indexName As String, _
                                ByVal fixingDatesRange As Range, _
                                ByVal fixingRatesRange As Range, _
                                ByRef fixing As VBAFixingInput, _
                                ByRef buffers As FixingBuffers)
    If fixingDatesRange.Count <> fixingRatesRange.Count Then
        Err.Raise vbObjectError + 1002, , "Fixing date and rate ranges must match"
    End If

    RangeToDoubleArray fixingDatesRange, buffers.fixingDates
    RangeToDoubleArray fixingRatesRange, buffers.fixingRates

    buffers.indexName = indexName
    fixing.indexName = StrPtr(buffers.indexName)
    fixing.fixingDateSerials = VarPtr(buffers.fixingDates(0))
    fixing.fixingRates = VarPtr(buffers.fixingRates(0))
    fixing.fixingCount = fixingDatesRange.Count
End Sub

Public Sub LoadBucketConfigFromSheet(ByVal curveId As String, _
                                     ByVal tenorRange As Range, _
                                     ByVal bumpSizeCell As Range, _
                                     ByRef bucket As VBABucketConfig, _
                                     ByRef buffers As BucketBuffers)
    RangeToStringArray tenorRange, buffers.tenorStrings
    BuildUtf16StringPointers buffers.tenorStrings, buffers.tenorPtrs

    buffers.curveId = curveId
    bucket.curveId = StrPtr(buffers.curveId)
    bucket.tenorStrings = VarPtr(buffers.tenorPtrs(0))
    bucket.tenorCount = tenorRange.Count
    bucket.bumpSize = CDbl(bumpSizeCell.Value)
End Sub

Public Sub LoadLegSpecFromRange(ByVal legRange As Range, _
                                ByRef leg As VBALegSpec, _
                                ByRef buffers As LegBuffers)
    Dim values As Variant

    If legRange.Columns.Count < 13 Then
        Err.Raise vbObjectError + 1003, , "Leg range must include 13 columns"
    End If

    values = legRange.Value

    leg.legType = CLng(values(1, 1))
    leg.payReceive = CLng(values(1, 2))
    leg.notional = CDbl(values(1, 3))
    leg.startDateSerial = CDbl(values(1, 4))
    leg.endDateSerial = CDbl(values(1, 5))
    leg.frequencyCode = CLng(values(1, 6))
    leg.dayCountCode = CLng(values(1, 7))
    leg.bdcCode = CLng(values(1, 8))
    leg.fixedRate = CDbl(values(1, 9))

    buffers.indexName = CStr(values(1, 10))
    If Len(buffers.indexName) > 0 Then
        leg.indexName = StrPtr(buffers.indexName)
    Else
        leg.indexName = 0
    End If

    leg.fixingDays = CLng(values(1, 11))
    leg.spread = CDbl(values(1, 12))
    leg.isCompounded = CLng(values(1, 13))
End Sub

Public Sub LoadSwapSpecFromSheet(ByVal swapType As Long, _
                                 ByVal discountCurveId As String, _
                                 ByVal valuationCurveId As String, _
                                 ByVal valuationDateCell As Range, _
                                 ByVal leg1Range As Range, _
                                 ByVal leg2Range As Range, _
                                 ByRef swap As VBASwapSpec, _
                                 ByRef buffers As SwapBuffers)
    buffers.discountCurveId = discountCurveId
    buffers.valuationCurveId = valuationCurveId

    swap.swapType = swapType
    swap.discountCurveId = StrPtr(buffers.discountCurveId)
    swap.valuationCurveId = StrPtr(buffers.valuationCurveId)
    swap.valuationDateSerial = CDbl(valuationDateCell.Value)

    LoadLegSpecFromRange leg1Range, swap.leg1, buffers.leg1
    LoadLegSpecFromRange leg2Range, swap.leg2, buffers.leg2
End Sub

Public Sub PriceSwapFromSheet(Optional ByVal curvesRangeName As String = "SwapCurves", _
                              Optional ByVal fixingsRangeName As String = "SwapFixings", _
                              Optional ByVal bucketsRangeName As String = "SwapBuckets", _
                              Optional ByVal swapSpecRangeName As String = "SwapSpec", _
                              Optional ByVal leg1RangeName As String = "SwapLeg1", _
                              Optional ByVal leg2RangeName As String = "SwapLeg2", _
                              Optional ByVal holidaysRangeName As String = "SwapHolidays", _
                              Optional ByVal npvOutputName As String = "SwapNPVOutput", _
                              Optional ByVal bucketOutputName As String = "SwapBucketOutput", _
                              Optional ByVal enableDebug As Boolean = False, _
                              Optional ByVal dryRun As Boolean = False, _
                              Optional ByVal logSheetName As String = "")
    Dim curveTable As Range
    Dim fixingTable As Range
    Dim bucketTable As Range
    Dim swapSpecRange As Range
    Dim leg1Range As Range
    Dim leg2Range As Range
    Dim holidaysRange As Range
    Dim npvOutput As Range
    Dim bucketOutput As Range
    Dim logSheet As Worksheet
    Dim swapSpec As VBASwapSpec
    Dim holidayCount As Long
    Dim holidaySerialsPtr As LongPtr
    Dim outPillarSerials() As Double
    Dim outDeltas() As Double
    Dim usedBuckets As Long
    Dim maxBuckets As Long
    Dim npv As Double
    Dim specValues As Variant
    Dim swapCount As Long
    Dim bucketRowsPerSwap As Long
    Dim leg1Row As Range
    Dim leg2Row As Range
    Dim npvCell As Range
    Dim bucketRange As Range
    Dim i As Long
    Dim expectedBucketRows As Long
    Dim bucketConfigIndex As Long
    Dim errorMessage As String

    If Len(logSheetName) > 0 Then
        Set logSheet = GetOrCreateLogSheet(logSheetName)
    End If

    Set curveTable = GetNamedRange(curvesRangeName)
    Set fixingTable = GetNamedRange(fixingsRangeName)
    Set bucketTable = GetNamedRange(bucketsRangeName)
    Set swapSpecRange = GetNamedRange(swapSpecRangeName)
    Set leg1Range = GetNamedRange(leg1RangeName)
    Set leg2Range = GetNamedRange(leg2RangeName)
    Set npvOutput = GetNamedRange(npvOutputName)
    Set bucketOutput = GetNamedRange(bucketOutputName)
    Set holidaysRange = TryGetNamedRange(holidaysRangeName)

    If bucketOutput.Columns.Count < 2 Then
        Err.Raise vbObjectError + 1110, , "Bucket output range must have at least 2 columns (PillarDate, Delta)"
    End If

    LoadCurvesFromTable curveTable, mCurves, mCurveBuffers
    LoadFixingsFromTable fixingTable, mFixings, mFixingBuffers
    LoadBucketsFromTable bucketTable, mBuckets, mBucketBuffers
    ' Do not modify mCurveBuffers/mFixingBuffers/mBucketBuffers after string pointers are built until DLL calls complete.

    If swapSpecRange.Columns.Count < 4 Or swapSpecRange.Rows.Count < 1 Then
        Err.Raise vbObjectError + 1111, , "Swap spec range must have 4 columns: SwapType, DiscountCurveId, ValuationCurveId, ValuationDate"
    End If

    swapCount = swapSpecRange.Rows.Count
    If leg1Range.Rows.Count < swapCount Then
        Err.Raise vbObjectError + 1117, , "SwapLeg1 range must have at least " & swapCount & " rows to match SwapSpec"
    End If
    If leg2Range.Rows.Count < swapCount Then
        Err.Raise vbObjectError + 1118, , "SwapLeg2 range must have at least " & swapCount & " rows to match SwapSpec"
    End If
    If npvOutput.Rows.Count < swapCount Then
        Err.Raise vbObjectError + 1119, , "SwapNPVOutput range must have at least " & swapCount & " rows to match SwapSpec"
    End If
    If bucketOutput.Rows.Count Mod swapCount <> 0 Then
        Err.Raise vbObjectError + 1120, , "SwapBucketOutput rows must be divisible by SwapSpec row count"
    End If
    bucketRowsPerSwap = bucketOutput.Rows.Count \ swapCount
    If bucketRowsPerSwap < 1 Then
        Err.Raise vbObjectError + 1121, , "SwapBucketOutput range must have at least one row per swap"
    End If
    expectedBucketRows = 0
    For bucketConfigIndex = 0 To UBound(mBuckets)
        expectedBucketRows = expectedBucketRows + mBuckets(bucketConfigIndex).tenorCount
    Next bucketConfigIndex
    If bucketRowsPerSwap < expectedBucketRows Then
        Err.Raise vbObjectError + 1122, , "SwapBucketOutput range must have at least " & expectedBucketRows & " rows per swap to hold bucket deltas"
    End If

    specValues = swapSpecRange.Value2

    holidayCount = 0
    If Not holidaysRange Is Nothing Then
        If holidaysRange.Count > 0 Then
            If holidaysRange.Columns.Count <> 1 Then
                Err.Raise vbObjectError + 1112, , "Holiday range must be a single column"
            End If
            RangeToDoubleArray holidaysRange, mHolidaySerials
            holidayCount = holidaysRange.Count
        End If
    End If
    If holidayCount > 0 Then
        holidaySerialsPtr = VarPtr(mHolidaySerials(0))
    Else
        holidaySerialsPtr = 0
    End If

    If enableDebug Then
        SetDebugMode True
    End If

    LogMessage "PriceSwapFromSheet start", logSheet
    LogMessage "Curves: " & UBound(mCurves) + 1 & ", Fixings: " & UBound(mFixings) + 1 & ", Buckets: " & UBound(mBuckets) + 1, logSheet
    LogMessage "SwapCount=" & swapCount, logSheet
    LogMessage "HolidayCount=" & holidayCount, logSheet

    If dryRun Then
        LogMessage "Dry-run enabled; skipping DLL call.", logSheet
        Exit Sub
    End If

    maxBuckets = bucketRowsPerSwap
    If maxBuckets < 1 Then
        Err.Raise vbObjectError + 1116, , "Bucket output range must have at least one row"
    End If
    ReDim outPillarSerials(0 To maxBuckets - 1)
    ReDim outDeltas(0 To maxBuckets - 1)

    For i = 1 To swapCount
        Set leg1Row = leg1Range.Rows(i)
        Set leg2Row = leg2Range.Rows(i)
        Set npvCell = npvOutput.Cells(i, 1)
        Set bucketRange = bucketOutput.Rows((i - 1) * maxBuckets + 1).Resize(maxBuckets, 2)

        LoadSwapSpecFromSheet CLng(specValues(i, 1)), _
                              CStr(specValues(i, 2)), _
                              CStr(specValues(i, 3)), _
                              swapSpecRange.Cells(i, 4), _
                              leg1Row, leg2Row, _
                              swapSpec, mSwapBuffers

        LogMessage "SwapRow=" & i & ", SwapType=" & specValues(i, 1) & ", DiscountCurveId=" & specValues(i, 2) & ", ValuationCurveId=" & specValues(i, 3) & ", ValuationDate=" & specValues(i, 4), logSheet

        npv = PriceSwapAndBucketsPtr(swapSpec, mCurves(0), UBound(mCurves) + 1, mFixings(0), UBound(mFixings) + 1, _
                                     mBuckets(0), UBound(mBuckets) + 1, holidaySerialsPtr, holidayCount, _
                                     outPillarSerials(0), outDeltas(0), maxBuckets, usedBuckets)

        npvCell.Value2 = npv

        If usedBuckets > maxBuckets Then
            Err.Raise vbObjectError + 1113, , "Bucket output range too small for returned deltas"
        End If

        Dim bucketValues() As Variant
        Dim bucketIndex As Long
        ReDim bucketValues(1 To maxBuckets, 1 To 2)
        For bucketIndex = 1 To maxBuckets
            If bucketIndex <= usedBuckets Then
                bucketValues(bucketIndex, 1) = outPillarSerials(bucketIndex - 1)
                bucketValues(bucketIndex, 2) = outDeltas(bucketIndex - 1)
            Else
                bucketValues(bucketIndex, 1) = vbNullString
                bucketValues(bucketIndex, 2) = vbNullString
            End If
        Next bucketIndex
        bucketRange.Value2 = bucketValues
    Next i

    errorMessage = PtrToStringW(IRS_LAST_ERROR())
    If Len(errorMessage) > 0 Then
        LogMessage "IRS_LAST_ERROR: " & errorMessage, logSheet
    End If

    LogMessage "PriceSwapFromSheet completed. UsedBuckets=" & usedBuckets & ", LastNPV=" & npv, logSheet
End Sub

Public Function PriceSwapAndBuckets(ByRef swapSpec As VBASwapSpec, _
                                    ByRef curveInputs As VBACurveInput, ByVal curveCount As Long, _
                                    ByRef fixingInputs As VBAFixingInput, ByVal fixingCount As Long, _
                                    ByRef bucketInputs As VBABucketConfig, ByVal bucketCount As Long, _
                                    ByRef holidaySerials As Double, ByVal holidayCount As Long, _
                                    ByRef outPillarSerials As Double, ByRef outDeltas As Double, _
                                    ByVal maxBuckets As Long, ByRef outUsedBuckets As Long) As Double
    Dim result As Double
    Dim errorPtr As LongPtr
    Dim errorMessage As String

    result = IRS_PRICE_AND_BUCKETS(swapSpec, curveInputs, curveCount, fixingInputs, fixingCount, _
                                   bucketInputs, bucketCount, holidaySerials, holidayCount, _
                                   outPillarSerials, outDeltas, maxBuckets, outUsedBuckets)

    errorPtr = IRS_LAST_ERROR()
    errorMessage = PtrToStringW(errorPtr)
    If Len(errorMessage) > 0 Then
        Err.Raise vbObjectError + 2000, , errorMessage
    End If

    If IRS_IS_NAN(result) <> 0 Then
        Err.Raise vbObjectError + 2000, , "IRS_PRICE_AND_BUCKETS returned NaN without error"
    End If

    PriceSwapAndBuckets = result
End Function

Public Function PriceSwapAndBucketsPtr(ByRef swapSpec As VBASwapSpec, _
                                       ByRef curveInputs As VBACurveInput, ByVal curveCount As Long, _
                                       ByRef fixingInputs As VBAFixingInput, ByVal fixingCount As Long, _
                                       ByRef bucketInputs As VBABucketConfig, ByVal bucketCount As Long, _
                                       ByVal holidaySerialsPtr As LongPtr, ByVal holidayCount As Long, _
                                       ByRef outPillarSerials As Double, ByRef outDeltas As Double, _
                                       ByVal maxBuckets As Long, ByRef outUsedBuckets As Long) As Double
    Dim result As Double
    Dim errorPtr As LongPtr
    Dim errorMessage As String

    result = IRS_PRICE_AND_BUCKETS_PTR(swapSpec, curveInputs, curveCount, fixingInputs, fixingCount, _
                                       bucketInputs, bucketCount, holidaySerialsPtr, holidayCount, _
                                       outPillarSerials, outDeltas, maxBuckets, outUsedBuckets)

    errorPtr = IRS_LAST_ERROR()
    errorMessage = PtrToStringW(errorPtr)
    If Len(errorMessage) > 0 Then
        Err.Raise vbObjectError + 2000, , errorMessage
    End If

    If IRS_IS_NAN(result) <> 0 Then
        Err.Raise vbObjectError + 2000, , "IRS_PRICE_AND_BUCKETS returned NaN without error"
    End If

    PriceSwapAndBucketsPtr = result
End Function

Public Sub SetDebugMode(ByVal enabled As Boolean, Optional ByVal logPath As String = "")
    Dim logPathUtf16() As Byte

    IRS_SET_DEBUG_MODE IIf(enabled, 1, 0)
    If Len(logPath) > 0 Then
        logPathUtf16 = StringToUtf16Buffer(logPath)
        IRS_SET_DEBUG_LOG_PATH VarPtr(logPathUtf16(0))
    End If
End Sub
