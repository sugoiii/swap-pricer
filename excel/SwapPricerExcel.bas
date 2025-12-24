Option Explicit

' VBA helpers for swap-pricer Excel DLL.
' These functions load curve/swap inputs from worksheet ranges
' and pin their memory so the C++ DLL can read the data safely.

#If VBA7 Then
    Private Const DLL_NAME As String = "C:\\path\\to\\swap_pricer.dll"
#Else
    Private Const DLL_NAME As String = "C:\\path\\to\\swap_pricer.dll"
#End If

' ====== Struct definitions (match src/Excel.cpp) ======
#If VBA7 Then
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
#Else
    Public Type VBACurveInput
        id As Long
        pillarSerials As Long
        discountRates As Long
        tenorStrings As Long
        pillarCount As Long
        dayCountCode As Long
    End Type

    Public Type VBAFixingInput
        indexName As Long
        fixingDateSerials As Long
        fixingRates As Long
        fixingCount As Long
    End Type

    Public Type VBABucketConfig
        curveId As Long
        tenorStrings As Long
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
        indexName As Long
        fixingDays As Long
        spread As Double
        isCompounded As Long
    End Type

    Public Type VBASwapSpec
        swapType As Long
        leg1 As VBALegSpec
        leg2 As VBALegSpec
        discountCurveId As Long
        valuationCurveId As Long
        valuationDateSerial As Double
    End Type
#End If

#If VBA7 Then
    Private Declare PtrSafe Function IRS_PRICE_AND_BUCKETS Lib DLL_NAME _
        (ByRef swapSpec As VBASwapSpec, _
         ByRef curveInputs As VBACurveInput, ByVal curveCount As Long, _
         ByRef fixingInputs As VBAFixingInput, ByVal fixingCount As Long, _
         ByRef bucketInputs As VBABucketConfig, ByVal bucketCount As Long, _
         ByRef holidaySerials As Double, ByVal holidayCount As Long, _
         ByRef outPillarSerials As Double, ByRef outDeltas As Double, _
         ByVal maxBuckets As Long, ByRef outUsedBuckets As Long) As Double
#Else
    Private Declare Function IRS_PRICE_AND_BUCKETS Lib DLL_NAME _
        (ByRef swapSpec As VBASwapSpec, _
         ByRef curveInputs As VBACurveInput, ByVal curveCount As Long, _
         ByRef fixingInputs As VBAFixingInput, ByVal fixingCount As Long, _
         ByRef bucketInputs As VBABucketConfig, ByVal bucketCount As Long, _
         ByRef holidaySerials As Double, ByVal holidayCount As Long, _
         ByRef outPillarSerials As Double, ByRef outDeltas As Double, _
         ByVal maxBuckets As Long, ByRef outUsedBuckets As Long) As Double
#End If

' ====== Memory pinning ======
Private gPinnedArrays As Collection
Private gAnsiStrings As Collection

Public Sub ResetPinnedBuffers()
    Set gPinnedArrays = New Collection
    Set gAnsiStrings = New Collection
End Sub

Private Sub EnsurePinned()
    If gPinnedArrays Is Nothing Then Set gPinnedArrays = New Collection
    If gAnsiStrings Is Nothing Then Set gAnsiStrings = New Collection
End Sub

#If VBA7 Then
    Private Function PinDoubleArray(ByRef values() As Double) As LongPtr
        EnsurePinned
        gPinnedArrays.Add values
        PinDoubleArray = VarPtr(values(LBound(values)))
    End Function

    Private Function PinLongPtrArray(ByRef values() As LongPtr) As LongPtr
        EnsurePinned
        gPinnedArrays.Add values
        PinLongPtrArray = VarPtr(values(LBound(values)))
    End Function

    Private Function PtrToAnsiString(ByVal value As String) As LongPtr
        EnsurePinned
        Dim bytes() As Byte
        bytes = StrConv(value & vbNullChar, vbFromUnicode)
        gAnsiStrings.Add bytes
        PtrToAnsiString = VarPtr(bytes(0))
    End Function
#Else
    Private Function PinDoubleArray(ByRef values() As Double) As Long
        EnsurePinned
        gPinnedArrays.Add values
        PinDoubleArray = VarPtr(values(LBound(values)))
    End Function

    Private Function PinLongPtrArray(ByRef values() As Long) As Long
        EnsurePinned
        gPinnedArrays.Add values
        PinLongPtrArray = VarPtr(values(LBound(values)))
    End Function

    Private Function PtrToAnsiString(ByVal value As String) As Long
        EnsurePinned
        Dim bytes() As Byte
        bytes = StrConv(value & vbNullChar, vbFromUnicode)
        gAnsiStrings.Add bytes
        PtrToAnsiString = VarPtr(bytes(0))
    End Function
#End If

' ====== Range helpers ======
Private Function RangeToDoubleArray(ByVal rng As Range) As Double()
    Dim result() As Double
    Dim count As Long
    count = rng.Count
    ReDim result(0 To count - 1)

    Dim i As Long
    Dim cell As Range
    i = 0
    For Each cell In rng.Cells
        result(i) = CDbl(cell.Value2)
        i = i + 1
    Next cell

    RangeToDoubleArray = result
End Function

Private Function RangeToStringArray(ByVal rng As Range) As String()
    Dim result() As String
    Dim count As Long
    count = rng.Count
    ReDim result(0 To count - 1)

    Dim i As Long
    Dim cell As Range
    i = 0
    For Each cell In rng.Cells
        result(i) = CStr(cell.Value2)
        i = i + 1
    Next cell

    RangeToStringArray = result
End Function

Private Function RangeToScalar(ByVal rng As Range) As Variant
    RangeToScalar = rng.Cells(1, 1).Value2
End Function

' ====== Loaders ======
Public Function LoadCurveInputFromSheet(ByVal curveIdCell As Range, _
                                       ByVal pillarRange As Range, _
                                       ByVal rateRange As Range, _
                                       ByVal tenorRange As Range, _
                                       ByVal dayCountCell As Range) As VBACurveInput
    Dim curve As VBACurveInput
    Dim curveId As String
    curveId = CStr(RangeToScalar(curveIdCell))

    Dim pillars() As Double
    Dim rates() As Double
    Dim tenors() As String

    pillars = RangeToDoubleArray(pillarRange)
    rates = RangeToDoubleArray(rateRange)
    tenors = RangeToStringArray(tenorRange)

    If UBound(pillars) <> UBound(rates) Or UBound(pillars) <> UBound(tenors) Then
        Err.Raise vbObjectError + 100, "LoadCurveInputFromSheet", "Curve ranges have mismatched lengths."
    End If

    #If VBA7 Then
        Dim tenorPtrs() As LongPtr
    #Else
        Dim tenorPtrs() As Long
    #End If
    ReDim tenorPtrs(0 To UBound(tenors))

    Dim i As Long
    For i = 0 To UBound(tenors)
        tenorPtrs(i) = PtrToAnsiString(tenors(i))
    Next i

    curve.id = PtrToAnsiString(curveId)
    curve.pillarSerials = PinDoubleArray(pillars)
    curve.discountRates = PinDoubleArray(rates)
    curve.tenorStrings = PinLongPtrArray(tenorPtrs)
    curve.pillarCount = UBound(pillars) + 1
    curve.dayCountCode = CLng(RangeToScalar(dayCountCell))

    LoadCurveInputFromSheet = curve
End Function

Public Function LoadFixingInputFromSheet(ByVal indexNameCell As Range, _
                                         ByVal fixingDateRange As Range, _
                                         ByVal fixingRateRange As Range) As VBAFixingInput
    Dim fix As VBAFixingInput
    Dim indexName As String
    indexName = CStr(RangeToScalar(indexNameCell))

    If Len(indexName) = 0 Then
        fix.indexName = 0
        fix.fixingCount = 0
        LoadFixingInputFromSheet = fix
        Exit Function
    End If

    Dim fixingDates() As Double
    Dim fixingRates() As Double

    fixingDates = RangeToDoubleArray(fixingDateRange)
    fixingRates = RangeToDoubleArray(fixingRateRange)

    If UBound(fixingDates) <> UBound(fixingRates) Then
        Err.Raise vbObjectError + 101, "LoadFixingInputFromSheet", "Fixing ranges have mismatched lengths."
    End If

    fix.indexName = PtrToAnsiString(indexName)
    fix.fixingDateSerials = PinDoubleArray(fixingDates)
    fix.fixingRates = PinDoubleArray(fixingRates)
    fix.fixingCount = UBound(fixingDates) + 1

    LoadFixingInputFromSheet = fix
End Function

Public Function LoadBucketConfigFromSheet(ByVal curveIdCell As Range, _
                                          ByVal tenorRange As Range, _
                                          ByVal bumpCell As Range) As VBABucketConfig
    Dim cfg As VBABucketConfig
    Dim curveId As String
    curveId = CStr(RangeToScalar(curveIdCell))

    If Len(curveId) = 0 Then
        cfg.curveId = 0
        cfg.tenorCount = 0
        cfg.bumpSize = 0#
        LoadBucketConfigFromSheet = cfg
        Exit Function
    End If

    Dim tenors() As String
    tenors = RangeToStringArray(tenorRange)

    #If VBA7 Then
        Dim tenorPtrs() As LongPtr
    #Else
        Dim tenorPtrs() As Long
    #End If
    ReDim tenorPtrs(0 To UBound(tenors))

    Dim i As Long
    For i = 0 To UBound(tenors)
        tenorPtrs(i) = PtrToAnsiString(tenors(i))
    Next i

    cfg.curveId = PtrToAnsiString(curveId)
    cfg.tenorStrings = PinLongPtrArray(tenorPtrs)
    cfg.tenorCount = UBound(tenors) + 1
    cfg.bumpSize = CDbl(RangeToScalar(bumpCell))

    LoadBucketConfigFromSheet = cfg
End Function

Public Function LoadLegSpecFromRange(ByVal specRange As Range) As VBALegSpec
    Dim values() As Variant
    Dim count As Long
    count = specRange.Count

    If count < 13 Then
        Err.Raise vbObjectError + 102, "LoadLegSpecFromRange", "Leg spec range must have 13 values."
    End If

    ReDim values(0 To count - 1)

    Dim i As Long
    Dim cell As Range
    i = 0
    For Each cell In specRange.Cells
        values(i) = cell.Value2
        i = i + 1
    Next cell

    Dim leg As VBALegSpec
    leg.legType = CLng(values(0))
    leg.payReceive = CLng(values(1))
    leg.notional = CDbl(values(2))
    leg.startDateSerial = CDbl(values(3))
    leg.endDateSerial = CDbl(values(4))
    leg.frequencyCode = CLng(values(5))
    leg.dayCountCode = CLng(values(6))
    leg.bdcCode = CLng(values(7))
    leg.fixedRate = CDbl(values(8))

    If Len(CStr(values(9))) > 0 Then
        leg.indexName = PtrToAnsiString(CStr(values(9)))
    Else
        leg.indexName = 0
    End If

    leg.fixingDays = CLng(values(10))
    leg.spread = CDbl(values(11))
    leg.isCompounded = CLng(values(12))

    LoadLegSpecFromRange = leg
End Function

Public Function LoadSwapSpecFromRange(ByVal headerRange As Range, _
                                      ByVal leg1Range As Range, _
                                      ByVal leg2Range As Range) As VBASwapSpec
    Dim headerValues() As Variant
    Dim count As Long
    count = headerRange.Count

    If count < 4 Then
        Err.Raise vbObjectError + 103, "LoadSwapSpecFromRange", "Swap header range must have 4 values."
    End If

    ReDim headerValues(0 To count - 1)

    Dim i As Long
    Dim cell As Range
    i = 0
    For Each cell In headerRange.Cells
        headerValues(i) = cell.Value2
        i = i + 1
    Next cell

    Dim swap As VBASwapSpec
    swap.swapType = CLng(headerValues(0))
    swap.discountCurveId = PtrToAnsiString(CStr(headerValues(1)))
    swap.valuationCurveId = PtrToAnsiString(CStr(headerValues(2)))
    swap.valuationDateSerial = CDbl(headerValues(3))
    swap.leg1 = LoadLegSpecFromRange(leg1Range)
    swap.leg2 = LoadLegSpecFromRange(leg2Range)

    LoadSwapSpecFromRange = swap
End Function

Public Function PriceSwapFromSheet(ByVal headerRange As Range, _
                                   ByVal leg1Range As Range, _
                                   ByVal leg2Range As Range, _
                                   ByVal curveIdCell As Range, _
                                   ByVal pillarRange As Range, _
                                   ByVal rateRange As Range, _
                                   ByVal tenorRange As Range, _
                                   ByVal dayCountCell As Range) As Double
    ResetPinnedBuffers

    Dim swap As VBASwapSpec
    swap = LoadSwapSpecFromRange(headerRange, leg1Range, leg2Range)

    Dim curve As VBACurveInput
    curve = LoadCurveInputFromSheet(curveIdCell, pillarRange, rateRange, tenorRange, dayCountCell)

    Dim fix As VBAFixingInput
    fix.indexName = 0
    fix.fixingCount = 0

    Dim bucket As VBABucketConfig
    bucket.curveId = 0
    bucket.tenorCount = 0
    bucket.bumpSize = 0#

    Dim outPillars(0 To 9) As Double
    Dim outDeltas(0 To 9) As Double
    Dim usedBuckets As Long

    PriceSwapFromSheet = IRS_PRICE_AND_BUCKETS( _
        swap, _
        curve, 1, _
        fix, 0, _
        bucket, 0, _
        outPillars(0), 0, _
        outPillars(0), outDeltas(0), 10, usedBuckets)
End Function
