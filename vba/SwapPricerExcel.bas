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
    pillarSerials() As Double
    discountRates() As Double
    tenorStrings() As String
    tenorPtrs() As LongPtr
End Type

Public Type FixingBuffers
    fixingDates() As Double
    fixingRates() As Double
End Type

Public Type BucketBuffers
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
    Public Declare PtrSafe Function IRS_LAST_ERROR Lib SWAP_PRICER_DLL () As LongPtr
    Public Declare PtrSafe Sub IRS_SET_DEBUG_MODE Lib SWAP_PRICER_DLL (ByVal enabled As Long)
    Public Declare PtrSafe Sub IRS_SET_DEBUG_LOG_PATH Lib SWAP_PRICER_DLL (ByVal path As LongPtr)
    Private Declare PtrSafe Function lstrlenA Lib "kernel32" (ByVal lpString As LongPtr) As Long
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
    Public Declare Function IRS_LAST_ERROR Lib SWAP_PRICER_DLL () As LongPtr
    Public Declare Sub IRS_SET_DEBUG_MODE Lib SWAP_PRICER_DLL (ByVal enabled As Long)
    Public Declare Sub IRS_SET_DEBUG_LOG_PATH Lib SWAP_PRICER_DLL (ByVal path As LongPtr)
    Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As LongPtr) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByVal dest As LongPtr, ByVal src As LongPtr, ByVal cb As LongPtr)
#End If

' ====== Helpers ======

Private Function PtrToCString(ByVal s As String) As LongPtr
    PtrToCString = StrPtr(s)
End Function

Private Function PtrToStringA(ByVal ptr As LongPtr) As String
    Dim length As Long
    Dim bytes() As Byte

    If ptr = 0 Then
        PtrToStringA = vbNullString
        Exit Function
    End If

    length = lstrlenA(ptr)
    If length <= 0 Then
        PtrToStringA = vbNullString
        Exit Function
    End If

    ReDim bytes(0 To length - 1)
    CopyMemory VarPtr(bytes(0)), ptr, length
    PtrToStringA = StrConv(bytes, vbUnicode)
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

Private Sub BuildStringPointers(ByRef strings() As String, ByRef ptrs() As LongPtr)
    Dim i As Long
    Dim n As Long

    n = UBound(strings) - LBound(strings) + 1
    ReDim ptrs(0 To n - 1)

    For i = 0 To n - 1
        ptrs(i) = PtrToCString(strings(i))
    Next i
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
    BuildStringPointers buffers.tenorStrings, buffers.tenorPtrs

    curve.id = PtrToCString(curveId)
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

    fixing.indexName = PtrToCString(indexName)
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
    BuildStringPointers buffers.tenorStrings, buffers.tenorPtrs

    bucket.curveId = PtrToCString(curveId)
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
        leg.indexName = PtrToCString(buffers.indexName)
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
    swap.discountCurveId = PtrToCString(buffers.discountCurveId)
    swap.valuationCurveId = PtrToCString(buffers.valuationCurveId)
    swap.valuationDateSerial = CDbl(valuationDateCell.Value)

    LoadLegSpecFromRange leg1Range, swap.leg1, buffers.leg1
    LoadLegSpecFromRange leg2Range, swap.leg2, buffers.leg2
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

    If result <> result Then
        errorPtr = IRS_LAST_ERROR()
        errorMessage = PtrToStringA(errorPtr)
        If Len(errorMessage) = 0 Then
            errorMessage = "Unknown error from IRS_PRICE_AND_BUCKETS"
        End If
        Err.Raise vbObjectError + 2000, , errorMessage
    End If

    PriceSwapAndBuckets = result
End Function

Public Sub SetDebugMode(ByVal enabled As Boolean, Optional ByVal logPath As String = "")
    IRS_SET_DEBUG_MODE IIf(enabled, 1, 0)
    If Len(logPath) > 0 Then
        IRS_SET_DEBUG_LOG_PATH PtrToCString(logPath)
    End If
End Sub
