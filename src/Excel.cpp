// ExcelBridge.cpp
//
// Windows-only Excel-DLL bridge for IRS pricing + bucketed delta.
// On non-Windows, this file compiles to nothing useful (no exports).

#include <string>
#include <vector>
#include <limits>

#include <ql/quantlib.hpp>

#include "SwapTypes.hpp"
#include "SwapSpec.hpp"
#include "MarketData.hpp"
#include "Pricer.hpp"

using namespace IRS;
using namespace QuantLib;

// ---------- Helpers shared for both test & Excel ----------

// This assumes you're using QuantLib serials directly;
// adapt if you really need "true" Excel serials later.
static Date fromExcelSerial(double serial)
{
    long s = static_cast<long>(serial);
    return Date(s);
}

static double toExcelSerial(const Date &d)
{
    return static_cast<double>(d.serialNumber());
}

static Period parseTenorString(const std::string &s)
{
    return PeriodParser::parse(s);
}

static std::vector<Date> toHolidayDates(const double *holidaySerials, int holidayCount)
{
    std::vector<Date> holidays;
    if (!holidaySerials || holidayCount <= 0)
    {
        return holidays;
    }

    holidays.reserve(static_cast<std::size_t>(holidayCount));
    for (int i = 0; i < holidayCount; ++i)
    {
        holidays.push_back(fromExcelSerial(holidaySerials[i]));
    }

    return holidays;
}

#ifdef _WIN32

#include <windows.h>

// ---------------- VBA-facing POD definitions ----------------
// VBA should marshal arrays using SAFEARRAY/ByRef pointers and fill these
// plain-old-data structs. All string pointers are expected to remain alive
// for the duration of the call (BSTR or char*). Unless otherwise noted, the
// "count" fields represent the number of elements in the corresponding array.
//
// VBACurveInput:
//   - id: curve identifier used in swap specs (e.g., "DISCOUNT", "FWD_KOFR").
//   - pillarSerials/discountRates/tenorStrings: aligned arrays of equal length
//     describing the curve pillars and their zero/discount rates.
//   - dayCountCode: 0 = Actual365Fixed, 1 = Actual360.
//
// VBAFixingInput:
//   - indexName: e.g., "KOFR" or "CD".
//   - fixingDateSerials/fixingRates: aligned arrays of historical fixings.
//
// VBABucketConfig:
//   - curveId: which curve to bump.
//   - tenorStrings: e.g., {"1M","3M","6M","1Y"}.
//   - bumpSize: absolute bump applied to zero rates.
//
// VBALegSpec + VBASwapSpec:
//   - swapType: 0 = Vanilla (fixed/ibor), 1 = OvernightIndexed (fixed/ON).
//   - leg type: 0 = Fixed, 1 = Ibor, 2 = Overnight.
//   - payReceive: 0 = Payer, 1 = Receiver.
//   - frequency: 0 = Annual, 1 = SemiAnnual, 2 = Quarterly, 3 = Monthly.
//   - dayCountCode: 0 = Actual365Fixed, 1 = Actual360.
//   - bdcCode: 0 = Following, 1 = ModifiedFollowing, 2 = Preceding.
//   - Floating legs must provide indexName; spreads are in decimal.
//   - Dates are QuantLib serials (compatible with Excel when using the
//     1900-date system). Bucketed deltas are returned as parallel arrays of
//     pillar serials and deltas per swap.

struct VBACurveInput
{
    const char *id;
    const double *pillarSerials;
    const double *discountRates;
    const char **tenorStrings;
    int pillarCount;
    int dayCountCode;
};

struct VBAFixingInput
{
    const char *indexName;
    const double *fixingDateSerials;
    const double *fixingRates;
    int fixingCount;
};

struct VBABucketConfig
{
    const char *curveId;
    const char **tenorStrings;
    int tenorCount;
    double bumpSize;
};

struct VBALegSpec
{
    int legType;
    int payReceive;
    double notional;
    double startDateSerial;
    double endDateSerial;
    int frequencyCode;
    int dayCountCode;
    int bdcCode;
    double fixedRate;
    const char *indexName;
    int fixingDays;
    double spread;
    int isCompounded;
};

struct VBASwapSpec
{
    int swapType;
    VBALegSpec leg1;
    VBALegSpec leg2;
    const char *discountCurveId;
    const char *valuationCurveId;
    double valuationDateSerial;
};

// ---------------- Mapping helpers ----------------

static bool mapCurveDayCount(int code, DayCounter &out)
{
    switch (code)
    {
    case 0:
        out = Actual365Fixed();
        return true;
    case 1:
        out = Actual360();
        return true;
    default:
        return false;
    }
}

static bool mapLegDayCount(int code, IRS::DayCount &out)
{
    switch (code)
    {
    case 0:
        out = IRS::DayCount::Actual365Fixed;
        return true;
    case 1:
        out = IRS::DayCount::Actual360;
        return true;
    default:
        return false;
    }
}

static bool mapBDC(int code, IRS::BusinessDayConv &out)
{
    switch (code)
    {
    case 0:
        out = IRS::BusinessDayConv::Following;
        return true;
    case 1:
        out = IRS::BusinessDayConv::ModifiedFollowing;
        return true;
    case 2:
        out = IRS::BusinessDayConv::Preceding;
        return true;
    default:
        return false;
    }
}

static bool mapFrequency(int code, IRS::Frequency &out)
{
    switch (code)
    {
    case 0:
        out = IRS::Frequency::Annual;
        return true;
    case 1:
        out = IRS::Frequency::SemiAnnual;
        return true;
    case 2:
        out = IRS::Frequency::Quarterly;
        return true;
    case 3:
        out = IRS::Frequency::Monthly;
        return true;
    default:
        return false;
    }
}

static bool mapLegType(int code, IRS::LegType &out)
{
    switch (code)
    {
    case 0:
        out = IRS::LegType::Fixed;
        return true;
    case 1:
        out = IRS::LegType::Ibor;
        return true;
    case 2:
        out = IRS::LegType::Overnight;
        return true;
    default:
        return false;
    }
}

static bool mapPayReceive(int code, IRS::PayReceive &out)
{
    switch (code)
    {
    case 0:
        out = IRS::PayReceive::Payer;
        return true;
    case 1:
        out = IRS::PayReceive::Receiver;
        return true;
    default:
        return false;
    }
}

static bool mapSwapType(int code, IRS::SwapType &out)
{
    switch (code)
    {
    case 0:
        out = IRS::SwapType::Vanilla;
        return true;
    case 1:
        out = IRS::SwapType::OvernightIndexed;
        return true;
    default:
        return false;
    }
}

static bool fillCurveInput(const VBACurveInput &raw,
                           const Calendar &calendar,
                           CurveInput &out,
                           std::string &error)
{
    if (!raw.id)
    {
        error = "Curve id is null";
        return false;
    }

    if (raw.pillarCount <= 0 || !raw.pillarSerials || !raw.discountRates)
    {
        error = "Curve pillars/rates are missing";
        return false;
    }

    if (!raw.tenorStrings)
    {
        error = "Curve tenors are missing";
        return false;
    }

    DayCounter dc;
    if (!mapCurveDayCount(raw.dayCountCode, dc))
    {
        error = "Unsupported curve day count";
        return false;
    }

    out.id = raw.id;
    out.dayCounter = dc;
    out.dates.clear();
    out.discountRates.clear();
    out.tenors.clear();

    out.dates.reserve(static_cast<std::size_t>(raw.pillarCount));
    out.discountRates.reserve(static_cast<std::size_t>(raw.pillarCount));
    out.tenors.reserve(static_cast<std::size_t>(raw.pillarCount));

    for (int i = 0; i < raw.pillarCount; ++i)
    {
        out.dates.push_back(calendar.adjust(fromExcelSerial(raw.pillarSerials[i])));
        out.discountRates.push_back(raw.discountRates[i]);
        out.tenors.push_back(parseTenorString(raw.tenorStrings[i]));
    }

    return true;
}

static bool fillFixings(const VBAFixingInput &raw, PricingContext &ctx, std::string &error)
{
    if (!raw.indexName || raw.fixingCount <= 0)
    {
        return true; // nothing to do
    }

    if (!raw.fixingDateSerials || !raw.fixingRates)
    {
        error = "Fixing arrays are missing";
        return false;
    }

    std::vector<std::pair<Date, double>> fixings;
    fixings.reserve(static_cast<std::size_t>(raw.fixingCount));

    for (int i = 0; i < raw.fixingCount; ++i)
    {
        fixings.emplace_back(fromExcelSerial(raw.fixingDateSerials[i]), raw.fixingRates[i]);
    }

    ctx.indexFixings[raw.indexName] = fixings;
    return true;
}

static bool fillLeg(const VBALegSpec &raw, IRS::LegSpec &out, std::string &error)
{
    if (raw.notional <= 0.0)
    {
        error = "Leg notional must be positive";
        return false;
    }

    if (raw.endDateSerial <= raw.startDateSerial)
    {
        error = "Leg end date must be after start date";
        return false;
    }

    if (!mapLegType(raw.legType, out.type))
    {
        error = "Unsupported leg type";
        return false;
    }

    if (!mapPayReceive(raw.payReceive, out.payReceive))
    {
        error = "Unsupported pay/receive flag";
        return false;
    }

    if (!mapFrequency(raw.frequencyCode, out.tenor.frequency))
    {
        error = "Unsupported frequency";
        return false;
    }

    if (!mapLegDayCount(raw.dayCountCode, out.tenor.daycount))
    {
        error = "Unsupported leg day count";
        return false;
    }

    if (!mapBDC(raw.bdcCode, out.tenor.bdc))
    {
        error = "Unsupported business day convention";
        return false;
    }

    out.notional = raw.notional;
    out.startDateSerial = static_cast<long>(raw.startDateSerial);
    out.endDateSerial = static_cast<long>(raw.endDateSerial);
    out.fixed.fixedRate = raw.fixedRate;
    out.floating.indexName = raw.indexName ? raw.indexName : "";
    out.floating.fixingDays = raw.fixingDays;
    out.floating.spread = raw.spread;
    out.floating.isCompounded = raw.isCompounded != 0;

    if ((out.type == IRS::LegType::Ibor || out.type == IRS::LegType::Overnight) && out.floating.indexName.empty())
    {
        error = "Floating leg requires an index name";
        return false;
    }

    return true;
}

static bool fillSwapSpec(const VBASwapSpec &raw, IRS::IRSwapSpec &out, std::string &error)
{
    if (!mapSwapType(raw.swapType, out.swapType))
    {
        error = "Unsupported swap type";
        return false;
    }

    if (!raw.discountCurveId || !raw.valuationCurveId)
    {
        error = "Swap requires discount and valuation curve ids";
        return false;
    }

    if (raw.valuationDateSerial <= 0)
    {
        error = "Valuation date is missing";
        return false;
    }

    out.discountCurveId = raw.discountCurveId;
    out.valuationCurveId = raw.valuationCurveId;
    out.valuationDateSerial = static_cast<long>(raw.valuationDateSerial);

    if (!fillLeg(raw.leg1, out.leg1, error))
    {
        return false;
    }

    if (!fillLeg(raw.leg2, out.leg2, error))
    {
        return false;
    }

    return true;
}

static bool fillBucketConfig(const VBABucketConfig &raw,
                             const PricingContext &ctx,
                             CurveBucketConfig &out,
                             std::string &error)
{
    if (!raw.curveId || !raw.tenorStrings || raw.tenorCount <= 0)
    {
        return true; // optional
    }

    if (raw.bumpSize == 0.0)
    {
        error = "Bucket bump size cannot be zero";
        return false;
    }

    out.curveId = raw.curveId;
    out.bumpSize = raw.bumpSize;
    out.buckets.clear();
    out.buckets.reserve(static_cast<std::size_t>(raw.tenorCount));

    for (int i = 0; i < raw.tenorCount; ++i)
    {
        TenorBucket bucket;
        bucket.tenor = parseTenorString(raw.tenorStrings[i]);
        bucket.date = ctx.calendar.adjust(ctx.calendar.advance(ctx.valuationDate, bucket.tenor, Following), Following);
        out.buckets.push_back(bucket);
    }

    return true;
}

static bool buildPricingContext(double valuationDateSerial,
                                const double *holidaySerials,
                                int holidayCount,
                                const VBACurveInput *curveInputs,
                                int curveCount,
                                const VBAFixingInput *fixingInputs,
                                int fixingCount,
                                const VBABucketConfig *bucketInputs,
                                int bucketCount,
                                PricingContext &ctx,
                                std::string &error)
{
    ctx = PricingContext();

    std::vector<Date> holidays = toHolidayDates(holidaySerials, holidayCount);
    ctx.calendar = buildCalendar(holidays);
    ctx.valuationDate = ctx.calendar.adjust(fromExcelSerial(valuationDateSerial));
    Settings::instance().evaluationDate() = ctx.valuationDate;

    if (!curveInputs || curveCount <= 0)
    {
        error = "No curves supplied";
        return false;
    }

    for (int i = 0; i < curveCount; ++i)
    {
        CurveInput input;
        if (!fillCurveInput(curveInputs[i], ctx.calendar, input, error))
        {
            return false;
        }

        Handle<YieldTermStructure> curveHandle = buildZeroCurve(input, ctx.calendar);
        ctx.curves[input.id] = curveHandle;
    }

    for (int i = 0; i < fixingCount; ++i)
    {
        if (!fillFixings(fixingInputs[i], ctx, error))
        {
            return false;
        }
    }

    for (int i = 0; i < bucketCount; ++i)
    {
        CurveBucketConfig cfg;
        if (!fillBucketConfig(bucketInputs[i], ctx, cfg, error))
        {
            return false;
        }

        if (!cfg.curveId.empty())
        {
            ctx.bucketConfigs.push_back(cfg);
        }
    }

    return true;
}

static void writeBuckets(const std::vector<IRS::BucketDelta> &buckets,
                         double *outPillarSerials,
                         double *outDeltas,
                         int maxBuckets,
                         int *outUsedBuckets)
{
    int n = static_cast<int>(buckets.size());
    if (n > maxBuckets)
    {
        n = maxBuckets;
    }

    for (int i = 0; i < n; ++i)
    {
        outPillarSerials[i] = toExcelSerial(buckets[i].pillar);
        outDeltas[i] = buckets[i].delta;
    }

    if (outUsedBuckets)
    {
        *outUsedBuckets = n;
    }
}

static void zeroBuckets(double *outPillarSerials,
                        double *outDeltas,
                        int maxBuckets,
                        int *outUsedBuckets)
{
    if (outPillarSerials && outDeltas)
    {
        for (int i = 0; i < maxBuckets; ++i)
        {
            outPillarSerials[i] = 0.0;
            outDeltas[i] = 0.0;
        }
    }
    if (outUsedBuckets)
    {
        *outUsedBuckets = 0;
    }
}

// ---------------- Exported functions ----------------

extern "C" __declspec(dllexport) double __stdcall IRS_PRICE_AND_BUCKETS(
    const VBASwapSpec *swapSpec,
    const VBACurveInput *curveInputs,
    int curveCount,
    const VBAFixingInput *fixingInputs,
    int fixingCount,
    const VBABucketConfig *bucketInputs,
    int bucketCount,
    const double *holidaySerials,
    int holidayCount,
    double *outPillarSerials,
    double *outDeltas,
    int maxBuckets,
    int *outUsedBuckets)
{
    try
    {
        if (!swapSpec)
        {
            zeroBuckets(outPillarSerials, outDeltas, maxBuckets, outUsedBuckets);
            return std::numeric_limits<double>::quiet_NaN();
        }

        IRSwapSpec spec;
        std::string error;
        if (!fillSwapSpec(*swapSpec, spec, error))
        {
            zeroBuckets(outPillarSerials, outDeltas, maxBuckets, outUsedBuckets);
            return std::numeric_limits<double>::quiet_NaN();
        }

        PricingContext ctx;
        if (!buildPricingContext(swapSpec->valuationDateSerial,
                                 holidaySerials,
                                 holidayCount,
                                 curveInputs,
                                 curveCount,
                                 fixingInputs,
                                 fixingCount,
                                 bucketInputs,
                                 bucketCount,
                                 ctx,
                                 error))
        {
            zeroBuckets(outPillarSerials, outDeltas, maxBuckets, outUsedBuckets);
            return std::numeric_limits<double>::quiet_NaN();
        }

        IRSwapPricer pricer;
        PriceResult result = pricer.price(spec, ctx);

        if (outPillarSerials && outDeltas && maxBuckets > 0)
        {
            writeBuckets(result.bucketedDeltas, outPillarSerials, outDeltas, maxBuckets, outUsedBuckets);
        }
        else if (outUsedBuckets)
        {
            *outUsedBuckets = 0;
        }

        return result.npv;
    }
    catch (...)
    {
        zeroBuckets(outPillarSerials, outDeltas, maxBuckets, outUsedBuckets);
        return std::numeric_limits<double>::quiet_NaN();
    }
}

extern "C" __declspec(dllexport) void __stdcall IRS_PRICE_AND_BUCKETS_BATCH(
    const VBASwapSpec *swapSpecs,
    int swapCount,
    const VBACurveInput *curveInputs,
    int curveCount,
    const VBAFixingInput *fixingInputs,
    int fixingCount,
    const VBABucketConfig *bucketInputs,
    int bucketCount,
    const double *holidaySerials,
    int holidayCount,
    double *outNpvs,
    double *outPillarSerials,
    double *outDeltas,
    int maxBucketsPerSwap,
    int *outBucketCounts)
{
    if (!swapSpecs || swapCount <= 0)
    {
        return;
    }

    for (int i = 0; i < swapCount; ++i)
    {
        double *pillarBase = outPillarSerials ? (outPillarSerials + i * maxBucketsPerSwap) : nullptr;
        double *deltaBase = outDeltas ? (outDeltas + i * maxBucketsPerSwap) : nullptr;
        int *usedBase = outBucketCounts ? (outBucketCounts + i) : nullptr;

        double npv = IRS_PRICE_AND_BUCKETS(
            &swapSpecs[i],
            curveInputs,
            curveCount,
            fixingInputs,
            fixingCount,
            bucketInputs,
            bucketCount,
            holidaySerials,
            holidayCount,
            pillarBase,
            deltaBase,
            maxBucketsPerSwap,
            usedBase);

        if (outNpvs)
        {
            outNpvs[i] = npv;
        }
    }
}

#else // !_WIN32

// Non-Windows: no Excel exports. We leave this file as a stub.
// You could also just not compile this file at all on Linux.

#endif // _WIN32
