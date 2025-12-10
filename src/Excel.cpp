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

static void buildSimpleFlatContext(
    double valuationDateSerial,
    double flatDiscountRate,
    const std::vector<std::string> &tenorStrings, // e.g. {"1M","3M","6M","1Y"}
    double bumpSize,                              // e.g. 0.0001
    PricingContext &ctx)
{
    ctx.valuationDate = fromExcelSerial(valuationDateSerial);
    Settings::instance().evaluationDate() = ctx.valuationDate;

    DayCounter dc = Actual365Fixed();
    boost::shared_ptr<YieldTermStructure> flatCurve(
        new FlatForward(ctx.valuationDate, flatDiscountRate, dc));
    Handle<YieldTermStructure> flatHandle(flatCurve);

    ctx.curves["DISCOUNT"] = flatHandle;
    ctx.curves["FWD_SOFR"] = flatHandle;
    ctx.curves["FWD_KOFR"] = flatHandle;

    CurveBucketConfig cfg;
    cfg.curveId = "DISCOUNT";
    cfg.bumpSize = bumpSize;

    for (const auto &ts : tenorStrings)
    {
        TenorBucket b;
        b.tenor = parseTenorString(ts);
        b.date = Date(); // filled below
        cfg.buckets.push_back(b);
    }

    ctx.bucketConfigs.push_back(cfg);

    Calendar cal = TARGET();

    for (auto &c : ctx.bucketConfigs)
    {
        for (auto &b : c.buckets)
        {
            Date d = cal.advance(ctx.valuationDate, b.tenor, Following);
            b.date = cal.adjust(d, Following);
        }
    }
}

static void buildSimpleFixedVsSOFRSwap(
    double valuationDateSerial,
    double notional,
    double fixedRate,
    double startDateSerial,
    double endDateSerial,
    IRSwapSpec &spec)
{
    spec.valuationDateSerial = static_cast<long>(valuationDateSerial);
    spec.discountCurveId = "DISCOUNT";
    spec.valuationCurveId = "DISCOUNT";

    spec.leg1.type = LegType::Fixed;
    spec.leg1.payReceive = PayReceive::Payer;
    spec.leg1.notional = notional;
    spec.leg1.startDateSerial = static_cast<long>(startDateSerial);
    spec.leg1.endDateSerial = static_cast<long>(endDateSerial);
    spec.leg1.tenor.frequency = IRS::Frequency::SemiAnnual;
    // spec.leg1.tenor.dayCount = DayCount::Actual365Fixed;
    spec.leg1.tenor.bdc = BusinessDayConv::ModifiedFollowing;
    spec.leg1.fixed.fixedRate = fixedRate;

    spec.leg2.type = LegType::Overnight;
    spec.leg2.payReceive = PayReceive::Receiver;
    spec.leg2.notional = notional;
    spec.leg2.startDateSerial = static_cast<long>(startDateSerial);
    spec.leg2.endDateSerial = static_cast<long>(endDateSerial);
    spec.leg2.tenor.frequency = IRS::Frequency::Quarterly;
    // spec.leg2.tenor.dayCount = DayCount::Actual365Fixed;
    spec.leg2.tenor.bdc = BusinessDayConv::ModifiedFollowing;
    spec.leg2.floating.indexName = "KOFR";
    spec.leg2.floating.fixingDays = 2;
    spec.leg2.floating.spread = 0.0;
    spec.leg2.floating.isCompounded = true;
}

#ifdef _WIN32

// Windows: we actually do Excel exports.

#include <windows.h>

// Simple NPV function
extern "C" __declspec(dllexport) double __stdcall IRS_NPV_SIMPLE_FIXED_FLOAT(
    double valuationDateSerial,
    double flatDiscountRate,
    double notional,
    double fixedRate,
    double startDateSerial,
    double endDateSerial)
{
    try
    {
        PricingContext ctx;
        std::vector<std::string> tenors = {"1M", "3M", "6M", "1Y", "2Y", "5Y"};
        double bumpSize = 0.0001;

        buildSimpleFlatContext(
            valuationDateSerial,
            flatDiscountRate,
            tenors,
            bumpSize,
            ctx);

        IRSwapSpec spec;
        buildSimpleFixedVsSOFRSwap(
            valuationDateSerial,
            notional,
            fixedRate,
            startDateSerial,
            endDateSerial,
            spec);

        IRSwapPricer pricer;
        PriceResult res = pricer.price(spec, ctx);
        return res.npv;
    }
    catch (...)
    {
        return std::numeric_limits<double>::quiet_NaN();
    }
}

// Bucketed delta function
extern "C" __declspec(dllexport) void __stdcall IRS_BUCKETED_DELTA_SIMPLE_FIXED_FLOAT(
    double valuationDateSerial,
    double flatDiscountRate,
    double notional,
    double fixedRate,
    double startDateSerial,
    double endDateSerial,
    double *outPillarSerials,
    double *outDeltas,
    int maxBuckets,
    int *outUsedBuckets)
{
    try
    {
        PricingContext ctx;
        std::vector<std::string> tenors = {"1M", "3M", "6M", "1Y", "2Y", "5Y"};
        double bumpSize = 0.0001;

        buildSimpleFlatContext(
            valuationDateSerial,
            flatDiscountRate,
            tenors,
            bumpSize,
            ctx);

        IRSwapSpec spec;
        buildSimpleFixedVsSOFRSwap(
            valuationDateSerial,
            notional,
            fixedRate,
            startDateSerial,
            endDateSerial,
            spec);

        IRSwapPricer pricer;
        PriceResult res = pricer.price(spec, ctx);

        int n = static_cast<int>(res.bucketedDeltas.size());
        if (n > maxBuckets)
        {
            n = maxBuckets;
        }

        for (int i = 0; i < n; ++i)
        {
            const BucketDelta &b = res.bucketedDeltas[i];
            outPillarSerials[i] = toExcelSerial(b.pillar);
            outDeltas[i] = b.delta;
        }

        *outUsedBuckets = n;
    }
    catch (...)
    {
        *outUsedBuckets = 0;
    }
}

#else // !_WIN32

// Non-Windows: no Excel exports. We leave this file as a stub.
// You could also just not compile this file at all on Linux.

#endif // _WIN32