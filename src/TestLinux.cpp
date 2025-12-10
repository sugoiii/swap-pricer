// TestMain.cpp
//
// Simple CLI to test IRSwapPricer on Linux (or anywhere).
// Build it against the same core library as ExcelBridge.

#include <iostream>
#include <vector>
#include <string>


#include "SwapTypes.hpp"
#include "SwapSpec.hpp"
#include "MarketData.hpp"
#include "Pricer.hpp"
#include <ql/utilities/dataparsers.hpp>
#include <ql/time/calendars/target.hpp>
#include <ql/termstructures/yield/flatforward.hpp>

using namespace IRS;
using namespace QuantLib;

static Date fromSerial(long s) {
    return Date(s);
}

// Reuse same helpers as ExcelBridge (you can put them in a shared header)

static Period parseTenorString(const std::string& s) {
    return PeriodParser::parse(s);
}

static void buildSimpleFlatContext(
    long valuationDateSerial,
    double flatDiscountRate,
    const std::vector<std::string>& tenorStrings,
    double bumpSize,
    PricingContext& ctx
) {
    ctx.valuationDate = fromSerial(valuationDateSerial);
    Settings::instance().evaluationDate() = ctx.valuationDate;

    DayCounter dc = Actual365Fixed();
    boost::shared_ptr<YieldTermStructure> flatCurve(
        new FlatForward(ctx.valuationDate, flatDiscountRate, dc)
    );
    Handle<YieldTermStructure> flatHandle(flatCurve);

    ctx.curves["DISCOUNT"] = flatHandle;
    ctx.curves["FWD_SOFR"] = flatHandle;
    ctx.curves["FWD_KOFR"] = flatHandle;

    CurveBucketConfig cfg;
    cfg.curveId  = "DISCOUNT";
    cfg.bumpSize = bumpSize;

    for (const auto& ts : tenorStrings) {
        TenorBucket b;
        b.tenor = parseTenorString(ts);
        b.date  = Date();
        cfg.buckets.push_back(b);
    }

    ctx.bucketConfigs.push_back(cfg);

    Calendar cal = TARGET();
    for (auto& c : ctx.bucketConfigs) {
        for (auto& b : c.buckets) {
            Date d = cal.advance(ctx.valuationDate, b.tenor, Following);
            b.date = cal.adjust(d, Following);
        }
    }
}

static void buildSimpleFixedVsKOFRSwap(
    long valuationDateSerial,
    double notional,
    double fixedRate,
    long startDateSerial,
    long endDateSerial,
    IRSwapSpec& spec
) {
    spec.valuationDateSerial = valuationDateSerial;
    spec.discountCurveId     = "DISCOUNT";
    spec.valuationCurveId    = "DISCOUNT";

    spec.leg1.type             = LegType::Fixed;
    spec.leg1.payReceive       = PayReceive::Payer;
    spec.leg1.notional         = notional;
    spec.leg1.startDateSerial  = startDateSerial;
    spec.leg1.endDateSerial    = endDateSerial;
    spec.leg1.tenor.frequency  = IRS::Frequency::SemiAnnual;
    // spec.leg1.tenor.dayCount   = DayCount::Actual365Fixed;
    spec.leg1.tenor.bdc        = BusinessDayConv::ModifiedFollowing;
    spec.leg1.fixed.fixedRate  = fixedRate;

    spec.leg2.type             = LegType::Overnight;
    spec.leg2.payReceive       = PayReceive::Receiver;
    spec.leg2.notional         = notional;
    spec.leg2.startDateSerial  = startDateSerial;
    spec.leg2.endDateSerial    = endDateSerial;
    spec.leg2.tenor.frequency  = IRS::Frequency::Quarterly;
    // spec.leg2.tenor.dayCount   = DayCount::Actual365Fixed;
    spec.leg2.tenor.bdc        = BusinessDayConv::ModifiedFollowing;
    spec.leg2.floating.indexName    = "KOFR";
    spec.leg2.floating.fixingDays   = 2;
    spec.leg2.floating.spread       = 0.0;
    spec.leg2.floating.isCompounded = true;
}

int main(int argc, char** argv) {
    try {
        // Just some toy params; you can parse from argv if you want.
        long valuationDateSerial = 80000;     // arbitrary QuantLib date
        long startDateSerial     = 80001;
        long endDateSerial       = 80000 + 365 * 5; // ~5Y
        double flatRate          = 0.03;
        double notional          = 1000000.0;
        double fixedRate         = 0.032;

        PricingContext ctx;
        std::vector<std::string> tenors = {
            "1M", "3M", "6M", "1Y", "2Y", "3Y"
        };
        double bumpSize = 0.0001;

        buildSimpleFlatContext(
            valuationDateSerial,
            flatRate,
            tenors,
            bumpSize,
            ctx
        );

        IRSwapSpec spec;
        buildSimpleFixedVsKOFRSwap(
            valuationDateSerial,
            notional,
            fixedRate,
            startDateSerial,
            endDateSerial,
            spec
        );

        IRSwapPricer pricer;
        PriceResult res = pricer.price(spec, ctx);

        std::cout << "NPV: " << res.npv << "\n";

        std::cout << "Bucketed deltas:\n";
        for (const auto& b : res.bucketedDeltas) {
            std::cout
                << "  curve=" << b.curveId
                << " tenor=" << b.tenor
                << " pillar=" << b.pillar
                << " delta=" << b.delta
                << "\n";
        }

        return 0;
    } catch (std::exception& e) {
        std::cerr << "Error: " << e.what() << "\n";
        return 1;
    }
}







