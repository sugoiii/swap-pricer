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
#include <ql/time/daycounters/actual365fixed.hpp>

using namespace IRS;
using namespace QuantLib;

static Date fromSerial(long s) {
    return Date(s);
}

// Reuse same helpers as ExcelBridge (you can put them in a shared header)

static Period parseTenorString(const std::string& s) {
    return PeriodParser::parse(s);
}

static CurveInput buildCurveInput(
    const std::string& curveId,
    const std::vector<long>& pillarSerials,
    const std::vector<double>& discountRates,
    const DayCounter& dc) {
    if (pillarSerials.size() != discountRates.size()) {
        QL_FAIL("Curve input size mismatch for " << curveId);
    }

    CurveInput input;
    input.id = curveId;
    input.dayCounter = dc;
    input.discountRates = discountRates;
    input.dates.reserve(pillarSerials.size());

    for (long serial : pillarSerials) {
        input.dates.push_back(fromSerial(serial));
    }

    return input;
}

static std::vector<long> buildPillarSerialsFromTenors(
    const Date& valuationDate,
    const std::vector<std::string>& tenors,
    const Calendar& calendar)
{
    std::vector<long> serials;
    serials.reserve(tenors.size());

    for (const auto& tenorString : tenors) {
        Period tenor = parseTenorString(tenorString);
        Date d = calendar.advance(valuationDate, tenor, Following);
        serials.push_back(calendar.adjust(d, Following).serialNumber());
    }

    return serials;
}

static CurveBucketConfig buildBucketConfig(
    const CurveInput& input,
    const Date& valuationDate,
    double bumpSize,
    const Calendar& calendar,
    const std::vector<std::string>& tenorStrings = {})
{
    CurveBucketConfig cfg;
    cfg.curveId = input.id;
    cfg.bumpSize = bumpSize;

    const bool hasTenors = tenorStrings.size() == input.dates.size();

    for (std::size_t i = 0; i < input.dates.size(); ++i) {
        TenorBucket bucket;
        bucket.date = input.dates[i];
        if (hasTenors) {
            bucket.tenor = parseTenorString(tenorStrings[i]);
        } else {
            Integer days = bucket.date - valuationDate;
            bucket.tenor = Period(days, Days);
        }
        // Ensure dates remain business-adjusted to match curve pillars
        bucket.date = calendar.adjust(bucket.date, Following);
        cfg.buckets.push_back(bucket);
    }

    return cfg;
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
        double notional          = 1000000.0;
        double fixedRate         = 0.032;

        DayCounter dc = Actual365Fixed();
        Date valuationDate = fromSerial(valuationDateSerial);

        std::vector<Date> holidays;
        PricingContext ctx;
        ctx.calendar = buildCalendar(holidays);
        ctx.valuationDate = ctx.calendar.adjust(valuationDate);
        Settings::instance().evaluationDate() = ctx.valuationDate;

        std::vector<std::string> tenors = {
            "1M", "3M", "6M", "1Y", "2Y", "3Y"
        };

        // Sample curve data; replace with real user input or CLI parsing as needed
        std::vector<long> discountPillars = buildPillarSerialsFromTenors(ctx.valuationDate, tenors, ctx.calendar);
        std::vector<double> discountRates = {0.0295, 0.02975, 0.0300, 0.0305, 0.0315, 0.0325};

        std::vector<long> sofrPillars = discountPillars;
        std::vector<double> sofrRates  = {0.0280, 0.0285, 0.02875, 0.0290, 0.02975, 0.0305};

        std::vector<long> kofrPillars = discountPillars;
        std::vector<double> kofrRates  = {0.02825, 0.0286, 0.0289, 0.0291, 0.02985, 0.0306};

        CurveInput discountCurve = buildCurveInput("DISCOUNT", discountPillars, discountRates, dc);
        CurveInput sofrCurve     = buildCurveInput("FWD_SOFR", sofrPillars, sofrRates, dc);
        CurveInput kofrCurve     = buildCurveInput("FWD_KOFR", kofrPillars, kofrRates, dc);

        ctx.curves[discountCurve.id] = buildZeroCurve(discountCurve, ctx.calendar);
        ctx.curves[sofrCurve.id]     = buildZeroCurve(sofrCurve, ctx.calendar);
        ctx.curves[kofrCurve.id]     = buildZeroCurve(kofrCurve, ctx.calendar);

        double bumpSize = 0.0001;
        ctx.bucketConfigs.push_back(
            buildBucketConfig(discountCurve, ctx.valuationDate, bumpSize, ctx.calendar, tenors)
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







