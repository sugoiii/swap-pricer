// TestMain.cpp
//
// Simple CLI to test IRSwapPricer on Linux (or anywhere).
// Build it against the same core library as ExcelBridge.

#include <iostream>
#include <vector>
#include <string>
#include <algorithm>
#include <cctype>

#include "SwapTypes.hpp"
#include "SwapSpec.hpp"
#include "MarketData.hpp"
#include "Pricer.hpp"
#include <ql/utilities/dataparsers.hpp>
#include <ql/time/calendars/target.hpp>
#include <ql/time/daycounters/actual365fixed.hpp>

using namespace IRS;
using namespace QuantLib;

static Date fromSerial(long s)
{
    return Date(s);
}

// Reuse same helpers as ExcelBridge (you can put them in a shared header)

static Period parseTenorString(const std::string &s)
{
    return PeriodParser::parse(s);
}

static Date parseHoliday(const std::string &token)
{
    const bool allDigits = std::all_of(token.begin(), token.end(), [](char c)
                                       { return std::isdigit(c); });
    if (allDigits)
    {
        return Date(std::stol(token));
    }

    return DateParser::parseISO(token);
}

static std::vector<Date> parseHolidayArgs(int argc, char **argv)
{
    std::vector<Date> holidays;

    for (int i = 1; i < argc; ++i)
    {
        const std::string token(argv[i]);
        if (token.rfind("--holiday=", 0) == 0)
        {
            holidays.push_back(parseHoliday(token.substr(std::string("--holiday=").size())));
        }
        else if (token == "--help")
        {
            std::cout
                << "Usage: " << argv[0] << " [--holiday=YYYY-MM-DD|SERIAL]..."
                << "\n"
                << "Provide QuantLib serials or ISO dates for holidays.\n";
            std::exit(0);
        }
    }

    if (holidays.empty())
    {
        // Default Swiss holidays for KOFR/CHF testing
        holidays.push_back(Date(1, January, 2025));
        holidays.push_back(Date(29, March, 2025)); // Good Friday
        holidays.push_back(Date(1, April, 2025));  // Easter Monday
        holidays.push_back(Date(9, May, 2025));    // Ascension
        holidays.push_back(Date(20, May, 2025));   // Whit Monday
        holidays.push_back(Date(25, December, 2025));
        holidays.push_back(Date(26, December, 2025));
    }

    return holidays;
}

static CurveInput buildCurveInput(
    const std::string &curveId,
    const std::vector<TenorBucket> &tenorBuckets,
    const std::vector<double> &discountRates,
    const DayCounter &dc)
{
    if (tenorBuckets.size() != discountRates.size())
    {
        QL_FAIL("Curve input size mismatch for " << curveId);
    }

    CurveInput input;
    input.id = curveId;
    input.dayCounter = dc;
    input.discountRates = discountRates;
    input.dates.reserve(tenorBuckets.size());
    input.tenors.reserve(tenorBuckets.size());

    for (auto bucket : tenorBuckets)
    {
        input.dates.push_back(bucket.date);
        input.tenors.push_back(bucket.tenor);
    }

    return input;
}

static std::vector<long> buildPillarSerialsFromTenors(
    const Date &valuationDate,
    const std::vector<std::string> &tenors,
    const Calendar &calendar)
{
    std::vector<long> serials;
    serials.reserve(tenors.size());

    for (const auto &tenorString : tenors)
    {
        Period tenor = parseTenorString(tenorString);
        Date d = calendar.advance(valuationDate, tenor, Following);
        serials.push_back(calendar.adjust(d, Following).serialNumber());
    }

    return serials;
}

static std::vector<TenorBucket> buildTenorBucketsFromTenors(
    const Date &valuationDate,
    const std::vector<std::string> &tenors,
    const Calendar &calendar)
{
    std::vector<TenorBucket> tenorBuckets;
    tenorBuckets.reserve(tenors.size());

    // TODO - Day Convention 제대로 설정.
    for (const auto &tenorString : tenors)
    {
        Period tenor = parseTenorString(tenorString);
        Date d = calendar.advance(valuationDate, tenor, Following);
        tenorBuckets.push_back({tenor, calendar.adjust(d, Following)});
    }

    return tenorBuckets;
}

static CurveBucketConfig buildBucketConfig(
    const CurveInput &input,
    const Date &valuationDate,
    double bumpSize,
    const Calendar &calendar,
    const std::vector<std::string> &tenorStrings = {})
{
    CurveBucketConfig cfg;
    cfg.curveId = input.id;
    cfg.bumpSize = bumpSize;

    const bool hasTenors = tenorStrings.size() == input.dates.size();

    for (std::size_t i = 0; i < input.dates.size(); ++i)
    {
        TenorBucket bucket;
        bucket.date = input.dates[i];
        if (hasTenors)
        {
            bucket.tenor = parseTenorString(tenorStrings[i]);
        }
        else
        {
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
    IRSwapSpec &spec)
{
    spec.valuationDateSerial = valuationDateSerial;
    spec.discountCurveId = "FWD_KOFR";
    spec.valuationCurveId = "FWD_KOFR";
    spec.swapType = SwapType::OvernightIndexed;

    spec.leg1.type = LegType::Fixed;
    spec.leg1.payReceive = PayReceive::Payer;
    spec.leg1.notional = notional;
    spec.leg1.startDateSerial = startDateSerial;
    spec.leg1.endDateSerial = endDateSerial;
    spec.leg1.tenor.frequency = IRS::Frequency::SemiAnnual;
    // spec.leg1.tenor.dayCount   = DayCount::Actual365Fixed;
    spec.leg1.tenor.bdc = BusinessDayConv::ModifiedFollowing;
    spec.leg1.fixed.fixedRate = fixedRate;

    spec.leg2.type = LegType::Overnight;
    spec.leg2.payReceive = PayReceive::Receiver;
    spec.leg2.notional = notional;
    spec.leg2.startDateSerial = startDateSerial;
    spec.leg2.endDateSerial = endDateSerial;
    spec.leg2.tenor.frequency = IRS::Frequency::Quarterly;
    // spec.leg2.tenor.dayCount   = DayCount::Actual365Fixed;
    spec.leg2.tenor.bdc = BusinessDayConv::ModifiedFollowing;
    spec.leg2.floating.indexName = "KOFR";
    spec.leg2.floating.fixingDays = 2;
    spec.leg2.floating.spread = 0.0;
    spec.leg2.floating.isCompounded = true;
}

static void buildSimpleFixedVsCDSwap(
    long valuationDateSerial,
    double notional,
    double fixedRate,
    long startDateSerial,
    long endDateSerial,
    IRSwapSpec &spec)
{
    spec.valuationDateSerial = valuationDateSerial;
    spec.discountCurveId = "FWD_CD";
    spec.valuationCurveId = "FWD_CD";
    spec.swapType = SwapType::Vanilla;

    spec.leg1.type = LegType::Fixed;
    spec.leg1.payReceive = PayReceive::Payer;
    spec.leg1.notional = notional;
    spec.leg1.startDateSerial = startDateSerial;
    spec.leg1.endDateSerial = endDateSerial;
    spec.leg1.tenor.frequency = IRS::Frequency::Quarterly;
    // spec.leg1.tenor.dayCount   = DayCount::Actual365Fixed;
    spec.leg1.tenor.bdc = BusinessDayConv::ModifiedFollowing;
    spec.leg1.fixed.fixedRate = fixedRate;

    spec.leg2.type = LegType::Ibor;
    spec.leg2.payReceive = PayReceive::Receiver;
    spec.leg2.notional = notional;
    spec.leg2.startDateSerial = startDateSerial;
    spec.leg2.endDateSerial = endDateSerial;
    spec.leg2.tenor.frequency = IRS::Frequency::Quarterly;
    // spec.leg2.tenor.dayCount   = DayCount::Actual365Fixed;
    spec.leg2.tenor.bdc = BusinessDayConv::ModifiedFollowing;
    spec.leg2.floating.indexName = "CD";
    spec.leg2.floating.fixingDays = 0;
    spec.leg2.floating.spread = 0.0;
    spec.leg2.floating.isCompounded = false; // TODO -
    
}

int main(int argc, char **argv)
{
    try
    {
        long valuationDateSerial = 46007; // arbitrary QuantLib date
        long startDateSerial = 44447;
        long endDateSerial = startDateSerial + 365 * 5; // ~5Y
        double notional = 10000000000.0;
        double fixedRate = 0.016;

        DayCounter dc = Actual365Fixed();
        Date valuationDate = fromSerial(valuationDateSerial);

        std::vector<Date> holidays = parseHolidayArgs(argc, argv);
        PricingContext ctx;
        ctx.calendar = buildCalendar(holidays);
        ctx.valuationDate = ctx.calendar.adjust(valuationDate);
        Settings::instance().evaluationDate() = ctx.valuationDate;

        std::vector<std::string> tenorStrs = {
            "1M", "3M", "6M", "1Y", "2Y", "3Y"};

        // Sample curve data; replace with real user input or CLI parsing as needed
        std::vector<TenorBucket> discountTenorBuckets = buildTenorBucketsFromTenors(ctx.valuationDate, tenorStrs, ctx.calendar);
        std::vector<double> kofrRates = {0.02825, 0.0286, 0.0289, 0.0291, 0.02985, 0.0306};
        std::vector<double> cdRates = {0.0283, 0.0283, 0.02745, 0.0277, 0.029175, 0.0305};

        CurveInput kofrCurveIn = buildCurveInput("FWD_KOFR", discountTenorBuckets, kofrRates, dc);
        CurveInput cdCurveIn = buildCurveInput("FWD_CD", discountTenorBuckets, cdRates, dc);

        ctx.curves[kofrCurveIn.id] = buildZeroCurve(kofrCurveIn, ctx.calendar);
        ctx.curves[cdCurveIn.id] = buildZeroCurve(cdCurveIn, ctx.calendar);

        std::vector<std::pair<Date, double>> kofrFixings;
        kofrFixings.emplace_back(ctx.calendar.adjust(ctx.valuationDate - 1), kofrRates.front());
        ctx.indexFixings["KOFR"] = kofrFixings;

        std::vector<std::pair<Date, double>> cdFixings;
        cdFixings.emplace_back(ctx.calendar.adjust(ctx.valuationDate - 1), cdRates.front());
        cdFixings.emplace_back(Date(45995), 0.0281);
        ctx.indexFixings["CD"] = cdFixings;

        double bumpSize = 0.0001;
        ctx.bucketConfigs.push_back(
            buildBucketConfig(kofrCurveIn, ctx.valuationDate, bumpSize, ctx.calendar, tenorStrs));

        IRSwapSpec spec;

        // buildSimpleFixedVsKOFRSwap(
        //     valuationDateSerial,
        //     notional,
        //     fixedRate,
        //     startDateSerial,
        //     endDateSerial,
        //     spec);

        buildSimpleFixedVsCDSwap(
            valuationDateSerial,
            notional,
            fixedRate,
            startDateSerial,
            endDateSerial,
            spec);

        IRSwapPricer pricer;
        PriceResult res = pricer.price(spec, ctx);

        std::cout << "NPV: " << res.npv << "\n";

        std::cout << "Bucketed deltas:\n";
        for (const auto &b : res.bucketedDeltas)
        {
            std::cout
                << "  curve=" << b.curveId
                << " tenor=" << b.tenor
                << " pillar=" << b.pillar
                << " delta=" << b.delta
                << "\n";
        }

        return 0;
    }
    catch (std::exception &e)
    {
        std::cerr << "Error: " << e.what() << "\n";
        return 1;
    }
}
