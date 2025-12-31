#include "Pricer.hpp"
#include <ql/time/calendars/southkorea.hpp>
#include <ql/time/calendars/weekendsonly.hpp>
#include <ql/time/daycounters/actual360.hpp>
#include <ql/indexes/iborindex.hpp>
#include <ql/currencies/asia.hpp>
#include <ql/currencies/exchangeratemanager.hpp>
#include <ql/cashflows/fixedratecoupon.hpp>
#include <ql/cashflows/iborcoupon.hpp>
#include <ql/cashflows/overnightindexedcoupon.hpp>
#include <ql/instruments/overnightindexedswap.hpp>
#include <ql/instruments/vanillaswap.hpp>
#include <ql/pricingengines/swap/discountingswapengine.hpp>
#include <ql/termstructures/yield/zerocurve.hpp>
#include <iostream>

namespace IRS
{

    using namespace QuantLib;

    Calendar IRSwapPricer::resolveCalendar(const std::string &indexName, const PricingContext &ctx) const
    {
        if (!ctx.calendar.empty())
        {
            return ctx.calendar;
        }

        if ((indexName == "KOFR") || (indexName == "CD"))
        {
            return SouthKorea(SouthKorea::Settlement);
        };

        return WeekendsOnly();
    }

    DayCounter IRSwapPricer::mapDayCount(DayCount dc) const
    {
        // TODO - 이것도 그냥 indexName으로 해도 될것 같은데
        switch (dc)
        {
        case DayCount::Actual365Fixed:
            return Actual365Fixed();
        case DayCount::Actual360:
            return Actual360();
        }
        QL_FAIL("Unsupported Daycount");
    };

    BusinessDayConvention IRSwapPricer::mapBDC(BusinessDayConv bdc) const
    {
        // 이것도 그냥 index이름으로 해도 도리듯...
        switch (bdc)
        {
        case BusinessDayConv::Following:
            return Following;
        case BusinessDayConv::ModifiedFollowing:
            return ModifiedFollowing;
        case BusinessDayConv::Preceding:
            return Preceding;
        }
        QL_FAIL("Unsupported BusinessDayConv");
    };

    QuantLib::Frequency IRSwapPricer::mapFrequency(Frequency f) const
    {
        switch (f)
        {
        case Frequency::Annual:
            return Annual;
        case Frequency::SemiAnnual:
            return Semiannual;
        case Frequency::Quarterly:
            return Quarterly;
        case Frequency::Monthly:
            return Monthly;
        }
        QL_FAIL("Unsupported Frequency");
    };

    void IRSwapPricer::applyIborFixings(const std::string &indexName,
                                        const PricingContext &ctx,
                                        const boost::shared_ptr<IborIndex> &index) const
    {
        auto fixIt = ctx.indexFixings.find(indexName);
        if (fixIt == ctx.indexFixings.end())
        {
            return;
        }

        const auto &fixings = fixIt->second;
        if (fixings.empty())
        {
            return;
        }

        std::vector<Date> fixingDates;
        std::vector<Real> fixingValues;
        fixingDates.reserve(fixings.size());
        fixingValues.reserve(fixings.size());

        for (const auto &fixing : fixings)
        {
            fixingDates.push_back(fixing.first);
            fixingValues.push_back(fixing.second);
        }

        index->addFixings(fixingDates.begin(), fixingDates.end(), fixingValues.begin(), true);
    }

    boost::shared_ptr<IborIndex> IRSwapPricer::resolveIborIndex(const FloatingLegSpec &floatSpec, const PricingContext &ctx) const
    {

        // 1) if already cached, use it
        // TODO - cached해주는 부분이 없는거같은데?
        auto it = ctx.indices.find(floatSpec.indexName);
        if (it != ctx.indices.end())
        {
            auto ibor = boost::dynamic_pointer_cast<IborIndex>(it->second);
            if (!ibor)
                QL_FAIL("Index " << floatSpec.indexName << " is not an IborIndex");
            return ibor;
        }

        // 2) Otherwise create a basic one from curves
        // in prod -> should have proper mapping: indexName -> curveId
        QuantLib::Handle<QuantLib::YieldTermStructure> fwdCurve = ctx.discountCurve("FWD_" + floatSpec.indexName);
        Calendar cal = resolveCalendar(floatSpec.indexName, ctx);
        DayCounter dc = Actual365Fixed(); // 왜 얘는 이거 고정으로 쓰냐?

        if (floatSpec.indexName == "CD")
        {
            // TODO - 캘린더 정보 제대로 넣어야함.
            auto index = boost::shared_ptr<IborIndex>(
                new QuantLib::IborIndex(
                    "CD", 3 * QuantLib::Months, 2, QuantLib::KRWCurrency(),
                    cal, QuantLib::ModifiedFollowing, false, QuantLib::Actual365Fixed(), fwdCurve));

            applyIborFixings(floatSpec.indexName, ctx, index);
            return index;
        }

        QL_FAIL("Unsupported Ibor index: " << floatSpec.indexName);
    };

    boost::shared_ptr<OvernightIndex> IRSwapPricer::resolveOvernightIndex(const FloatingLegSpec &floatSpec, const PricingContext &ctx) const
    {

        // 1) if already cached, use it
        auto it = ctx.indices.find(floatSpec.indexName);
        if (it != ctx.indices.end())
        {
            auto on = boost::dynamic_pointer_cast<OvernightIndex>(it->second);
            if (!on)
                QL_FAIL("Index " << floatSpec.indexName << " is not an OvernightIndex");
            return on;
        }

        // 2) Otherwise create a basic one from curves
        // in prod -> should have proper mapping: indexName -> curveId
        QuantLib::Handle<QuantLib::YieldTermStructure> fwdCurve = ctx.discountCurve("FWD_" + floatSpec.indexName);
        Calendar cal = resolveCalendar(floatSpec.indexName, ctx);
        DayCounter dc = Actual365Fixed(); // 왜 얘는 이거 고정으로 쓰냐?

        if (floatSpec.indexName == "KOFR")
        {
            return boost::shared_ptr<OvernightIndex>(
                new OvernightIndex("KOFR", 0, QuantLib::KRWCurrency(), cal, dc, fwdCurve));
        }
    };

    Leg IRSwapPricer::buildLeg(const LegSpec &legSpec, const PricingContext &ctx, const Date &valuationDate) const
    {
        // Convert Excel serials to Quantlib Date
        Date startDate = Date(legSpec.startDateSerial);
        Date endDate = Date(legSpec.endDateSerial);

        BusinessDayConvention bdc = mapBDC(legSpec.tenor.bdc);
        QuantLib::Frequency freq = mapFrequency(legSpec.tenor.frequency);
        DayCounter dc = mapDayCount(legSpec.tenor.daycount);

        // schedule
        // TODO - Forward vs Backward
        Schedule schedule(startDate, endDate, Period(freq), ctx.calendar, bdc, bdc, DateGeneration::Backward, false);

        Leg leg;

        switch (legSpec.type)
        {
        case LegType::Fixed:
        {
            leg = FixedRateLeg(schedule)
                      .withNotionals(legSpec.notional)
                      .withCouponRates(legSpec.fixed.fixedRate, dc);
            break;
        }

        case LegType::Ibor:
        {
            const auto &flt = legSpec.floating;
            auto index = resolveIborIndex(flt, ctx);
            leg = IborLeg(schedule, index)
                      .withNotionals(legSpec.notional)
                      .withSpreads(flt.spread);

            break;
        }
        case LegType::Overnight:
        {
            const auto &flt = legSpec.floating;
            auto onIndex = resolveOvernightIndex(flt, ctx);

            OvernightLeg onLeg(schedule, onIndex);
            onLeg.withNotionals(legSpec.notional).withSpreads(flt.spread);
            leg = onLeg;
            break;
        }
        default:
            QL_FAIL("Unsupported LegType");
        }

        return leg;
    }

    double IRSwapPricer::npvInternal(const IRSwapSpec &spec, const PricingContext &ctx) const
    {
        Date valDate = Date(spec.valuationDateSerial);
        Settings::instance().evaluationDate() = valDate;

        Handle<YieldTermStructure> disc = ctx.discountCurve(spec.discountCurveId);

        boost::shared_ptr<PricingEngine> engine(
            new DiscountingSwapEngine(disc));

        auto buildSchedule = [&](const LegSpec &legSpec)
        {
            BusinessDayConvention bdc = mapBDC(legSpec.tenor.bdc);
            QuantLib::Frequency freq = mapFrequency(legSpec.tenor.frequency);

            return Schedule(Date(legSpec.startDateSerial),
                            Date(legSpec.endDateSerial),
                            Period(freq),
                            ctx.calendar,
                            bdc,
                            bdc,
                            DateGeneration::Backward,
                            false);
        };

        switch (spec.swapType)
        {
        case SwapType::Vanilla:
        {
            const bool leg1Fixed = spec.leg1.type == LegType::Fixed;
            const bool leg2Fixed = spec.leg2.type == LegType::Fixed;
            const bool leg1Ibor = spec.leg1.type == LegType::Ibor;
            const bool leg2Ibor = spec.leg2.type == LegType::Ibor;

            const bool validFixedFloat = (leg1Fixed && leg2Ibor) || (leg2Fixed && leg1Ibor);
            QL_REQUIRE(validFixedFloat, "Vanilla swap requires one fixed leg and one Ibor leg");

            const LegSpec &fixedSpec = leg1Fixed ? spec.leg1 : spec.leg2;
            const LegSpec &floatSpec = leg1Fixed ? spec.leg2 : spec.leg1;

            Schedule fixedSchedule = buildSchedule(fixedSpec);
            Schedule floatSchedule = buildSchedule(floatSpec);

            DayCounter fixedDc = mapDayCount(fixedSpec.tenor.daycount);
            auto floatIndex = resolveIborIndex(floatSpec.floating, ctx);

            VanillaSwap::Type swapType =
                fixedSpec.payReceive == PayReceive::Payer ? VanillaSwap::Payer : VanillaSwap::Receiver;

            VanillaSwap vanillaSwap(
                swapType,
                fixedSpec.notional,
                fixedSchedule,
                fixedSpec.fixed.fixedRate,
                fixedDc,
                floatSchedule,
                floatIndex,
                floatSpec.floating.spread,
                floatIndex->dayCounter());

            vanillaSwap.setPricingEngine(engine);
            return vanillaSwap.NPV();
        }

        case SwapType::OvernightIndexed:
        {
            const bool leg1Fixed = spec.leg1.type == LegType::Fixed;
            const bool leg2Fixed = spec.leg2.type == LegType::Fixed;
            const bool leg1ON = spec.leg1.type == LegType::Overnight;
            const bool leg2ON = spec.leg2.type == LegType::Overnight;

            const bool validFixedOn = (leg1Fixed && leg2ON) || (leg2Fixed && leg1ON);
            QL_REQUIRE(validFixedOn, "Overnight indexed swap requires one fixed leg and one overnight leg");

            const LegSpec &fixedSpec = leg1Fixed ? spec.leg1 : spec.leg2;
            const LegSpec &onSpec = leg1Fixed ? spec.leg2 : spec.leg1;

            Schedule fixedSchedule = buildSchedule(fixedSpec);
            Schedule onSchedule = buildSchedule(onSpec);

            DayCounter fixedDc = mapDayCount(fixedSpec.tenor.daycount);
            auto onIndex = resolveOvernightIndex(onSpec.floating, ctx);

            OvernightIndexedSwap::Type swapType =
                fixedSpec.payReceive == PayReceive::Payer ? OvernightIndexedSwap::Payer : OvernightIndexedSwap::Receiver;

            OvernightIndexedSwap ois(
                swapType,
                fixedSpec.notional,
                fixedSchedule,
                fixedSpec.fixed.fixedRate,
                fixedDc,
                onSchedule,
                onIndex,
                onSpec.floating.spread,
                0,
                mapBDC(onSpec.tenor.bdc),
                ctx.calendar);

            ois.setPricingEngine(engine);
            return ois.NPV();
        }
        }

        QL_FAIL("Unsupported swap type");
    }

    PriceResult IRSwapPricer::price(const IRSwapSpec &spec, const PricingContext &ctx) const
    {
        PriceResult result;

        const double baseNpv = npvInternal(spec, ctx);
        result.npv = baseNpv;

        for (const auto &cfg : ctx.bucketConfigs)
        {
            auto curveIt = ctx.curves.find(cfg.curveId);
            if (curveIt == ctx.curves.end())
            {
                // 근데 없을수가 있나?
                continue;
            }

            const Handle<YieldTermStructure> &baseCurveHandle = curveIt->second;
            boost::shared_ptr<YieldTermStructure> baseCurve = baseCurveHandle.currentLink();

            std::vector<Date> dates;
            std::vector<Rate> zeroRates;

            dates.reserve(cfg.buckets.size() + 1);
            zeroRates.reserve(cfg.buckets.size() + 1);

            dates.push_back(baseCurve->referenceDate());
            Rate zr = baseCurve->zeroRate(baseCurve->referenceDate(), baseCurve->dayCounter(), Continuous).rate();
            zeroRates.push_back(zr);

            for (const auto &b : cfg.buckets)
            {
                dates.push_back(b.date);
                Rate zr = baseCurve->zeroRate(b.date, baseCurve->dayCounter(), Continuous).rate();
                zeroRates.push_back(zr);
            }

            for (std::size_t i = 0; i < cfg.buckets.size(); ++i)
            {
                const auto &bucket = cfg.buckets[i];

                // Copy base zeros and bump only this bucket i
                // TODO - bumpsize 빼서 1/2
                std::vector<Rate> bumpedZeros = zeroRates;
                bumpedZeros[i] += cfg.bumpSize;

                boost::shared_ptr<YieldTermStructure> bumpedCurve(
                    new ZeroCurve(dates, bumpedZeros, baseCurve->dayCounter(), baseCurve->calendar()));
                bumpedCurve->enableExtrapolation();
                Handle<YieldTermStructure> bumpedHandle(bumpedCurve);

                // ctx도 새로만드는듯?
                PricingContext bumpedCtx = ctx;
                bumpedCtx.curves[cfg.curveId] = bumpedHandle;

                const double bumpedNpv = npvInternal(spec, bumpedCtx);

                BucketDelta bd;
                bd.curveId = cfg.curveId;
                bd.pillar = bucket.date;
                bd.tenor = bucket.tenor;
                bd.delta = (bumpedNpv - baseNpv) / cfg.bumpSize;

                result.bucketedDeltas.push_back(bd);
            }
        }

        return result;
    }

}