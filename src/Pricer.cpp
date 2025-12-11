#include "Pricer.hpp"
#include <ql/time/calendars/southkorea.hpp>
#include <ql/time/calendars/target.hpp>
#include <ql/time/daycounters/actual360.hpp>
#include <ql/indexes/iborindex.hpp>
#include <ql/currencies/asia.hpp>
#include <ql/currencies/exchangeratemanager.hpp>
#include <ql/cashflows/fixedratecoupon.hpp>
#include <ql/cashflows/iborcoupon.hpp>
#include <ql/cashflows/overnightindexedcoupon.hpp>
#include <ql/instruments/swap.hpp>
#include <ql/pricingengines/swap/discountingswapengine.hpp>
#include <ql/termstructures/yield/zerocurve.hpp>


namespace IRS {

    using namespace QuantLib;

    Calendar IRSwapPricer::resolveCalendar(const std::string& indexName, const PricingContext& ctx) const {
        if (!ctx.calendar.empty()) {
            return ctx.calendar;
        }

        // Extend LATER ->

        // TODO - SouthKorea가 아닌 실제 휴일 정보 받아서 캘린더 만들기.
        if ((indexName == "KOFR") || (indexName=="CD")) {
            return SouthKorea(SouthKorea::Settlement);
        };

        return TARGET();
    }

    DayCounter IRSwapPricer::mapDayCount(DayCount dc) const {
        // TODO - 이것도 그냥 indexName으로 해도 될것 같은데
        switch(dc) {
            case DayCount::Actual365Fixed: return Actual365Fixed();
            case DayCount::Actual360: return Actual360();
        }
        QL_FAIL("Unsupported Daycount");
    };

    BusinessDayConvention IRSwapPricer::mapBDC(BusinessDayConv bdc) const {
        // 이것도 그냥 index이름으로 해도 도리듯...
        switch(bdc) {
            case BusinessDayConv::Following: return Following;
            case BusinessDayConv::ModifiedFollowing: return ModifiedFollowing;
            case BusinessDayConv::Preceding: return Preceding;
        }
        QL_FAIL("Unsupported BusinessDayConv");
    };

    QuantLib::Frequency IRSwapPricer::mapFrequency(Frequency f) const {
        switch(f) {
            case Frequency::Annual: return Annual;
            case Frequency::SemiAnnual: return Semiannual;
            case Frequency::Quarterly: return Quarterly;
            case Frequency::Monthly: return Monthly;
        }
        QL_FAIL("Unsupported Frequency");
    };

    boost::shared_ptr<IborIndex> IRSwapPricer::resolveIborIndex(const FloatingLegSpec& floatSpec, const PricingContext& ctx) const {
        
        // 1) if already cached, use it
        auto it = ctx.indices.find(floatSpec.indexName);
        if (it != ctx.indices.end()) {
            auto ibor = boost::dynamic_pointer_cast<IborIndex>(it->second);
            if (!ibor) QL_FAIL("Index " << floatSpec.indexName << " is not an IborIndex");
            return ibor;
        }

        // 2) Otherwise create a basic one from curves
        // in prod -> should have proper mapping: indexName -> curveId
        QuantLib::Handle<QuantLib::YieldTermStructure> fwdCurve = ctx.discountCurve("FWD_" + floatSpec.indexName);
        Calendar cal = resolveCalendar(floatSpec.indexName, ctx);
        DayCounter dc = Actual365Fixed(); // 왜 얘는 이거 고정으로 쓰냐?

        if (floatSpec.indexName == "CD") {
            // TODO - 캘린더 정보 제대로 넣어야함.
            return boost::shared_ptr<IborIndex>(
                new QuantLib::IborIndex(
                    "CD91", 3 * QuantLib::Months, 2, QuantLib::KRWCurrency(),
                    TARGET(), QuantLib::ModifiedFollowing, false, QuantLib::Actual365Fixed(), fwdCurve)
            );
        }
    };

    boost::shared_ptr<OvernightIndex> IRSwapPricer::resolveOvernightIndex(const FloatingLegSpec& floatSpec, const PricingContext& ctx) const {
        
        // 1) if already cached, use it
        auto it = ctx.indices.find(floatSpec.indexName);
        if (it != ctx.indices.end()) {
            auto on = boost::dynamic_pointer_cast<OvernightIndex>(it->second);
            if (!on) QL_FAIL("Index " << floatSpec.indexName << " is not an OvernightIndex");
            return on;
        }

        // 2) Otherwise create a basic one from curves
        // in prod -> should have proper mapping: indexName -> curveId
        QuantLib::Handle<QuantLib::YieldTermStructure> fwdCurve = ctx.discountCurve("FWD_" + floatSpec.indexName);
        Calendar cal = resolveCalendar(floatSpec.indexName, ctx);
        DayCounter dc = Actual365Fixed(); // 왜 얘는 이거 고정으로 쓰냐?

        if (floatSpec.indexName == "KOFR") {
            return boost::shared_ptr<OvernightIndex>(
                new OvernightIndex("KOFR", 0, QuantLib::KRWCurrency(), cal, dc, fwdCurve)
            );
        }

    };

    Leg IRSwapPricer::buildLeg(const LegSpec& legSpec, const PricingContext& ctx, const Date& valuationDate) const {
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

        switch (legSpec.type) {
            case LegType::Fixed: {
                leg = FixedRateLeg(schedule)
                .withNotionals(legSpec.notional)
                .withCouponRates(legSpec.fixed.fixedRate, dc);
                break;
            }

            case LegType::Ibor: {
                const auto& flt = legSpec.floating;
                auto index = resolveIborIndex(flt, ctx);
                leg = IborLeg(schedule, index)
                    .withNotionals(legSpec.notional)
                    .withSpreads(flt.spread);

                break;
            }
            case LegType::Overnight: {
                const auto& flt = legSpec.floating;
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

    double IRSwapPricer::npvInternal(const IRSwapSpec& spec, const PricingContext& ctx) const {
        Date valDate = Date(spec.valuationDateSerial);
        Settings::instance().evaluationDate() = valDate;

        // Build Legs
        Leg leg1 = buildLeg(spec.leg1, ctx, valDate);
        Leg leg2 = buildLeg(spec.leg2, ctx, valDate);

        std::vector<Leg> legs{ leg1, leg2};
        std::vector<bool> payerFlags {
            spec.leg1.payReceive==PayReceive::Payer,
            spec.leg2.payReceive==PayReceive::Payer
        };

        Swap swap(legs, payerFlags);

        Handle<YieldTermStructure> disc = ctx.discountCurve(spec.discountCurveId);

        boost::shared_ptr<PricingEngine> engine (
            new DiscountingSwapEngine(disc)
        );
        swap.setPricingEngine(engine);

        return swap.NPV();
    }

    PriceResult IRSwapPricer::price(const IRSwapSpec& spec, const PricingContext& ctx) const {
        PriceResult result;

        const double baseNpv = npvInternal(spec, ctx);
        result.npv = baseNpv;

        for (const auto& cfg : ctx.bucketConfigs) {
            auto curveIt = ctx.curves.find(cfg.curveId);
            if (curveIt == ctx.curves.end()) {
                // 근데 없을수가 있나?
                continue;
            }

            const Handle<YieldTermStructure>& baseCurveHandle = curveIt ->second;
            boost::shared_ptr<YieldTermStructure> baseCurve = baseCurveHandle.currentLink();

            std::vector<Date> dates;
            std::vector<Rate> zeroRates;

            dates.reserve(cfg.buckets.size());
            zeroRates.reserve(cfg.buckets.size());

            for (const auto& b : cfg.buckets) {
                dates.push_back(b.date);
                Rate zr = baseCurve->zeroRate(b.date, baseCurve->dayCounter(), Continuous).rate();
                zeroRates.push_back(zr);
            }

            for (std::size_t i = 0; i < cfg.buckets.size(); ++i) {
                const auto& bucket = cfg.buckets[i];

                // Copy base zeros and bump only this bucket i
                // TODO - bumpsize 빼서 1/2
                std::vector<Rate> bumpedZeros = zeroRates;
                bumpedZeros[i] += cfg.bumpSize;

                boost::shared_ptr<YieldTermStructure> bumpedCurve(
                    new ZeroCurve(dates, bumpedZeros, baseCurve->dayCounter(), baseCurve->calendar())
                );
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
    }

}