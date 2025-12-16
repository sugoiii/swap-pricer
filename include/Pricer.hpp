#pragma once

#include "SwapSpec.hpp"
#include "MarketData.hpp"
#include <ql/cashflow.hpp>

namespace IRS {
    struct BucketDelta {
        std::string curveId;
        QuantLib::Period tenor;
        QuantLib::Date pillar;
        double delta;
    };

    struct PriceResult {
        double npv;
        std::vector<BucketDelta> bucketedDeltas;
    };

    class IRSwapPricer {
    public:
        PriceResult price(const IRSwapSpec& spec, const PricingContext& ctx) const;

    private:
        double npvInternal(const IRSwapSpec &spec, const PricingContext& ctx) const;

        QuantLib::Leg buildLeg(const LegSpec& legSpec,
                               const PricingContext& ctx,
                               const QuantLib::Date& valuationDate) const;

        QuantLib::Calendar resolveCalendar(const std::string& indexName,
                                           const PricingContext& ctx) const;
        QuantLib::DayCounter mapDayCount(DayCount dc) const;
        QuantLib::BusinessDayConvention mapBDC(BusinessDayConv bdc) const;
        QuantLib::Frequency mapFrequency(Frequency f) const;

        void applyIborFixings(const std::string& indexName,
                              const PricingContext& ctx,
                              const boost::shared_ptr<QuantLib::IborIndex>& index) const;

        boost::shared_ptr<QuantLib::IborIndex> resolveIborIndex(const FloatingLegSpec& floatSpec, const PricingContext& ctx) const;
        boost::shared_ptr<QuantLib::OvernightIndex> resolveOvernightIndex(const FloatingLegSpec& floatSpec, const PricingContext& ctx) const;


    };
}