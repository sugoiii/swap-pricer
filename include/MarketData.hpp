#pragma once

#include <string>
#include <vector>
#include <memory>
#include <unordered_map>
#include <ql/time/date.hpp>
#include <ql/time/daycounter.hpp>
#include <ql/time/period.hpp>
#include <ql/termstructures/yieldtermstructure.hpp>
#include <ql/indexes/interestrateindex.hpp>
#include <ql/indexes/iborindex.hpp>

namespace IRS {
    struct TenorBucket {
        QuantLib::Period tenor;
        QuantLib::Date date;
    };

    struct CurveBucketConfig{
        std::string curveId;
        std::vector<TenorBucket> buckets;
        double bumpSize;
    };
     
    struct CurveInput {
        std::string id;
        std::vector<QuantLib::Date> dates;
        std::vector<double> discountRates;
        QuantLib::DayCounter dayCounter;
    };

    struct PricingContext {
        QuantLib::Date valuationDate;

        // Curve handles keyed by id
        std::unordered_map<std::string,
            QuantLib::Handle<QuantLib::YieldTermStructure>> curves;

        // Optional: index objects keyed by name
        std::unordered_map<std::string,
            boost::shared_ptr<QuantLib::InterestRateIndex>> indices;

        std::vector<CurveBucketConfig> bucketConfigs;

        // convenience lookup
        QuantLib::Handle<QuantLib::YieldTermStructure> discountCurve(const std::string& id) const {
            auto it = curves.find(id);
            if (it == curves.end()) {
                QL_FAIL("Discount curve not found: " << id);
            }
            return it->second;
        }
        
    };

    QuantLib::Handle<QuantLib::YieldTermStructure> buildZeroCurve(const CurveInput& input);
}