#include "MarketData.hpp"

#include <algorithm>
#include <ql/settings.hpp>
#include <ql/termstructures/yield/piecewiseyieldcurve.hpp>
#include <ql/termstructures/yield/zerocurve.hpp>
#include <ql/math/interpolations/linearinterpolation.hpp>
#include <ql/time/calendars/target.hpp>

namespace IRS {

    using namespace QuantLib;

    Handle<YieldTermStructure> buildZeroCurve(const CurveInput& input) {
        if (input.dates.size() != input.discountRates.size()) {
            QL_FAIL("Curve input dates/discountRates size mismatch for " << input.id);
        }

        if (input.dates.empty()) {
            QL_FAIL("Curve input is empty for " << input.id);
        }

        std::vector<std::pair<Date, Rate>> pillars;
        pillars.reserve(input.dates.size());
        for (std::size_t i = 0; i < input.dates.size(); ++i) {
            pillars.emplace_back(input.dates[i], input.discountRates[i]);
        }

        std::sort(pillars.begin(), pillars.end(), [](const auto& lhs, const auto& rhs) {
            return lhs.first < rhs.first;
        });

        std::vector<Date> dates;
        std::vector<Rate> zeroRates;
        dates.reserve(pillars.size());
        zeroRates.reserve(pillars.size());

        for (const auto& pillar : pillars) {
            dates.push_back(pillar.first);
            zeroRates.push_back(pillar.second);
        }

        Calendar calendar = TARGET();
        if (Settings::instance().evaluationDate() == Date()) {
            Settings::instance().evaluationDate() = calendar.adjust(dates.front());
        }

        using ZeroInterpolation = PiecewiseYieldCurve<ZeroYield, Linear>;

        boost::shared_ptr<YieldTermStructure> curve(
            new ZeroInterpolation(dates.front(), dates, zeroRates, input.dayCounter, calendar)
        );

        return Handle<YieldTermStructure>(curve);
    }
}