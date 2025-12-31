#include "MarketData.hpp"

#include <algorithm>
#include <ql/settings.hpp>
#include <ql/termstructures/yield/piecewiseyieldcurve.hpp>
#include <ql/termstructures/yield/zerocurve.hpp>
#include <ql/math/interpolations/linearinterpolation.hpp>
#include <ql/time/calendars/weekendsonly.hpp>
#include <ql/termstructures/yield/ratehelpers.hpp>
#include <ql/termstructures/yield/oisratehelper.hpp>
#include <ql/currencies/asia.hpp>
#include <iostream>

namespace IRS
{

    using namespace QuantLib;

    Calendar buildCalendar(const std::vector<Date> &holidays)
    {
        Calendar calendar = WeekendsOnly();

        for (const auto &holiday : holidays)
        {
            calendar.addHoliday(holiday);
        }

        return calendar;
    }

    Handle<YieldTermStructure> buildZeroCurve(const CurveInput &input, const Calendar &calendar)
    {
        if (input.dates.size() != input.discountRates.size())
        {
            QL_FAIL("Curve input dates/discountRates size mismatch for " << input.id);
        }

        if (input.dates.empty())
        {
            QL_FAIL("Curve input is empty for " << input.id);
        }

        // TODO - 이거 pillars 그냥 안쓰이는것 같은데...?
        std::vector<std::pair<Date, Rate>> pillars;
        pillars.reserve(input.dates.size());
        for (std::size_t i = 0; i < input.dates.size(); ++i)
        {
            pillars.emplace_back(input.dates[i], input.discountRates[i]);
        }

        std::sort(pillars.begin(), pillars.end(), [](const auto &lhs, const auto &rhs)
                  { return lhs.first < rhs.first; });

        std::vector<Date> dates;
        std::vector<Rate> discountRates;
        dates.reserve(pillars.size());
        discountRates.reserve(pillars.size());

        for (const auto &pillar : pillars)
        {
            dates.push_back(pillar.first);
            discountRates.push_back(pillar.second);
        }

        // TODO - 이거는 발생해선 안된다...!!
        if (Settings::instance().evaluationDate() == Date())
        {
            QL_FAIL("NO EVAL DATE");
            Settings::instance().evaluationDate() = calendar.adjust(dates.front());
        }

        QuantLib::RelinkableHandle<QuantLib::YieldTermStructure> forwardingTermStructure;

        if (input.id == "FWD_KOFR")
        {
            QuantLib::ext::shared_ptr<QuantLib::OvernightIndex> kofrIndex(new QuantLib::OvernightIndex(
                "KOFR", 0, QuantLib::KRWCurrency(), calendar, input.dayCounter, forwardingTermStructure));
            std::vector<QuantLib::ext::shared_ptr<QuantLib::SimpleQuote>> quotes;
            std::vector<boost::shared_ptr<QuantLib::RateHelper>> helpers;
            std::vector<Period> helperTenors;
            for (QuantLib::Size i = 0; i < input.tenors.size(); ++i)
            {
                auto q = QuantLib::ext::make_shared<QuantLib::SimpleQuote>(input.discountRates[i]);
                quotes.push_back(q);
                auto quoteHandle = QuantLib::Handle<QuantLib::Quote>(q);
                helperTenors.push_back(input.tenors[i]);
                helpers.push_back(boost::make_shared<QuantLib::OISRateHelper>(
                    0,
                    input.tenors[i],
                    quoteHandle,
                    kofrIndex,
                    forwardingTermStructure,
                    false,
                    2,
                    QuantLib::ModifiedFollowing,
                    QuantLib::Quarterly,
                    calendar));
            }

            // TODO - Zeroyield / discount, Linear / LogLinear 확실하게 결정.
            QuantLib::ext::shared_ptr<QuantLib::YieldTermStructure> discountCurve(
                new QuantLib::PiecewiseYieldCurve<QuantLib::ZeroYield, QuantLib::Linear>(
                    Settings::instance().evaluationDate(),
                    helpers,
                    input.dayCounter));
            discountCurve->enableExtrapolation();
            forwardingTermStructure.linkTo(discountCurve);
        }
        else if (input.id == "FWD_CD")
        {
            QuantLib::ext::shared_ptr<QuantLib::IborIndex> cdIndex(new QuantLib::IborIndex(
                "CD", 3 * QuantLib::Months, 2, QuantLib::KRWCurrency(),
                calendar, QuantLib::ModifiedFollowing, false, QuantLib::Actual365Fixed(), forwardingTermStructure));

            std::vector<QuantLib::ext::shared_ptr<QuantLib::RateHelper>> helpers;
            for (QuantLib::Size i = 0; i < input.tenors.size(); ++i)
            {
                helpers.push_back(QuantLib::ext::make_shared<QuantLib::SwapRateHelper>(
                    QuantLib::Handle<QuantLib::Quote>(QuantLib::ext::make_shared<QuantLib::SimpleQuote>(input.discountRates[i])),
                    input.tenors[i],
                    calendar,
                    QuantLib::Quarterly,
                    QuantLib::ModifiedFollowing,
                    QuantLib::Actual365Fixed(),
                    cdIndex));
            }
            QuantLib::ext::shared_ptr<QuantLib::YieldTermStructure> discountCurve(
                new QuantLib::PiecewiseYieldCurve<QuantLib::Discount, QuantLib::Linear>(
                    Settings::instance().evaluationDate(),
                    helpers,
                    input.dayCounter));
            discountCurve->enableExtrapolation();
            forwardingTermStructure.linkTo(discountCurve);
        }

        return forwardingTermStructure;
    }
}