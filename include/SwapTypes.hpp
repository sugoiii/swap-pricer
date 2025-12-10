#pragma once

#include<string>

namespace IRS {

    enum class LegType {
        Fixed,
        Ibor,
        Overnight
    };

    enum class PayReceive {
        Payer,
        Receiver
    };

    enum class Frequency {
        Annual,
        SemiAnnual,
        Quarterly,
        Monthly
    };

    enum class DayCount {
        Actual365Fixed,
        Actual360
    };

    enum class BusinessDayConv {
        Following,
        ModifiedFollowing,
        Preceding
    };

    struct LegTenor {
        Frequency frequency;
        DayCount daycount;
        BusinessDayConv bdc;
    };

    struct FixedLegSpec {
        double fixedRate;
    };

    struct FloatingLegSpec {
        std::string indexName;
        int fixingDays;
        double spread;
        bool isCompounded;
    };

    struct LegSpec {
        LegType type;
        PayReceive payReceive;
        double notional;

        long startDateSerial;
        long endDateSerial;

        LegTenor tenor;

        FixedLegSpec fixed;
        FloatingLegSpec floating;
    };

}