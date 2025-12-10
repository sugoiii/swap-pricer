#pragma once
#include "SwapTypes.hpp"

namespace IRS {

    struct IRSwapSpec {
        LegSpec leg1;
        LegSpec leg2;

        std::string discountCurveId;
        std::string valuationCurveId;

        long valuationDateSerial;
    };
}