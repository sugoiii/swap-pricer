// ExcelBridge.cpp
//
// Windows-only Excel-DLL bridge for IRS pricing + bucketed delta.
// On non-Windows, this file compiles to nothing useful (no exports).

#include <exception>
#include <string>
#include <vector>
#include <limits>
#include <fstream>
#include <mutex>
#include <sstream>
#include <cmath>
#include <cstddef>

#include <ql/utilities/dataparsers.hpp>
#include <ql/time/daycounters/actual360.hpp>
#include <ql/time/calendars/weekendsonly.hpp>

#include "SwapTypes.hpp"
#include "SwapSpec.hpp"
#include "MarketData.hpp"
#include "Pricer.hpp"

using namespace IRS;
using namespace QuantLib;

// ---------- Helpers shared for both test & Excel ----------

// This assumes you're using QuantLib serials directly;
// adapt if you really need "true" Excel serials later.
static Date fromExcelSerial(double serial)
{
    if (!std::isfinite(serial) || serial <= 0.0)
    {
        return Date();
    }
    long s = static_cast<long>(serial);
    if (s <= 0)
    {
        return Date();
    }
    return Date(s);
}

static double toExcelSerial(const Date &d)
{
    return static_cast<double>(d.serialNumber());
}

static bool parseTenorString(const std::string &tenor,
                             const char *contextLabel,
                             const std::string &contextId,
                             int index,
                             Period &out,
                             std::string &error)
{
    try
    {
        out = PeriodParser::parse(tenor);
        return true;
    }
    catch (const std::exception &ex)
    {
        std::ostringstream message;
        message << "Invalid tenor string '" << tenor << "' for " << contextLabel
                << " '" << contextId << "' at index " << index;
        if (ex.what() && *ex.what())
        {
            message << ": " << ex.what();
        }
        error = message.str();
        return false;
    }
    catch (...)
    {
        std::ostringstream message;
        message << "Invalid tenor string '" << tenor << "' for " << contextLabel
                << " '" << contextId << "' at index " << index;
        error = message.str();
        return false;
    }
}

static bool isFiniteNumber(double value)
{
    return std::isfinite(value);
}

static bool validateFinite(const char *label, double value, std::string &error)
{
    if (!isFiniteNumber(value))
    {
        error = std::string(label) + " must be a finite number";
        return false;
    }
    return true;
}

static bool validatePositiveSerial(double serial, const char *label, std::string &error)
{
    if (!isFiniteNumber(serial) || serial <= 0.0)
    {
        error = std::string(label) + " must be a positive serial";
        return false;
    }
    return true;
}

static bool fillHolidayDates(const double *holidaySerials,
                             int holidayCount,
                             std::vector<Date> &holidays,
                             std::string &error)
{
    holidays.clear();
    if (holidayCount <= 0)
    {
        return true;
    }

    if (!holidaySerials)
    {
        error = "Holiday serials are missing";
        return false;
    }

    holidays.reserve(static_cast<std::size_t>(holidayCount));
    for (int i = 0; i < holidayCount; ++i)
    {
        if (!validatePositiveSerial(holidaySerials[i], "Holiday date", error))
        {
            return false;
        }
        holidays.push_back(fromExcelSerial(holidaySerials[i]));
    }

    return true;
}

#ifdef _WIN32

#include <windows.h>

// ---------------- VBA-facing POD definitions ----------------
// VBA should marshal arrays using SAFEARRAY/ByRef pointers and fill these
// plain-old-data structs. All string pointers are expected to remain alive
// for the duration of the call (BSTR / UTF-16 wchar_t*). Unless otherwise noted, the
// "count" fields represent the number of elements in the corresponding array.
//
// VBACurveInput:
//   - id: curve identifier used in swap specs (e.g., "DISCOUNT", "FWD_KOFR").
//   - pillarSerials/discountRates/tenorStrings: aligned arrays of equal length
//     describing the curve pillars and their zero/discount rates.
//   - dayCountCode: 0 = Actual365Fixed, 1 = Actual360.
//
// VBAFixingInput:
//   - indexName: e.g., "KOFR" or "CD".
//   - fixingDateSerials/fixingRates: aligned arrays of historical fixings.
//
// VBABucketConfig:
//   - curveId: which curve to bump.
//   - tenorStrings: e.g., {"1M","3M","6M","1Y"}.
//   - bumpSize: absolute bump applied to zero rates.
//
// VBALegSpec + VBASwapSpec:
//   - swapType: 0 = Vanilla (fixed/ibor), 1 = OvernightIndexed (fixed/ON).
//   - leg type: 0 = Fixed, 1 = Ibor, 2 = Overnight.
//   - payReceive: 0 = Payer, 1 = Receiver.
//   - frequency: 0 = Annual, 1 = SemiAnnual, 2 = Quarterly, 3 = Monthly.
//   - dayCountCode: 0 = Actual365Fixed, 1 = Actual360.
//   - bdcCode: 0 = Following, 1 = ModifiedFollowing, 2 = Preceding.
//   - Floating legs must provide indexName; spreads are in decimal.
//   - Dates are QuantLib serials (compatible with Excel when using the
//     1900-date system). Bucketed deltas are returned as parallel arrays of
//     pillar serials and deltas per swap.

// VBA UDTs are 4-byte packed on x86 and 8-byte aligned on x64 (default MSVC).
#if defined(_WIN64)
#pragma pack(push, 8)
#else
#pragma pack(push, 4)
#endif
struct VBACurveInput
{
    const wchar_t *id;
    const double *pillarSerials;
    const double *discountRates;
    const wchar_t **tenorStrings;
    int pillarCount;
    int dayCountCode;
};

struct VBAFixingInput
{
    const wchar_t *indexName;
    const double *fixingDateSerials;
    const double *fixingRates;
    int fixingCount;
};

struct VBABucketConfig
{
    const wchar_t *curveId;
    const wchar_t **tenorStrings;
    int tenorCount;
    double bumpSize;
};

struct VBALegSpec
{
    int legType;
    int payReceive;
    double notional;
    double startDateSerial;
    double endDateSerial;
    int frequencyCode;
    int dayCountCode;
    int bdcCode;
    double fixedRate;
    const wchar_t *indexName;
    int fixingDays;
    double spread;
    int isCompounded;
};

struct VBASwapSpec
{
    int swapType;
    VBALegSpec leg1;
    VBALegSpec leg2;
    const wchar_t *discountCurveId;
    const wchar_t *valuationCurveId;
    double valuationDateSerial;
};
#pragma pack(pop)

static_assert(sizeof(void *) == sizeof(LONG_PTR),
              "Pointer size must match VBA LongPtr");
static_assert(sizeof(wchar_t) == 2,
              "VBA strings are UTF-16 (wchar_t must be 2 bytes)");

#if defined(_WIN64)
// VBA x64 UDTs are 8-byte aligned, so pointer + double fields shift offsets
// (e.g., discountCurveId in VBASwapSpec starts at 184).
static_assert(sizeof(VBACurveInput) == 40, "VBACurveInput size mismatch (x64)");
static_assert(offsetof(VBACurveInput, pillarCount) == 32, "VBACurveInput.pillarCount offset mismatch (x64)");
static_assert(offsetof(VBACurveInput, dayCountCode) == 36, "VBACurveInput.dayCountCode offset mismatch (x64)");

static_assert(sizeof(VBAFixingInput) == 32, "VBAFixingInput size mismatch (x64)");
static_assert(offsetof(VBAFixingInput, fixingCount) == 24, "VBAFixingInput.fixingCount offset mismatch (x64)");

static_assert(sizeof(VBABucketConfig) == 32, "VBABucketConfig size mismatch (x64)");
static_assert(offsetof(VBABucketConfig, tenorCount) == 16, "VBABucketConfig.tenorCount offset mismatch (x64)");
static_assert(offsetof(VBABucketConfig, bumpSize) == 24, "VBABucketConfig.bumpSize offset mismatch (x64)");

static_assert(sizeof(VBALegSpec) == 88, "VBALegSpec size mismatch (x64)");
static_assert(offsetof(VBALegSpec, startDateSerial) == 16, "VBALegSpec.startDateSerial offset mismatch (x64)");
static_assert(offsetof(VBALegSpec, fixedRate) == 48, "VBALegSpec.fixedRate offset mismatch (x64)");
static_assert(offsetof(VBALegSpec, indexName) == 56, "VBALegSpec.indexName offset mismatch (x64)");
static_assert(offsetof(VBALegSpec, spread) == 72, "VBALegSpec.spread offset mismatch (x64)");
static_assert(offsetof(VBALegSpec, isCompounded) == 80, "VBALegSpec.isCompounded offset mismatch (x64)");

static_assert(sizeof(VBASwapSpec) == 208, "VBASwapSpec size mismatch (x64)");
static_assert(offsetof(VBASwapSpec, discountCurveId) == 184,
              "VBASwapSpec.discountCurveId offset mismatch (x64)");
static_assert(offsetof(VBASwapSpec, valuationDateSerial) == 200,
              "VBASwapSpec.valuationDateSerial offset mismatch (x64)");
#else
static_assert(sizeof(VBACurveInput) == 24, "VBACurveInput size mismatch (x86)");
static_assert(offsetof(VBACurveInput, pillarCount) == 16, "VBACurveInput.pillarCount offset mismatch (x86)");
static_assert(offsetof(VBACurveInput, dayCountCode) == 20, "VBACurveInput.dayCountCode offset mismatch (x86)");

static_assert(sizeof(VBAFixingInput) == 16, "VBAFixingInput size mismatch (x86)");
static_assert(offsetof(VBAFixingInput, fixingCount) == 12, "VBAFixingInput.fixingCount offset mismatch (x86)");

static_assert(sizeof(VBABucketConfig) == 20, "VBABucketConfig size mismatch (x86)");
static_assert(offsetof(VBABucketConfig, tenorCount) == 8, "VBABucketConfig.tenorCount offset mismatch (x86)");
static_assert(offsetof(VBABucketConfig, bumpSize) == 12, "VBABucketConfig.bumpSize offset mismatch (x86)");

static_assert(sizeof(VBALegSpec) == 72, "VBALegSpec size mismatch (x86)");
static_assert(offsetof(VBALegSpec, startDateSerial) == 16, "VBALegSpec.startDateSerial offset mismatch (x86)");
static_assert(offsetof(VBALegSpec, fixedRate) == 44, "VBALegSpec.fixedRate offset mismatch (x86)");
static_assert(offsetof(VBALegSpec, indexName) == 52, "VBALegSpec.indexName offset mismatch (x86)");
static_assert(offsetof(VBALegSpec, spread) == 60, "VBALegSpec.spread offset mismatch (x86)");
static_assert(offsetof(VBALegSpec, isCompounded) == 68, "VBALegSpec.isCompounded offset mismatch (x86)");

static_assert(sizeof(VBASwapSpec) == 164, "VBASwapSpec size mismatch (x86)");
static_assert(offsetof(VBASwapSpec, discountCurveId) == 148,
              "VBASwapSpec.discountCurveId offset mismatch (x86)");
static_assert(offsetof(VBASwapSpec, valuationDateSerial) == 156,
              "VBASwapSpec.valuationDateSerial offset mismatch (x86)");
#endif

// ---------------- Mapping helpers ----------------

static thread_local std::string lastError;
static thread_local std::wstring lastErrorWide;
static std::mutex logMutex;
static bool debugEnabled = false;
static std::string debugLogPath = "swap_pricer_debug.log";

static void logDebugLine(const std::string &message)
{
    if (!debugEnabled)
    {
        return;
    }

    std::lock_guard<std::mutex> lock(logMutex);
    std::ofstream out(debugLogPath, std::ios::out | std::ios::app);
    if (out)
    {
        out << message << "\n";
    }
}

static const char *legTypeLabel(IRS::LegType type)
{
    switch (type)
    {
    case IRS::LegType::Fixed:
        return "Fixed";
    case IRS::LegType::Ibor:
        return "Ibor";
    case IRS::LegType::Overnight:
        return "Overnight";
    default:
        return "Unknown";
    }
}

static const char *payReceiveLabel(IRS::PayReceive payReceive)
{
    switch (payReceive)
    {
    case IRS::PayReceive::Payer:
        return "Payer";
    case IRS::PayReceive::Receiver:
        return "Receiver";
    default:
        return "Unknown";
    }
}

static const char *swapTypeLabel(IRS::SwapType swapType)
{
    switch (swapType)
    {
    case IRS::SwapType::Vanilla:
        return "Vanilla";
    case IRS::SwapType::OvernightIndexed:
        return "OvernightIndexed";
    default:
        return "Unknown";
    }
}

static void setLastError(const std::string &message)
{
    lastError = message;
    int needed = MultiByteToWideChar(CP_UTF8, 0, lastError.c_str(), -1, nullptr, 0);
    if (needed > 0)
    {
        lastErrorWide.assign(static_cast<std::size_t>(needed - 1), L'\0');
        MultiByteToWideChar(CP_UTF8, 0, lastError.c_str(), -1, lastErrorWide.data(), needed);
    }
    else
    {
        lastErrorWide.clear();
    }
}

static void setLastError(const char *message)
{
    lastError = message ? message : "";
    int needed = MultiByteToWideChar(CP_UTF8, 0, lastError.c_str(), -1, nullptr, 0);
    if (needed > 0)
    {
        lastErrorWide.assign(static_cast<std::size_t>(needed - 1), L'\0');
        MultiByteToWideChar(CP_UTF8, 0, lastError.c_str(), -1, lastErrorWide.data(), needed);
    }
    else
    {
        lastErrorWide.clear();
    }
}

static bool convertWideToUtf8(const wchar_t *value,
                              const char *label,
                              std::string &out,
                              std::string &error)
{   

    if (!value || !*value)
    {
        error = std::string(label) + " is missing";
        return false;
    }

    int needed = WideCharToMultiByte(CP_UTF8, 0, value, -1, nullptr, 0, nullptr, nullptr);
    if (needed <= 0)
    {
        error = std::string("Failed to convert ") + label + " to UTF-8";
        return false;
    }

    std::string buffer(static_cast<std::size_t>(needed), '\0');
    int written = WideCharToMultiByte(CP_UTF8, 0, value, -1, buffer.data(), needed, nullptr, nullptr);
    if (written <= 0)
    {
        error = std::string("Failed to convert ") + label + " to UTF-8";
        return false;
    }

    buffer.resize(static_cast<std::size_t>(written - 1));
    out = std::move(buffer);
    return true;
}

static bool mapCurveDayCount(int code, DayCounter &out, std::string &error)
{
    switch (code)
    {
    case 0:
        out = Actual365Fixed();
        return true;
    case 1:
        out = Actual360();
        return true;
    default:
        error = "Unsupported curve day count code: " + std::to_string(code);
        return false;
    }
}

static bool mapLegDayCount(int code, IRS::DayCount &out, std::string &error)
{
    switch (code)
    {
    case 0:
        out = IRS::DayCount::Actual365Fixed;
        return true;
    case 1:
        out = IRS::DayCount::Actual360;
        return true;
    default:
        error = "Unsupported leg day count code: " + std::to_string(code);
        return false;
    }
}

static bool mapBDC(int code, IRS::BusinessDayConv &out, std::string &error)
{
    switch (code)
    {
    case 0:
        out = IRS::BusinessDayConv::Following;
        return true;
    case 1:
        out = IRS::BusinessDayConv::ModifiedFollowing;
        return true;
    case 2:
        out = IRS::BusinessDayConv::Preceding;
        return true;
    default:
        error = "Unsupported business day convention code: " + std::to_string(code);
        return false;
    }
}

static bool mapFrequency(int code, IRS::Frequency &out, std::string &error)
{
    switch (code)
    {
    case 0:
        out = IRS::Frequency::Annual;
        return true;
    case 1:
        out = IRS::Frequency::SemiAnnual;
        return true;
    case 2:
        out = IRS::Frequency::Quarterly;
        return true;
    case 3:
        out = IRS::Frequency::Monthly;
        return true;
    default:
        error = "Unsupported frequency code: " + std::to_string(code);
        return false;
    }
}

static bool mapLegType(int code, IRS::LegType &out, std::string &error)
{
    switch (code)
    {
    case 0:
        out = IRS::LegType::Fixed;
        return true;
    case 1:
        out = IRS::LegType::Ibor;
        return true;
    case 2:
        out = IRS::LegType::Overnight;
        return true;
    default:
        error = "Unsupported leg type code: " + std::to_string(code);
        return false;
    }
}

static bool mapPayReceive(int code, IRS::PayReceive &out, std::string &error)
{
    switch (code)
    {
    case 0:
        out = IRS::PayReceive::Payer;
        return true;
    case 1:
        out = IRS::PayReceive::Receiver;
        return true;
    default:
        error = "Unsupported pay/receive code: " + std::to_string(code);
        return false;
    }
}

static bool mapSwapType(int code, IRS::SwapType &out, std::string &error)
{
    switch (code)
    {
    case 0:
        out = IRS::SwapType::Vanilla;
        return true;
    case 1:
        out = IRS::SwapType::OvernightIndexed;
        return true;
    default:
        error = "Unsupported swap type code: " + std::to_string(code);
        return false;
    }
}

static bool fillCurveInput(const VBACurveInput &raw,
                           const Calendar &calendar,
                           CurveInput &out,
                           std::string &error)
{
    std::string curveId;
    if (!convertWideToUtf8(raw.id, "Curve id", curveId, error))
    {
        return false;
    }

    if (raw.pillarCount <= 0)
    {
        error = "Curve pillar count must be positive";
        return false;
    }

    if (!raw.pillarSerials || !raw.discountRates || !raw.tenorStrings)
    {
        error = "Curve pillar arrays are missing";
        return false;
    }

    DayCounter dc;
    if (!mapCurveDayCount(raw.dayCountCode, dc, error))
    {
        return false;
    }

    out.id = curveId;
    out.dayCounter = dc;
    out.dates.clear();
    out.discountRates.clear();
    out.tenors.clear();

    out.dates.reserve(static_cast<std::size_t>(raw.pillarCount));
    out.discountRates.reserve(static_cast<std::size_t>(raw.pillarCount));
    out.tenors.reserve(static_cast<std::size_t>(raw.pillarCount));

    for (int i = 0; i < raw.pillarCount; ++i)
    {
        if (!validatePositiveSerial(raw.pillarSerials[i], "Curve pillar date", error))
        {
            return false;
        }

        if (!validateFinite("Curve discount rate", raw.discountRates[i], error))
        {
            return false;
        }

        if (!raw.tenorStrings[i])
        {
            error = "Curve tenor string is null";
            return false;
        }

        std::string tenorString;
        if (!convertWideToUtf8(raw.tenorStrings[i], "Curve tenor string", tenorString, error))
        {
            return false;
        }

        Period tenor;
        if (!parseTenorString(tenorString, "curve", curveId, i, tenor, error))
        {
            return false;
        }

        out.dates.push_back(calendar.adjust(fromExcelSerial(raw.pillarSerials[i])));
        out.discountRates.push_back(raw.discountRates[i]);
        out.tenors.push_back(tenor);
    }
    return true;
}

static bool fillFixings(const VBAFixingInput &raw, PricingContext &ctx, std::string &error)
{
    if (raw.fixingCount <= 0)
    {
        return true; // nothing to do
    }

    std::string indexName;
    if (!convertWideToUtf8(raw.indexName, "Fixing index name", indexName, error))
    {
        return false;
    }

    if (!raw.fixingDateSerials || !raw.fixingRates)
    {
        error = "Fixing arrays are missing";
        return false;
    }

    std::vector<std::pair<Date, double>> fixings;
    fixings.reserve(static_cast<std::size_t>(raw.fixingCount));

    for (int i = 0; i < raw.fixingCount; ++i)
    {
        if (!validatePositiveSerial(raw.fixingDateSerials[i], "Fixing date", error))
        {
            return false;
        }
        if (!validateFinite("Fixing rate", raw.fixingRates[i], error))
        {
            return false;
        }
        fixings.emplace_back(fromExcelSerial(raw.fixingDateSerials[i]), raw.fixingRates[i]);
    }

    ctx.indexFixings[indexName] = fixings;
    return true;
}

static bool fillLeg(const VBALegSpec &raw, IRS::LegSpec &out, std::string &error)
{
    if (!validateFinite("Leg notional", raw.notional, error) || raw.notional <= 0.0)
    {
        error = "Leg notional must be positive";
        return false;
    }

    if (!validatePositiveSerial(raw.startDateSerial, "Leg start date", error) ||
        !validatePositiveSerial(raw.endDateSerial, "Leg end date", error))
    {
        return false;
    }

    if (raw.endDateSerial <= raw.startDateSerial)
    {
        error = "Leg end date must be after start date";
        return false;
    }

    if (!mapLegType(raw.legType, out.type, error))
    {
        return false;
    }

    if (!mapPayReceive(raw.payReceive, out.payReceive, error))
    {
        return false;
    }

    if (!mapFrequency(raw.frequencyCode, out.tenor.frequency, error))
    {
        return false;
    }

    if (!mapLegDayCount(raw.dayCountCode, out.tenor.daycount, error))
    {
        return false;
    }

    if (!mapBDC(raw.bdcCode, out.tenor.bdc, error))
    {
        return false;
    }

    if (!validateFinite("Leg fixed rate", raw.fixedRate, error))
    {
        return false;
    }

    if (!validateFinite("Leg spread", raw.spread, error))
    {
        return false;
    }

    if (raw.fixingDays < 0)
    {
        error = "Leg fixing days must be non-negative";
        return false;
    }

    out.notional = raw.notional;
    out.startDateSerial = static_cast<long>(raw.startDateSerial);
    out.endDateSerial = static_cast<long>(raw.endDateSerial);
    out.fixed.fixedRate = raw.fixedRate;
    std::string indexName;
    if (raw.indexName && *raw.indexName)
    {
        if (!convertWideToUtf8(raw.indexName, "Floating leg index name", indexName, error))
        {
            return false;
        }
    }

    out.floating.indexName = indexName;
    out.floating.fixingDays = raw.fixingDays;
    out.floating.spread = raw.spread;
    out.floating.isCompounded = raw.isCompounded != 0;

    if (out.type == IRS::LegType::Ibor || out.type == IRS::LegType::Overnight)
    {
        if (indexName.empty())
        {
            error = "Floating leg index name is missing";
            return false;
        }
    }

    return true;
}

static bool fillSwapSpec(const VBASwapSpec &raw, IRS::IRSwapSpec &out, std::string &error)
{
    if (!mapSwapType(raw.swapType, out.swapType, error))
    {
        setLastError(error);
        return false;
    }

    std::string discountCurveId;
    if (!convertWideToUtf8(raw.discountCurveId, "Discount curve id", discountCurveId, error))
    {
        return false;
    }

    std::string valuationCurveId;
    if (!convertWideToUtf8(raw.valuationCurveId, "Valuation curve id", valuationCurveId, error))
    {
        return false;
    }

    if (!validatePositiveSerial(raw.valuationDateSerial, "Valuation date", error))
    {
        return false;
    }

    out.discountCurveId = discountCurveId;
    out.valuationCurveId = valuationCurveId;
    out.valuationDateSerial = static_cast<long>(raw.valuationDateSerial);

    {
        std::ostringstream details;
        details << "fillSwapSpec: swapType=" << swapTypeLabel(out.swapType)
                << " discountCurveId=" << out.discountCurveId
                << " valuationCurveId=" << out.valuationCurveId
                << " valuationDateSerial=" << out.valuationDateSerial;
        logDebugLine(details.str());
    }

    if (!fillLeg(raw.leg1, out.leg1, error))
    {
        setLastError(error);
        return false;
    }

    if (!fillLeg(raw.leg2, out.leg2, error))
    {
        setLastError(error);
        return false;
    }

    {
        std::ostringstream details;
        details << "fillSwapSpec: leg1 type=" << legTypeLabel(out.leg1.type)
                << " payReceive=" << payReceiveLabel(out.leg1.payReceive)
                << " notional=" << out.leg1.notional
                << " index=" << out.leg1.floating.indexName;
        logDebugLine(details.str());
    }
    {
        std::ostringstream details;
        details << "fillSwapSpec: leg2 type=" << legTypeLabel(out.leg2.type)
                << " payReceive=" << payReceiveLabel(out.leg2.payReceive)
                << " notional=" << out.leg2.notional
                << " index=" << out.leg2.floating.indexName;
        logDebugLine(details.str());

    }

    return true;
}

static bool fillBucketConfig(const VBABucketConfig &raw,
                             const PricingContext &ctx,
                             CurveBucketConfig &out,
                             std::string &error)
{
    if (raw.tenorCount <= 0)
    {
        return true; // optional
    }

    std::string curveId;
    if (!convertWideToUtf8(raw.curveId, "Bucket curve id", curveId, error))
    {
        return false;
    }

    if (!raw.tenorStrings)
    {
        error = "Bucket tenor strings are missing";
        return false;
    }

    if (!validateFinite("Bucket bump size", raw.bumpSize, error))
    {
        return false;
    }

    if (raw.bumpSize == 0.0)
    {
        error = "Bucket bump size cannot be zero";
        return false;
    }

    out.curveId = curveId;
    out.bumpSize = raw.bumpSize;
    out.buckets.clear();
    out.buckets.reserve(static_cast<std::size_t>(raw.tenorCount));

    for (int i = 0; i < raw.tenorCount; ++i)
    {
        if (!raw.tenorStrings[i])
        {
            error = "Bucket tenor string is null";
            return false;
        }

        std::string tenorString;
        if (!convertWideToUtf8(raw.tenorStrings[i], "Bucket tenor string", tenorString, error))
        {
            return false;
        }

        TenorBucket bucket;
        if (!parseTenorString(tenorString, "bucket", curveId, i, bucket.tenor, error))
        {
            return false;
        }
        bucket.date = ctx.calendar.adjust(ctx.calendar.advance(ctx.valuationDate, bucket.tenor, Following), Following);
        out.buckets.push_back(bucket);
    }

    return true;
}

static bool buildPricingContext(double valuationDateSerial,
                                const double *holidaySerials,
                                int holidayCount,
                                const VBACurveInput *curveInputs,
                                int curveCount,
                                const VBAFixingInput *fixingInputs,
                                int fixingCount,
                                const VBABucketConfig *bucketInputs,
                                int bucketCount,
                                PricingContext &ctx,
                                std::string &error)
{
    ctx = PricingContext();

    if (!validatePositiveSerial(valuationDateSerial, "Valuation date", error))
    {
        return false;
    }

    std::vector<Date> holidays;
    if (holidayCount > 0)
    {
        if (!fillHolidayDates(holidaySerials, holidayCount, holidays, error))
        {
            return false;
        }
    }
    ctx.calendar = buildCalendar(holidays);
    ctx.valuationDate = ctx.calendar.adjust(fromExcelSerial(valuationDateSerial));
    Settings::instance().evaluationDate() = ctx.valuationDate;
    {
        std::ostringstream details;
        details << "buildPricingContext: valuationDateSerial=" << valuationDateSerial
                << " valuationDate=" << ctx.valuationDate
                << " holidayCount=" << holidayCount;
        logDebugLine(details.str());
    }

    if (curveCount <= 0)
    {
        error = "No curves supplied";
        setLastError(error);
        return false;
    }

    if (!curveInputs)
    {
        error = "Curve inputs are missing";
        setLastError(error);
        return false;
    }

    if (fixingCount > 0 && !fixingInputs)
    {
        error = "Fixing inputs are missing";
        setLastError(error);
        return false;
    }

    if (bucketCount > 0 && !bucketInputs)
    {
        error = "Bucket inputs are missing";
        setLastError(error);
        return false;
    }

    {
        std::ostringstream details;
        details << "buildPricingContext: curveCount=" << curveCount
                << " fixingCount=" << fixingCount
                << " bucketCount=" << bucketCount;
        logDebugLine(details.str());
    }

    for (int i = 0; i < curveCount; ++i)
    {
        CurveInput input;
        if (!fillCurveInput(curveInputs[i], ctx.calendar, input, error))
        {
            setLastError(error);
            return false;
        }

        {
            std::ostringstream details;
            details << "buildPricingContext: curveId=" << input.id
                    << " pillars=" << input.dates.size()
                    << " dayCountCode=" << curveInputs[i].dayCountCode;
            logDebugLine(details.str());
        }

        Handle<YieldTermStructure> curveHandle = buildZeroCurve(input, ctx.calendar);
        ctx.curves[input.id] = curveHandle;
    }

    for (int i = 0; i < fixingCount; ++i)
    {
        if (!fillFixings(fixingInputs[i], ctx, error))
        {
            setLastError(error);
            return false;
        }

        if (fixingInputs && fixingInputs[i].indexName)
        {
            std::string fixingName;
            std::string conversionError;
            if (!convertWideToUtf8(fixingInputs[i].indexName, "Fixing index name", fixingName, conversionError))
            {
                fixingName = "<invalid>";
            }
            std::ostringstream details;
            details << "buildPricingContext: fixings index=" << fixingName
                    << " count=" << fixingInputs[i].fixingCount;
            logDebugLine(details.str());
        }
    }

    for (int i = 0; i < bucketCount; ++i)
    {
        CurveBucketConfig cfg;
        if (!fillBucketConfig(bucketInputs[i], ctx, cfg, error))
        {
            setLastError(error);
            return false;
        }

        if (!cfg.curveId.empty())
        {
            std::ostringstream details;
            details << "buildPricingContext: bucketConfig curveId=" << cfg.curveId
                    << " buckets=" << cfg.buckets.size()
                    << " bumpSize=" << cfg.bumpSize;
            logDebugLine(details.str());
            ctx.bucketConfigs.push_back(cfg);
        }
    }

    return true;
}

static void writeBuckets(const std::vector<IRS::BucketDelta> &buckets,
                         double *outPillarSerials,
                         double *outDeltas,
                         int maxBuckets,
                         int *outUsedBuckets)
{
    int n = static_cast<int>(buckets.size());
    if (n > maxBuckets)
    {
        n = maxBuckets;
    }

    for (int i = 0; i < n; ++i)
    {
        outPillarSerials[i] = toExcelSerial(buckets[i].pillar);
        outDeltas[i] = buckets[i].delta;
    }

    if (outUsedBuckets)
    {
        *outUsedBuckets = n;
    }
}

static void zeroBuckets(double *outPillarSerials,
                        double *outDeltas,
                        int maxBuckets,
                        int *outUsedBuckets)
{
    if (outPillarSerials && outDeltas)
    {
        for (int i = 0; i < maxBuckets; ++i)
        {
            outPillarSerials[i] = 0.0;
            outDeltas[i] = 0.0;
        }
    }
    if (outUsedBuckets)
    {
        *outUsedBuckets = 0;
    }
}

static double failAndReturnNaN(const std::string &message,
                               double *outPillarSerials,
                               double *outDeltas,
                               int maxBuckets,
                               int *outUsedBuckets)
{
    setLastError(message);
    zeroBuckets(outPillarSerials, outDeltas, maxBuckets, outUsedBuckets);
    return std::numeric_limits<double>::quiet_NaN();
}

// ---------------- Exported functions ----------------
#if defined(_WIN64)
#define IRS_EXCEL_CALL
#else
#define IRS_EXCEL_CALL __stdcall
#endif
#define IRS_EXCEL_EXPORT extern "C" __declspec(dllexport)

IRS_EXCEL_EXPORT double IRS_EXCEL_CALL IRS_PRICE_AND_BUCKETS(
    const VBASwapSpec *swapSpec,
    const VBACurveInput *curveInputs,
    int curveCount,
    const VBAFixingInput *fixingInputs,
    int fixingCount,
    const VBABucketConfig *bucketInputs,
    int bucketCount,
    const double *holidaySerials,
    int holidayCount,
    double *outPillarSerials,
    double *outDeltas,
    int maxBuckets,
    int *outUsedBuckets)
{
    try
    {
        lastError.clear();
        lastErrorWide.clear();
        if (!swapSpec)
        {
            return failAndReturnNaN("Swap spec is null", outPillarSerials, outDeltas, maxBuckets, outUsedBuckets);
        }

        IRSwapSpec spec;
        std::string error;
        if (!fillSwapSpec(*swapSpec, spec, error))
        {
            return failAndReturnNaN(error, outPillarSerials, outDeltas, maxBuckets, outUsedBuckets);
        }

        const double *safeHolidaySerials = holidayCount > 0 ? holidaySerials : nullptr;
        int safeHolidayCount = holidayCount > 0 ? holidayCount : 0;
        const VBAFixingInput *safeFixingInputs = fixingCount > 0 ? fixingInputs : nullptr;
        int safeFixingCount = fixingCount > 0 ? fixingCount : 0;
        const VBABucketConfig *safeBucketInputs = bucketCount > 0 ? bucketInputs : nullptr;
        int safeBucketCount = bucketCount > 0 ? bucketCount : 0;

        PricingContext ctx;
        if (!buildPricingContext(swapSpec->valuationDateSerial,
                                 safeHolidaySerials,
                                 safeHolidayCount,
                                 curveInputs,
                                 curveCount,
                                 safeFixingInputs,
                                 safeFixingCount,
                                 safeBucketInputs,
                                 safeBucketCount,
                                 ctx,
                                 error))
        {
            return failAndReturnNaN(error, outPillarSerials, outDeltas, maxBuckets, outUsedBuckets);
        }

        IRSwapPricer pricer;
        PriceResult result = pricer.price(spec, ctx);

        if (outPillarSerials && outDeltas && maxBuckets > 0)
        {
            writeBuckets(result.bucketedDeltas, outPillarSerials, outDeltas, maxBuckets, outUsedBuckets);
        }
        else if (outUsedBuckets)
        {
            *outUsedBuckets = 0;
        }

        return result.npv;
    }
    catch (const std::exception &ex)
    {
        setLastError(ex.what());
        zeroBuckets(outPillarSerials, outDeltas, maxBuckets, outUsedBuckets);
        return std::numeric_limits<double>::quiet_NaN();
    }
    catch (...)
    {
        setLastError("Unhandled exception in IRS_PRICE_AND_BUCKETS");
        zeroBuckets(outPillarSerials, outDeltas, maxBuckets, outUsedBuckets);
        return std::numeric_limits<double>::quiet_NaN();
    }
}

IRS_EXCEL_EXPORT const wchar_t *IRS_EXCEL_CALL IRS_LAST_ERROR()
{
    return lastErrorWide.c_str();
}

IRS_EXCEL_EXPORT int IRS_EXCEL_CALL IRS_IS_NAN(double value)
{
    return std::isnan(value) ? 1 : 0;
}

IRS_EXCEL_EXPORT void IRS_EXCEL_CALL IRS_SET_DEBUG_MODE(int enabled)
{
    debugEnabled = enabled != 0;
}

IRS_EXCEL_EXPORT void IRS_EXCEL_CALL IRS_SET_DEBUG_LOG_PATH(const wchar_t *path)
{
    std::string converted;
    std::string conversionError;
    if (path && *path &&
        convertWideToUtf8(path, "Debug log path", converted, conversionError))
    {
        debugLogPath = converted;
    }
}

IRS_EXCEL_EXPORT void IRS_EXCEL_CALL IRS_PRICE_AND_BUCKETS_BATCH(
    const VBASwapSpec *swapSpecs,
    int swapCount,
    const VBACurveInput *curveInputs,
    int curveCount,
    const VBAFixingInput *fixingInputs,
    int fixingCount,
    const VBABucketConfig *bucketInputs,
    int bucketCount,
    const double *holidaySerials,
    int holidayCount,
    double *outNpvs,
    double *outPillarSerials,
    double *outDeltas,
    int maxBucketsPerSwap,
    int *outBucketCounts)
{
    if (!swapSpecs || swapCount <= 0)
    {
        return;
    }

    for (int i = 0; i < swapCount; ++i)
    {
        double *pillarBase = outPillarSerials ? (outPillarSerials + i * maxBucketsPerSwap) : nullptr;
        double *deltaBase = outDeltas ? (outDeltas + i * maxBucketsPerSwap) : nullptr;
        int *usedBase = outBucketCounts ? (outBucketCounts + i) : nullptr;

        double npv = IRS_PRICE_AND_BUCKETS(
            &swapSpecs[i],
            curveInputs,
            curveCount,
            fixingInputs,
            fixingCount,
            bucketInputs,
            bucketCount,
            holidaySerials,
            holidayCount,
            pillarBase,
            deltaBase,
            maxBucketsPerSwap,
            usedBase);

        if (outNpvs)
        {
            outNpvs[i] = npv;
        }
    }
}

#else // !_WIN32

// Non-Windows: no Excel exports. We leave this file as a stub.
// You could also just not compile this file at all on Linux.

#endif // _WIN32
