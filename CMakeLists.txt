cmake_minimum_required(VERSION 3.16)
project(irs_pricer CXX)

set(CMAKE_CXX_STANDARD 17)

find_package(QuantLib REQUIRED)

add_library(irs_core
    src/Pricer.cpp
    src/MarketData.cpp
)
target_include_directories(irs_core PUBLIC
    ${CMAKE_CURRENT_SOURCE_DIR}/include
)
target_link_libraries(irs_core PUBLIC QuantLib::QuantLib)

# CLI test binary (Linux/Windows)
add_executable(irs_test src/TestLinux.cpp)
target_link_libraries(irs_test PRIVATE irs_core)

# Excel DLL only on Windows
if (WIN32)
    add_library(irs_excel SHARED src/Excel.cpp)
    target_link_libraries(irs_excel PRIVATE irs_core)
    if (CMAKE_SIZEOF_VOID_P EQUAL 4)
        target_sources(irs_excel PRIVATE src/irs_excel.def)
    endif()
endif()
