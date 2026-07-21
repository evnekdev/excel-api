cmake_minimum_required(VERSION 3.20)
project(excel_com_native_control LANGUAGES CXX)

if(NOT MSVC)
    message(FATAL_ERROR "Prompt 05E requires the x64 MSVC toolchain; configure from a Visual Studio x64 developer environment.")
endif()

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

add_library(excel_com_native_shim SHARED
    src/excel_com_native_control.cpp
)
target_include_directories(excel_com_native_shim PUBLIC include)
target_compile_definitions(excel_com_native_shim PRIVATE EXCEL_COM_NATIVE_BUILD_DLL)
target_compile_options(excel_com_native_shim PRIVATE /W4 /permissive- /EHsc)
set_property(TARGET excel_com_native_shim PROPERTY MSVC_RUNTIME_LIBRARY "MultiThreadedDLL")
target_link_libraries(excel_com_native_shim PRIVATE ole32 oleaut32 user32)

add_executable(excel_com_native_control
    src/main.cpp
    src/excel_com_native_control.cpp
)
target_include_directories(excel_com_native_control PRIVATE include)
target_compile_definitions(excel_com_native_control PRIVATE EXCEL_COM_NATIVE_BUILD_DLL)
target_compile_options(excel_com_native_control PRIVATE /W4 /permissive- /EHsc)
set_property(TARGET excel_com_native_control PROPERTY MSVC_RUNTIME_LIBRARY "MultiThreadedDLL")
target_link_libraries(excel_com_native_control PRIVATE ole32 oleaut32 user32)
