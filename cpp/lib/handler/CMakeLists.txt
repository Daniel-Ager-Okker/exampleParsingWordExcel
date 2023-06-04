cmake_minimum_required(VERSION 3.15.0 FATAL_ERROR)

project(handler)


set(HEADERS
    ${CMAKE_CURRENT_SOURCE_DIR}/include/handler/handler.h
)
foreach(headerFile IN LISTS HEADERS)
    message("There is header-file: ${headerFile}")
endforeach()

set(SOURCES
    ${CMAKE_CURRENT_SOURCE_DIR}/src/handler.cpp
)
foreach(srcFile IN LISTS SOURCES)
    message("There is source-file: ${srcFile}")
endforeach()

add_library(${PROJECT_NAME} STATIC ${HEADERS} ${SOURCES})

target_include_directories(${PROJECT_NAME} PUBLIC include)

target_compile_features(${PROJECT_NAME} PUBLIC cxx_std_17)
