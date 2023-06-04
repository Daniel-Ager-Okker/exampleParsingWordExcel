cmake_minimum_required(VERSION 3.15.0 FATAL_ERROR)
include(uic_handling.cmake)

project(main_window_pl)

set(CMAKE_INCLUDE_CURRENT_DIR ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)

find_package(Qt5 REQUIRED COMPONENTS Core Gui Widgets)

set(HEADERS
    ${CMAKE_CURRENT_SOURCE_DIR}/include/main_window_pl/main_window.h
    ${CMAKE_CURRENT_SOURCE_DIR}/include/main_window_pl/ui_main_window.h
)

set(SOURCES
    ${CMAKE_CURRENT_SOURCE_DIR}/src/main_window.cpp
)

# qt5_wrap_ui(UI_HEADERS ${FORMS} TARGET ${PROJECT_NAME})

add_library(${PROJECT_NAME} STATIC ${HEADERS} ${SOURCES})

target_include_directories(${PROJECT_NAME} PUBLIC include)

target_link_libraries(${PROJECT_NAME}
    PUBLIC
    Qt5::Core
    Qt5::Gui
    Qt5::Widgets
)

target_compile_features(${PROJECT_NAME} PUBLIC cxx_std_17)
