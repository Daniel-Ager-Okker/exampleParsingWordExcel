cmake_minimum_required(VERSION 3.15.0 FATAL_ERROR)
project(VelesStroyConverterApp)

set(SOURCES
    ${CMAKE_CURRENT_SOURCE_DIR}/main.cpp
)

add_executable(${PROJECT_NAME} ${SOURCES})

add_subdirectory(../lib/handler ${CMAKE_CURRENT_BINARY_DIR}/subproject/handler)
add_subdirectory(../plugin/main_window_pl ${CMAKE_CURRENT_BINARY_DIR}/subproject/main_window_pl)

target_link_libraries(${PROJECT_NAME}
    PRIVATE
    handler
    main_window_pl
)