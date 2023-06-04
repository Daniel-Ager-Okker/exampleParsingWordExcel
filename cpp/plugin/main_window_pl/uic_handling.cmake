function(ui2Header pathToUiFile pathToHeaderFolder)
    get_filename_component(fileName ${pathToUiFile} NAME_WE)
    set(pathToHeaderFile "${pathToHeaderFolder}/ui_${fileName}.h")

    execute_process(
    COMMAND uic ${pathToUiFile} -o ${pathToHeaderFile}
    WORKING_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}
    )
endfunction()

function(deleteUiHeader pathToUiHeader)
    if(EXISTS ${pathToUiHeader})
        message("There is already ${pathToUiHeader}. Need to delete")
        file(REMOVE ${pathToUiHeader})
    else()
        message("No ${pathToUiHeader}. Nothing to delete")
    endif()
endfunction()


set(FORMS
    ${CMAKE_CURRENT_SOURCE_DIR}/forms/main_window.ui
)

set(PATH_TO_HEADERS ${CMAKE_CURRENT_SOURCE_DIR}/include/main_window_pl)

foreach(uiFile IN LISTS FORMS)
    message("File-ui: ${uiFile}. Need to make header-file from it\n")

    get_filename_component(fileName ${uiFile} NAME_WE)
    set(pathToHeaderFile "${PATH_TO_HEADERS}/ui_${fileName}.h")
    message("Generated header-file will be: ${pathToHeaderFile}")

    deleteUiHeader(${pathToHeaderFile})
    ui2Header(${uiFile} ${PATH_TO_HEADERS})
endforeach()

