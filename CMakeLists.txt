cmake_minimum_required(VERSION 3.10)
project(DocxToPdfConverter VERSION 1.0)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED True)

find_package(PkgConfig REQUIRED)

include_directories(include)

if(POLICY CMP0167)
    cmake_policy(SET CMP0167 OLD)
endif()

# Platform-Specific Settings
if(${CMAKE_SYSTEM_NAME} MATCHES "Darwin")  # macOS
    message(STATUS "Configuring for macOS")

    # Define possible Homebrew prefixes
    set(HOMEBREW_PREFIXES /usr/local /opt/homebrew)

    # libharu for macOS
    find_path(LIBHARU_INCLUDE_DIR hpdf.h
        HINTS
            ${HOMEBREW_PREFIXES}/include
    )

    find_library(LIBHARU_LIBRARY
        NAMES hpdf
        HINTS
            ${HOMEBREW_PREFIXES}/lib
    )

    if(NOT LIBHARU_INCLUDE_DIR)
        message(FATAL_ERROR "Could not find libharu headers (hpdf.h)")
    endif()

    if(NOT LIBHARU_LIBRARY)
        message(FATAL_ERROR "Could not find libharu library (libhpdf)")
    endif()

    # Include the directories
    include_directories(${LIBHARU_INCLUDE_DIR})

    # libzip for macOS
    find_path(LIBZIP_INCLUDE_DIR zip.h
        HINTS
            ${HOMEBREW_PREFIXES}/include
    )
    find_library(LIBZIP_LIBRARY
        NAMES zip
        HINTS
            ${HOMEBREW_PREFIXES}/lib
    )

    if(NOT LIBZIP_INCLUDE_DIR)
        message(FATAL_ERROR "Could not find libzip headers (zip.h)")
    endif()

    if(NOT LIBZIP_LIBRARY)
        message(FATAL_ERROR "Could not find libzip library (libzip)")
    endif()

    include_directories(${LIBZIP_INCLUDE_DIR})

    # TinyXML2 for macOS
    find_path(TINYXML2_INCLUDE_DIR tinyxml2.h
        HINTS
            ${HOMEBREW_PREFIXES}/include
    )
    find_library(TINYXML2_LIBRARY
        NAMES tinyxml2
        HINTS
            ${HOMEBREW_PREFIXES}/lib
    )
    if(NOT TINYXML2_INCLUDE_DIR)
        message(FATAL_ERROR "Could not find TinyXML2 headers (tinyxml2.h)")
    endif()
    if(NOT TINYXML2_LIBRARY)
        message(FATAL_ERROR "Could not find TinyXML2 library (tinyxml2)")
    endif()

    include_directories(${TINYXML2_INCLUDE_DIR})

    # Set TINYXML2_LIB for macOS
    set(TINYXML2_LIB ${TINYXML2_LIBRARY})

    # Set LIBZIP_LIB for macOS
    set(LIBZIP_LIB ${LIBZIP_LIBRARY})

elseif(${CMAKE_SYSTEM_NAME} MATCHES "Linux")  # Ubuntu/Linux
    message(STATUS "Configuring for Linux")

    include_directories(/usr/include /usr/include/x86_64-linux-gnu)

    # libharu for Ubuntu
    find_path(LIBHARU_INCLUDE_DIR hpdf.h
        HINTS /usr/include
    )
    find_library(LIBHARU_LIBRARY
        NAMES hpdf
        HINTS /usr/lib/x86_64-linux-gnu
    )

    if(NOT LIBHARU_INCLUDE_DIR OR NOT LIBHARU_LIBRARY)
        message(FATAL_ERROR "libharu not found on Linux")
    endif()

    include_directories(${LIBHARU_INCLUDE_DIR})

    # libzip for Ubuntu
    pkg_check_modules(LibZip REQUIRED libzip)
    if(LibZip_FOUND)
        include_directories(${LibZip_INCLUDE_DIRS})
        # Set LIBZIP_LIB for Linux
        set(LIBZIP_LIB ${LibZip_LIBRARIES})
    else()
        message(FATAL_ERROR "libzip not found on Linux")
    endif()

    # TinyXML2 for Ubuntu
    pkg_check_modules(TinyXML2 REQUIRED tinyxml2)
    if(TinyXML2_FOUND)
        include_directories(${TinyXML2_INCLUDE_DIRS})
        # Set TINYXML2_LIB for Linux
        set(TINYXML2_LIB TinyXML2::TinyXML2)
    else()
        message(FATAL_ERROR "TinyXML2 not found on Linux")
    endif()

else()
    message(FATAL_ERROR "Unsupported platform")
endif()

# Boost
find_package(Boost COMPONENTS filesystem system REQUIRED)
if(Boost_FOUND)
    include_directories(${Boost_INCLUDE_DIRS})
else()
    message(FATAL_ERROR "Boost not found")
endif()

# Add the executable
add_executable(DocxToPdfConverter src/main.cpp src/DocxParser.cpp src/DocxToPdfConverter.cpp)

# Link libraries conditionally based on platform
target_link_libraries(DocxToPdfConverter
    ${LIBHARU_LIBRARY}
    ${LIBZIP_LIB}
    ${TINYXML2_LIB}
    Boost::filesystem
    Boost::system
    z
)
