cmake_minimum_required(VERSION 3.10)
project(DocxToPdfConverter VERSION 1.0)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED True)

include_directories(include)

find_package(PkgConfig REQUIRED)

# Platform-Specific Settings
if(${CMAKE_SYSTEM_NAME} MATCHES "Darwin")  # macOS
    message(STATUS "Configuring for macOS")

    # libharu for macOS 
    find_path(LIBHARU_INCLUDE_DIR haru.h
        HINTS /usr/local/include /opt/homebrew/include
    )
    find_library(LIBHARU_LIBRARY
        NAMES hpdf
        HINTS /usr/local/lib /opt/homebrew/lib
    )

    # libzip for macOS 
    find_path(LIBZIP_INCLUDE_DIR zip.h
        HINTS /usr/local/include /opt/homebrew/include
    )
    find_library(LIBZIP_LIBRARY
        NAMES zip
        HINTS /usr/local/lib /opt/homebrew/lib
    )

    # TinyXML2 for macOS 
    find_package(TinyXML2 REQUIRED)
    if(TinyXML2_FOUND)
        include_directories(${TinyXML2_INCLUDE_DIRS})
    else()
        message(FATAL_ERROR "TinyXML2 not found on macOS")
    endif()

elseif(${CMAKE_SYSTEM_NAME} MATCHES "Linux")  # Ubuntu/Linux
    message(STATUS "Configuring for Linux")

    # libharu for Ubuntu 
    find_path(LIBHARU_INCLUDE_DIR hpdf.h
        HINTS /usr/include
    )
    find_library(LIBHARU_LIBRARY
        NAMES hpdf
        HINTS /usr/lib/x86_64-linux-gnu
    )

    # libzip for Ubuntu 
    pkg_check_modules(LibZip REQUIRED libzip)
    if(LibZip_FOUND)
        include_directories(${LibZip_INCLUDE_DIRS})
        link_directories(${LibZip_LIBRARY_DIRS})
    else()
        message(FATAL_ERROR "libzip not found on Linux")
    endif()

    # TinyXML2 for Ubuntu 
    pkg_check_modules(TinyXML2 REQUIRED tinyxml2)
    if(TinyXML2_FOUND)
        include_directories(${TinyXML2_INCLUDE_DIRS})
        link_directories(${TinyXML2_LIBRARY_DIRS})
    else()
        message(FATAL_ERROR "TinyXML2 not found on Linux")
    endif()

else()
    message(FATAL_ERROR "Unsupported platform")
endif()

# libharu 
if (LIBHARU_INCLUDE_DIR AND LIBHARU_LIBRARY)
    include_directories(${LIBHARU_INCLUDE_DIR})
else()
    message(FATAL_ERROR "libharu not found")
endif()

# Boost 
find_package(Boost COMPONENTS filesystem system REQUIRED)
if(Boost_FOUND)
    include_directories(${Boost_INCLUDE_DIRS})
else()
    message(FATAL_ERROR "Boost not found")
endif()

# Add the executable
add_executable(DocxToPdfConverter src/main.cpp)

# Link libraries
target_link_libraries(DocxToPdfConverter
    ${LIBHARU_LIBRARY}
    ${TinyXML2_LIBRARIES}
    ${LibZip_LIBRARIES}
    Boost::filesystem
    Boost::system
    z
)
