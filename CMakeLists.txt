cmake_minimum_required(VERSION 3.10)

# Project name and version
project(DocxToPdfConverter VERSION 1.0)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED True)

# directory for header files
include_directories(include)

find_package(PkgConfig REQUIRED)

# libharu
find_path(LIBHARU_INCLUDE_DIR haru.h
    HINTS /usr/local/include /opt/homebrew/include
)

find_library(LIBHARU_LIBRARY
    NAMES haru
    HINTS /usr/local/lib /opt/homebrew/lib
)

# TinyXML2
find_package(TinyXML2 REQUIRED)
if(TinyXML2_FOUND)
    include_directories(${TinyXML2_INCLUDE_DIRS})
else()
    message(FATAL_ERROR "TinyXML2 not found")
endif()

# libzip
find_package(PkgConfig REQUIRED)
pkg_check_modules(LibZip REQUIRED libzip)
if(LibZip_FOUND)
    include_directories(${LibZip_INCLUDE_DIRS})
    link_directories(${LibZip_LIBRARY_DIRS})
else()
    message(FATAL_ERROR "libzip not found")
endif()

# Boost 
find_package(Boost COMPONENTS filesystem system REQUIRED)
if(Boost_FOUND)
    include_directories(${Boost_INCLUDE_DIRS})
endif()

add_executable(DocxToPdfConverter src/main.cpp)

# Link libraries
target_link_libraries(DocxToPdfConverter
    ${LIBHaru_LIBRARIES}
    ${TinyXML2_LIBRARIES}
    ${LibZip_LIBRARIES}
    Boost::filesystem
    Boost::system
    z
)
