cmake_minimum_required(VERSION 3.20)
project(adosql VERSION 0.1.0 LANGUAGES CXX RC)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")

set(ADOSQL_SOURCES
  src/main.cpp
  src/cli.cpp
  src/console.cpp
  src/paths.cpp
  src/ado_session.cpp
  src/repl.cpp
  src/ini_settings.cpp
  src/script_input.cpp
)

add_executable(adosql ${ADOSQL_SOURCES} src/version.rc)

target_include_directories(adosql PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}/src)

target_compile_definitions(adosql PRIVATE
  UNICODE
  _UNICODE
  WIN32_LEAN_AND_MEAN
  NOMINMAX
  _WIN32_WINNT=0x0A00
  ADOSQL_VERSION_MAJOR=${PROJECT_VERSION_MAJOR}
  ADOSQL_VERSION_MINOR=${PROJECT_VERSION_MINOR}
  ADOSQL_VERSION_PATCH=${PROJECT_VERSION_PATCH}
)

if(MSVC)
  target_compile_options(adosql PRIVATE /W4 /permissive- /EHsc)
endif()

target_link_libraries(adosql PRIVATE
  ole32
  oleaut32
  uuid
  comsuppw
)
