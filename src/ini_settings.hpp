#pragma once

#include "settings.hpp"

#include <string>

namespace adosql {

bool LoadIniNextToExe(DisplaySettings& out, std::wstring& err);
bool SaveIniNextToExe(const DisplaySettings& in, std::wstring& err);

}  // namespace adosql
