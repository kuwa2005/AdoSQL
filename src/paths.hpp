#pragma once

#include <string>

namespace adosql {

bool NormalizePathToFull(const wchar_t* path, std::wstring& out, std::wstring& error);
bool FileExists(const std::wstring& path);
bool DirectoryExists(const std::wstring& path);

}  // namespace adosql
