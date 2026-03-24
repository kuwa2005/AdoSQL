#include "paths.hpp"

#include <windows.h>

#include <vector>

namespace adosql {

bool NormalizePathToFull(const wchar_t* path, std::wstring& out, std::wstring& error) {
  out.clear();
  error.clear();
  if (!path || !*path) {
    error = L"Empty path.";
    return false;
  }
  DWORD n = GetFullPathNameW(path, 0, nullptr, nullptr);
  if (n == 0) {
    error = L"Invalid path.";
    return false;
  }
  std::vector<wchar_t> buf(static_cast<size_t>(n));
  DWORD m = GetFullPathNameW(path, static_cast<DWORD>(buf.size()), buf.data(), nullptr);
  if (m == 0 || m >= buf.size()) {
    error = L"Failed to resolve full path.";
    return false;
  }
  out.assign(buf.data());
  return true;
}

bool FileExists(const std::wstring& p) {
  DWORD a = GetFileAttributesW(p.c_str());
  if (a == INVALID_FILE_ATTRIBUTES) return false;
  return (a & FILE_ATTRIBUTE_DIRECTORY) == 0;
}

bool DirectoryExists(const std::wstring& p) {
  DWORD a = GetFileAttributesW(p.c_str());
  if (a == INVALID_FILE_ATTRIBUTES) return false;
  return (a & FILE_ATTRIBUTE_DIRECTORY) != 0;
}

}  // namespace adosql
