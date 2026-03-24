#include "script_input.hpp"

#include "paths.hpp"

#include <windows.h>

#include <sstream>
#include <string_view>

namespace adosql {

namespace {

bool Utf8BytesToWide(std::string_view bytes, std::wstring& out, std::wstring& err) {
  out.clear();
  err.clear();
  if (bytes.empty()) return true;
  int n = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, bytes.data(), static_cast<int>(bytes.size()), nullptr, 0);
  if (n == 0) {
    n = MultiByteToWideChar(CP_UTF8, 0, bytes.data(), static_cast<int>(bytes.size()), nullptr, 0);
    if (n == 0) {
      err = L"Invalid UTF-8 in script.";
      return false;
    }
  }
  out.resize(static_cast<size_t>(n));
  if (MultiByteToWideChar(CP_UTF8, 0, bytes.data(), static_cast<int>(bytes.size()), out.data(), n) == 0) {
    err = L"UTF-8 conversion failed.";
    return false;
  }
  return true;
}

}  // namespace

bool ScriptStack::Push(const std::wstring& path, std::wstring& err) {
  err.clear();
  std::wstring full;
  if (!NormalizePathToFull(path.c_str(), full, err)) return false;
  if (!FileExists(full)) {
    err = L"Script file not found: " + full;
    return false;
  }
  Level lv;
  lv.path = full;
  lv.in.open(full, std::ios::binary);
  if (!lv.in) {
    err = L"Cannot open script: " + full;
    return false;
  }
  levels_.push_back(std::move(lv));
  return true;
}

void ScriptStack::Pop() {
  if (!levels_.empty()) levels_.pop_back();
}

bool ScriptStack::ReadLine(std::wstring& out, bool& eof, bool& overflowed, std::wstring& err) {
  out.clear();
  eof = false;
  overflowed = false;
  err.clear();
  if (levels_.empty()) {
    err = L"Internal: script stack empty.";
    return false;
  }
  Level& lv = levels_.back();
  std::string line;
  if (!std::getline(lv.in, line)) {
    eof = true;
    return true;
  }
  eof = false;
  if (lv.strip_bom) {
    lv.strip_bom = false;
    if (line.size() >= 3 && static_cast<unsigned char>(line[0]) == 0xEF && static_cast<unsigned char>(line[1]) == 0xBB &&
        static_cast<unsigned char>(line[2]) == 0xBF) {
      line.erase(0, 3);
    }
  }
  while (!line.empty() && line.back() == '\r') line.pop_back();

  constexpr size_t kMax = 2048;
  if (line.size() > kMax) {
    overflowed = true;
    return true;
  }

  return Utf8BytesToWide(line, out, err);
}

}  // namespace adosql
