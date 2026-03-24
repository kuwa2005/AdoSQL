#include "ini_settings.hpp"

#include <algorithm>
#include <cctype>
#include <cwctype>
#include <fstream>
#include <sstream>
#include <string>
#include <vector>

#include <windows.h>

namespace adosql {

namespace {

std::wstring IniPath(std::wstring& err) {
  err.clear();
  wchar_t mod[MAX_PATH];
  DWORD n = GetModuleFileNameW(nullptr, mod, MAX_PATH);
  if (n == 0 || n >= MAX_PATH) {
    err = L"GetModuleFileNameW failed.";
    return {};
  }
  std::wstring p(mod);
  const auto slash = p.find_last_of(L"\\/");
  if (slash == std::wstring::npos) {
    err = L"Invalid module path.";
    return {};
  }
  return p.substr(0, slash + 1) + L"adosql.ini";
}

void TrimAsciiWs(std::string& s) {
  while (!s.empty() && std::isspace(static_cast<unsigned char>(s.front()))) s.erase(s.begin());
  while (!s.empty() && std::isspace(static_cast<unsigned char>(s.back()))) s.pop_back();
}

bool IcEqAscii(std::string_view a, std::string_view b) {
  if (a.size() != b.size()) return false;
  for (size_t i = 0; i < a.size(); ++i) {
    if (std::tolower(static_cast<unsigned char>(a[i])) != std::tolower(static_cast<unsigned char>(b[i]))) return false;
  }
  return true;
}

}  // namespace

bool LoadIniNextToExe(DisplaySettings& out, std::wstring& err) {
  err.clear();
  const std::wstring path = IniPath(err);
  if (path.empty()) return false;

  std::ifstream f(path, std::ios::binary);
  if (!f) return true;  // optional file

  std::string line;
  bool bom = true;
  while (std::getline(f, line)) {
    if (bom) {
      bom = false;
      if (line.size() >= 3 && static_cast<unsigned char>(line[0]) == 0xEF && static_cast<unsigned char>(line[1]) == 0xBB &&
          static_cast<unsigned char>(line[2]) == 0xBF) {
        line.erase(0, 3);
      }
    }
    TrimAsciiWs(line);
    if (line.empty() || line[0] == '#' || (line.size() >= 2 && line[0] == '/' && line[1] == '/')) continue;
    const auto eq = line.find('=');
    if (eq == std::string::npos) continue;
    std::string key = line.substr(0, eq);
    std::string val = line.substr(eq + 1);
    TrimAsciiWs(key);
    TrimAsciiWs(val);
    try {
      if (IcEqAscii(key, "pagesize")) out.pagesize = std::stoi(val);
      else if (IcEqAscii(key, "linesize"))
        out.linesize = std::stoi(val);
      else if (IcEqAscii(key, "maxrows"))
        out.maxrows = std::stoi(val);
      else if (IcEqAscii(key, "widthclamp"))
        out.widthclamp = IcEqAscii(val, "on") || IcEqAscii(val, "1") || IcEqAscii(val, "true");
    } catch (...) {
    }
  }
  out.pagesize = std::clamp(out.pagesize, 5, 500);
  out.linesize = std::clamp(out.linesize, 20, 500);
  if (out.maxrows < 0) out.maxrows = 0;
  return true;
}

bool SaveIniNextToExe(const DisplaySettings& in, std::wstring& err) {
  err.clear();
  const std::wstring path = IniPath(err);
  if (path.empty()) return false;

  DisplaySettings s = in;
  s.pagesize = std::clamp(s.pagesize, 5, 500);
  s.linesize = std::clamp(s.linesize, 20, 500);
  if (s.maxrows < 0) s.maxrows = 0;

  std::ostringstream os;
  os << "# adosql.ini (UTF-8)\n";
  os << "pagesize=" << s.pagesize << "\n";
  os << "linesize=" << s.linesize << "\n";
  os << "maxrows=" << s.maxrows << "\n";
  os << "widthclamp=" << (s.widthclamp ? "on" : "off") << "\n";

  const std::string utf8 = os.str();
  std::ofstream f(path, std::ios::binary | std::ios::trunc);
  if (!f) {
    err = L"Failed to write adosql.ini.";
    return false;
  }
  f.write(utf8.data(), static_cast<std::streamsize>(utf8.size()));
  return true;
}

}  // namespace adosql
