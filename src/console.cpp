#include "console.hpp"

#include <windows.h>

#include <cwctype>
#include <cstring>
#include <string>

namespace adosql {

namespace {

HANDLE StdOut() { return GetStdHandle(STD_OUTPUT_HANDLE); }
HANDLE StdIn() { return GetStdHandle(STD_INPUT_HANDLE); }

std::string WideToUtf8(std::wstring_view w) {
  if (w.empty()) return {};
  int n = WideCharToMultiByte(CP_UTF8, 0, w.data(), static_cast<int>(w.size()), nullptr, 0, nullptr, nullptr);
  if (n <= 0) return {};
  std::string s(static_cast<size_t>(n), '\0');
  WideCharToMultiByte(CP_UTF8, 0, w.data(), static_cast<int>(w.size()), s.data(), n, nullptr, nullptr);
  return s;
}

}  // namespace

void PrintUtf8(const char* utf8) {
  if (!utf8) return;
  DWORD written = 0;
  WriteFile(StdOut(), utf8, static_cast<DWORD>(std::strlen(utf8)), &written, nullptr);
}

void PrintUtf8(const std::string& utf8) {
  DWORD written = 0;
  WriteFile(StdOut(), utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr);
}

void PrintUtf8Wide(const std::wstring& ws) { PrintUtf8(WideToUtf8(ws)); }

void PrintUtf8WideLine(const std::wstring& ws) {
  PrintUtf8Wide(ws);
  PrintUtf8("\n");
}

void ClearConsoleScreen() {
  HANDLE h = StdOut();
  if (h == INVALID_HANDLE_VALUE) return;
  CONSOLE_SCREEN_BUFFER_INFO csbi{};
  if (!GetConsoleScreenBufferInfo(h, &csbi)) return;
  const DWORD cells = static_cast<DWORD>(csbi.dwSize.X) * static_cast<DWORD>(csbi.dwSize.Y);
  COORD home{0, 0};
  DWORD written = 0;
  FillConsoleOutputCharacterW(h, L' ', cells, home, &written);
  FillConsoleOutputAttribute(h, csbi.wAttributes, cells, home, &written);
  SetConsoleCursorPosition(h, home);
}

bool ReadConsoleLineWide(std::wstring& out, bool& eof, bool& overflowed, std::wstring& err) {
  out.clear();
  eof = false;
  overflowed = false;
  HANDLE hin = StdIn();
  if (hin == INVALID_HANDLE_VALUE) {
    err = L"Invalid standard input handle.";
    return false;
  }

  constexpr size_t kMax = 2048;
  std::wstring acc;
  bool had_overflow = false;

  for (;;) {
    wchar_t buf[1024];
    DWORD read = 0;
    if (!ReadConsoleW(hin, buf, static_cast<DWORD>(std::size(buf)), &read, nullptr)) {
      err = L"ReadConsoleW failed.";
      return false;
    }
    if (read == 0) {
      eof = true;
      break;
    }

    size_t i = 0;
    while (i < read) {
      wchar_t c = buf[i++];
      if (c == L'\r') continue;
      if (c == L'\n') {
        if (had_overflow) {
          overflowed = true;
          out.clear();
        } else {
          out = std::move(acc);
        }
        return true;
      }
      if (acc.size() < kMax) {
        acc.push_back(c);
      } else {
        had_overflow = true;
      }
    }
  }

  if (had_overflow) {
    overflowed = true;
    out.clear();
  } else {
    out = std::move(acc);
  }
  return true;
}

bool ReadConsoleCharYesNo(bool& yes, std::wstring& err) {
  std::wstring line;
  bool eof = false;
  bool overflow = false;
  if (!ReadConsoleLineWide(line, eof, overflow, err)) return false;
  if (overflow) {
    PrintUtf8("Input line too long.\n");
    yes = false;
    return true;
  }
  if (eof || line.empty()) {
    yes = false;
    return true;
  }
  wchar_t c = static_cast<wchar_t>(std::towlower(line[0]));
  yes = (c == L'y');
  return true;
}

}  // namespace adosql
