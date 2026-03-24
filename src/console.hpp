#pragma once

#include <string>

namespace adosql {

void PrintUtf8(const char* utf8);
void PrintUtf8(const std::string& utf8);
void PrintUtf8Wide(const std::wstring& ws);
void PrintUtf8WideLine(const std::wstring& ws);

// Reads one line (max 2048 chars excluding NUL). Discards overflow per spec; sets overflowed.
bool ReadConsoleLineWide(std::wstring& out, bool& eof, bool& overflowed, std::wstring& err);

bool ReadConsoleCharYesNo(bool& yes, std::wstring& err);

}  // namespace adosql
