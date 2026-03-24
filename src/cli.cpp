#include "cli.hpp"

#include "app_version.hpp"
#include "console.hpp"
#include "paths.hpp"

#include <algorithm>
#include <cwctype>

namespace adosql {

namespace {

bool IcEq(std::wstring_view a, std::wstring_view b) {
  if (a.size() != b.size()) return false;
  for (size_t i = 0; i < a.size(); ++i) {
    if (std::towlower(a[i]) != std::towlower(b[i])) return false;
  }
  return true;
}

}  // namespace

void PrintUsage() {
  PrintUtf8(
      "Usage:\n"
      "  adosql.exe <database_file> [password]\n"
      "  adosql.exe --version | -v\n"
      "  adosql.exe /?\n"
      "\n"
      "Password on the command line may be visible in process lists and shell history.\n");
}

void PrintVersion() {
  PrintUtf8WideLine(std::wstring(kProductName) + L" " + kVersionString);
}

std::optional<CliResult> ParseCli(int argc, wchar_t* argv[]) {
  if (argc <= 1) {
    PrintUsage();
    return CliResult{.action = CliAction::UsageOk};
  }

  const std::wstring_view a1 = argv[1];
  if (IcEq(a1, L"--version") || IcEq(a1, L"-v")) {
    PrintVersion();
    return CliResult{.action = CliAction::VersionOk};
  }
  if (a1 == L"/?" || IcEq(a1, L"-?") || IcEq(a1, L"/help") || IcEq(a1, L"-help")) {
    PrintUsage();
    return CliResult{.action = CliAction::UsageOk};
  }

  CliResult r;
  r.action = CliAction::Run;
  std::wstring err;
  if (!NormalizePathToFull(argv[1], r.db_path, err)) {
    PrintUtf8WideLine(L"Error: " + err);
    return std::nullopt;
  }

  if (argc >= 3) {
    r.db_password = std::wstring(argv[2]);
  }
  return r;
}

}  // namespace adosql
