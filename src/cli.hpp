#pragma once

#include <optional>
#include <string>

namespace adosql {

enum class CliAction {
  UsageOk,       // exit 0
  VersionOk,   // exit 0
  Run,           // connect + repl
};

struct CliResult {
  CliAction action = CliAction::UsageOk;
  std::wstring db_path;           // full path when Run
  std::optional<std::wstring> db_password;
};

// Parses argv. On failure returns nullopt (caller prints generic error, exit 1).
std::optional<CliResult> ParseCli(int argc, wchar_t* argv[]);

void PrintUsage();
void PrintVersion();

}  // namespace adosql
