#include "ado_session.hpp"
#include "cli.hpp"
#include "console.hpp"
#include "ini_settings.hpp"
#include "paths.hpp"
#include "repl.hpp"
#include "settings.hpp"

#include <windows.h>

#include <objbase.h>
#include <optional>

int wmain(int argc, wchar_t* argv[]) {
  SetConsoleCP(65001);
  SetConsoleOutputCP(65001);

  const std::optional<adosql::CliResult> opt = adosql::ParseCli(argc, argv);
  if (!opt) return 1;
  if (opt->action != adosql::CliAction::Run) return 0;

  const HRESULT hrCom = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
  if (FAILED(hrCom)) {
    adosql::PrintUtf8("Failed to initialize COM.\n");
    return 1;
  }

  {
    adosql::DisplaySettings settings;
    std::wstring ierr;
    adosql::LoadIniNextToExe(settings, ierr);

    adosql::AdoSession session;
    std::wstring db = opt->db_path;

    if (!adosql::FileExists(db)) {
      adosql::PrintUtf8Wide(L"Database file not found:\n");
      adosql::PrintUtf8WideLine(db);
      adosql::PrintUtf8("Create a new database? [y/n]: ");
      bool yes = false;
      if (!adosql::ReadConsoleCharYesNo(yes, ierr)) {
        adosql::PrintUtf8WideLine(L"Error: " + ierr);
        CoUninitialize();
        return 1;
      }
      if (!yes) {
        CoUninitialize();
        return 1;
      }
      if (!session.CreateDatabase(db, ierr)) {
        adosql::PrintUtf8WideLine(L"Error: " + ierr);
        CoUninitialize();
        return 1;
      }
    }

    std::wstring err;
    if (!session.Connect(db, opt->db_password.has_value() ? &*opt->db_password : nullptr, err)) {
      adosql::PrintUtf8WideLine(L"Error: " + err);
      CoUninitialize();
      return 1;
    }

    adosql::RunRepl(session, settings);
  }

  CoUninitialize();
  return 0;
}
