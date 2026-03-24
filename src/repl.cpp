#include "repl.hpp"

#include "ado_session.hpp"
#include "console.hpp"
#include "ini_settings.hpp"
#include "script_input.hpp"
#include "settings.hpp"

#include <algorithm>
#include <cwctype>
#include <sstream>
#include <string_view>
#include <unordered_set>
#include <vector>

namespace adosql {

namespace {

void Trim(std::wstring& s) {
  while (!s.empty() && std::iswspace(s.front())) s.erase(s.begin());
  while (!s.empty() && std::iswspace(s.back())) s.pop_back();
}

std::wstring FirstToken(std::wstring_view sv) {
  size_t i = 0;
  while (i < sv.size() && std::iswspace(sv[i])) ++i;
  size_t j = i;
  while (j < sv.size() && !std::iswspace(sv[j])) ++j;
  return std::wstring(sv.substr(i, j - i));
}

std::wstring FirstTokenLower(std::wstring_view sv) {
  std::wstring t = FirstToken(sv);
  for (wchar_t& c : t) c = static_cast<wchar_t>(std::towlower(c));
  return t;
}

bool ReadOneLine(ScriptStack& scripts, std::wstring& line, bool& eof, bool& overflow, std::wstring& err) {
  if (scripts.InScript()) {
    bool seof = false;
    if (!scripts.ReadLine(line, seof, overflow, err)) return false;
    if (overflow) return true;
    if (seof) {
      scripts.Pop();
      return ReadOneLine(scripts, line, eof, overflow, err);
    }
    return true;
  }
  return ReadConsoleLineWide(line, eof, overflow, err);
}

void PrintHelp() {
  PrintUtf8(
      "Commands (first word, case-insensitive):\n"
      "  exit | quit          Leave (rolls back uncommitted DML)\n"
      "  help | ? | /?        This help\n"
      "  // ...               Comment (line starting with //)\n"
      "  @path                Run script file (UTF-8, relative to CWD)\n"
      "  set pagesize N       5..500\n"
      "  set linesize N       20..500\n"
      "  set maxrows N        0 = unlimited; >0 pauses with --More--\n"
      "  set widthclamp on|off\n"
      "  show settings\n"
      "  prompt TEXT\n"
      "  tablelist | querylist | describe NAME | showsql NAME\n"
      "  commit | rollback\n"
      "  SQL                  select/insert/update/delete/create/alter/drop/...\n"
      "\n"
      "Input: end a statement with ';'. Empty line cancels a multi-line buffer.\n"
      "Tab completion: planned (see docs).\n"
      "Tip: use AS aliases on expressions in SELECT.\n");
}

bool IsKnownFirstToken(const std::wstring& ft) {
  static const std::unordered_set<std::wstring> k = {
      L"exit",    L"quit",     L"help",     L"?",        L"/?",        L"set",        L"show",
      L"prompt",  L"select",   L"insert",   L"update",   L"delete",    L"create",     L"alter",
      L"drop",    L"commit",   L"rollback", L"describe", L"descrive",  L"desc",       L"showsql",    L"tablelist",
      L"querylist", L"with",   L"cls",      L"clear",
  };
  return k.find(ft) != k.end();
}

bool HandleSlashSlashLine(std::wstring_view line) {
  size_t i = 0;
  while (i < line.size() && std::iswspace(line[i])) ++i;
  return line.size() >= i + 2 && line[i] == L'/' && line[i + 1] == L'/';
}

bool ParseSet(const std::wstring& stmt, DisplaySettings& settings, std::wstring& err) {
  err.clear();
  std::wistringstream iss(stmt);
  std::wstring a, b, c, d;
  iss >> a >> b;
  if (!iss) {
    err = L"Invalid set syntax.";
    return false;
  }
  for (auto& ch : a) ch = static_cast<wchar_t>(std::towlower(ch));
  for (auto& ch : b) ch = static_cast<wchar_t>(std::towlower(ch));
  if (a != L"set") {
    err = L"Invalid set syntax.";
    return false;
  }
  if (b == L"pagesize") {
    iss >> c;
    try {
      settings.pagesize = std::stoi(c);
    } catch (...) {
      err = L"Invalid number for pagesize.";
      return false;
    }
    settings.pagesize = std::clamp(settings.pagesize, 5, 500);
    return true;
  }
  if (b == L"linesize") {
    iss >> c;
    try {
      settings.linesize = std::stoi(c);
    } catch (...) {
      err = L"Invalid number for linesize.";
      return false;
    }
    settings.linesize = std::clamp(settings.linesize, 20, 500);
    return true;
  }
  if (b == L"maxrows") {
    iss >> c;
    try {
      settings.maxrows = std::stoi(c);
    } catch (...) {
      err = L"Invalid number for maxrows.";
      return false;
    }
    if (settings.maxrows < 0) settings.maxrows = 0;
    return true;
  }
  if (b == L"widthclamp") {
    iss >> c;
    for (auto& ch : c) ch = static_cast<wchar_t>(std::towlower(ch));
    if (c == L"on")
      settings.widthclamp = true;
    else if (c == L"off")
      settings.widthclamp = false;
    else {
      err = L"widthclamp expects on or off.";
      return false;
    }
    return true;
  }
  err = L"Unknown set target.";
  return false;
}

bool IsShowSettings(const std::wstring& stmt) {
  std::wistringstream iss(stmt);
  std::wstring a, b;
  iss >> a >> b;
  for (auto& ch : a) ch = static_cast<wchar_t>(std::towlower(ch));
  for (auto& ch : b) ch = static_cast<wchar_t>(std::towlower(ch));
  return a == L"show" && b == L"settings";
}

bool IsImmediateCommand(const std::wstring& stmt) {
  const std::wstring ft = FirstTokenLower(stmt);
  if (ft.empty()) return false;
  if (ft == L"help" || ft == L"?" || ft == L"/?" || ft == L"tablelist" || ft == L"querylist" || ft == L"describe" ||
      ft == L"descrive" || ft == L"desc" || ft == L"showsql" || ft == L"show" || ft == L"set" || ft == L"prompt" || ft == L"commit" ||
      ft == L"rollback" || ft == L"exit" || ft == L"quit" || ft == L"cls" || ft == L"clear") {
    return true;
  }
  if (!stmt.empty() && stmt[0] == L'@') return true;
  if (HandleSlashSlashLine(stmt)) return true;
  return false;
}

void DispatchStatement(const std::wstring& stmt_in, AdoSession& session, DisplaySettings& settings, ScriptStack& scripts,
                       bool& quit) {
  quit = false;
  std::wstring stmt = stmt_in;
  Trim(stmt);
  if (stmt.empty()) return;
  if (HandleSlashSlashLine(stmt)) return;

  if (!stmt.empty() && stmt[0] == L'@') {
    std::wstring p = stmt.substr(1);
    Trim(p);
    std::wstring err;
    if (!scripts.Push(p, err)) PrintUtf8WideLine(L"Error: " + err);
    return;
  }

  const std::wstring ft = FirstTokenLower(stmt);

  if (ft == L"exit" || ft == L"quit") {
    quit = true;
    return;
  }
  if (ft == L"help" || ft == L"?" || ft == L"/?" || stmt == L"/?") {
    PrintHelp();
    return;
  }

  if (ft == L"set") {
    std::wstring err;
    if (!ParseSet(stmt, settings, err)) {
      PrintUtf8WideLine(L"Error: " + err);
      return;
    }
    std::wstring ierr;
    if (!SaveIniNextToExe(settings, ierr)) PrintUtf8WideLine(L"Warning: " + ierr);
    return;
  }

  if (IsShowSettings(stmt)) {
    std::wostringstream os;
    os << L"pagesize=" << settings.pagesize << L" linesize=" << settings.linesize << L" maxrows=" << settings.maxrows
       << L" widthclamp=" << (settings.widthclamp ? L"on" : L"off");
    PrintUtf8WideLine(os.str());
    return;
  }

  if (ft == L"prompt") {
    size_t i = 0;
    while (i < stmt.size() && std::iswspace(stmt[i])) ++i;
    std::wstring rest = stmt.substr(i);
    const std::wstring tok = FirstToken(rest);
    if (tok.size() < rest.size()) {
      rest = rest.substr(tok.size());
      Trim(rest);
    } else {
      rest.clear();
    }
    PrintUtf8WideLine(rest);
    return;
  }

  if (ft == L"cls" || ft == L"clear") {
    ClearConsoleScreen();
    return;
  }

  if (ft == L"commit") {
    std::wstring err;
    session.Commit(err);
    if (!err.empty()) PrintUtf8WideLine(L"Error: " + err);
    return;
  }
  if (ft == L"rollback") {
    std::wstring err;
    session.Rollback(err);
    if (!err.empty()) PrintUtf8WideLine(L"Error: " + err);
    return;
  }

  if (ft == L"tablelist") {
    std::wstring err;
    if (!session.RunTableList(err)) PrintUtf8WideLine(L"Error: " + err);
    return;
  }
  if (ft == L"querylist") {
    std::wstring err;
    if (!session.RunQueryList(err)) PrintUtf8WideLine(L"Error: " + err);
    return;
  }

  if (ft == L"describe" || ft == L"descrive" || ft == L"desc") {
    std::wistringstream iss(stmt);
    std::wstring a, name;
    iss >> a >> name;
    if (name.empty()) {
      PrintUtf8("Error: describe requires a name.\n");
      return;
    }
    std::wstring err;
    if (!session.RunDescribe(name, err)) PrintUtf8WideLine(L"Error: " + err);
    return;
  }

  if (ft == L"showsql") {
    std::wistringstream iss(stmt);
    std::wstring a, name;
    iss >> a >> name;
    if (name.empty()) {
      PrintUtf8("Error: showsql requires a name.\n");
      return;
    }
    std::wstring err;
    if (!session.RunShowSql(name, err)) PrintUtf8WideLine(L"Error: " + err);
    return;
  }

  if (!IsKnownFirstToken(ft)) {
    PrintUtf8("Invalid command.\n");
    return;
  }

  std::wstring err;
  if (!session.ExecuteInteractive(stmt, settings, err)) PrintUtf8WideLine(L"Error: " + err);
}

}  // namespace

void RunRepl(AdoSession& session, DisplaySettings& settings) {
  ScriptStack scripts;
  std::wstring acc;
  int depth = 1;

  for (;;) {
    std::wostringstream prompt;
    prompt << (depth == 1 ? L"SQL>" : std::to_wstring(depth) + L">");
    PrintUtf8Wide(prompt.str());

    std::wstring line;
    bool eof = false;
    bool overflow = false;
    std::wstring err;
    if (!ReadOneLine(scripts, line, eof, overflow, err)) {
      PrintUtf8WideLine(L"Error: " + err);
      break;
    }
    if (overflow) {
      PrintUtf8("Input line too long.\n");
      acc.clear();
      depth = 1;
      continue;
    }
    if (eof && line.empty()) {
      session.OnExitSession();
      break;
    }

    Trim(line);
    if (acc.empty() && line.empty()) {
      continue;
    }
    if (!acc.empty() && line.empty()) {
      acc.clear();
      depth = 1;
      continue;
    }

    if (!acc.empty()) acc.push_back(L'\n');
    if (acc.empty() && line.find(L';') == std::wstring::npos && IsImmediateCommand(line)) {
      bool quit = false;
      DispatchStatement(line, session, settings, scripts, quit);
      if (quit) {
        session.OnExitSession();
        return;
      }
      continue;
    }

    acc += line;

    if (acc.find(L';') != std::wstring::npos) {
      std::wstring work = acc;
      acc.clear();
      depth = 1;
      while (!work.empty()) {
        const size_t semi = work.find(L';');
        if (semi == std::wstring::npos) {
          acc = work;
          depth = acc.empty() ? 1 : static_cast<int>(std::count(acc.begin(), acc.end(), L'\n') + 1);
          if (depth < 1) depth = 1;
          break;
        }
        std::wstring chunk = work.substr(0, semi + 1);
        work.erase(0, semi + 1);
        Trim(work);
        std::wstring exec = chunk;
        Trim(exec);
        if (!exec.empty() && exec.back() == L';') exec.pop_back();
        Trim(exec);
        bool quit = false;
        DispatchStatement(exec, session, settings, scripts, quit);
        if (quit) {
          session.OnExitSession();
          return;
        }
      }
      continue;
    }

    depth = static_cast<int>(std::count(acc.begin(), acc.end(), L'\n') + 1);
  }
}

}  // namespace adosql
