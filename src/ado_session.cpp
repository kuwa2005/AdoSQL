#include "ado_session.hpp"

#include "console.hpp"

#include <algorithm>
#include <cwchar>
#include <cwctype>
#include <iomanip>
#include <sstream>
#include <string_view>
#include <utility>
#include <vector>

#ifdef _WIN64
#import "C:\\Program Files\\Common Files\\System\\ado\\msado15.dll" rename("EOF", "ADO_EOF") rename_namespace("AdoDb")
#import "C:\\Program Files\\Common Files\\System\\ado\\msadox.dll" rename("EOF", "ADOX_EOF") rename_namespace("AdoX")
#else
#import "C:\\Program Files (x86)\\Common Files\\System\\ado\\msado15.dll" rename("EOF", "ADO_EOF") rename_namespace("AdoDb")
#import "C:\\Program Files (x86)\\Common Files\\System\\ado\\msadox.dll" rename("EOF", "ADOX_EOF") rename_namespace("AdoX")
#endif

namespace adosql {

namespace {
constexpr long kAdStateClosed = 0;
constexpr long kAdConnectUnspecified = -1;
constexpr long kAdCmdText = 1;
constexpr AdoDb::SchemaEnum kAdSchemaTables = static_cast<AdoDb::SchemaEnum>(20);
constexpr AdoDb::SchemaEnum kAdSchemaViews = static_cast<AdoDb::SchemaEnum>(23);
constexpr AdoDb::SchemaEnum kAdSchemaColumns = static_cast<AdoDb::SchemaEnum>(4);

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

bool LooksLikeMsys(std::wstring_view name) {
  return name.size() >= 4 && std::towlower(name[0]) == L'm' && std::towlower(name[1]) == L's' &&
         std::towlower(name[2]) == L'y' && std::towlower(name[3]) == L's';
}

std::wstring FormatComError(const _com_error& e) {
  std::wstring msg;
  const wchar_t* d = e.Description();
  if (d && *d) msg = d;
  else msg = e.ErrorMessage();
  msg += L" [0x";
  wchar_t buf[32];
  swprintf_s(buf, L"%08lX", static_cast<unsigned long>(e.Error()));
  msg += buf;
  msg += L"]";
  return msg;
}

std::wstring MetaStr(const _variant_t& v) {
  if (v.vt == VT_NULL || v.vt == VT_EMPTY) return {};
  try {
    _variant_t tmp;
    tmp.ChangeType(VT_BSTR, &v);
    if (tmp.vt == VT_BSTR && tmp.bstrVal) return std::wstring(tmp.bstrVal, SysStringLen(tmp.bstrVal));
  } catch (...) {
  }
  return {};
}

std::wstring VariantToText(const _variant_t& v) {
  if (v.vt == VT_NULL || v.vt == VT_EMPTY) return L"NULL";
  try {
    _variant_t tmp;
    tmp.ChangeType(VT_BSTR, &v);
    if (tmp.vt == VT_BSTR && tmp.bstrVal) {
      std::wstring s(tmp.bstrVal, SysStringLen(tmp.bstrVal));
      std::wstring out;
      out.reserve(s.size() * 2);
      for (wchar_t c : s) {
        if (c == L'\r')
          out += L"<cr>";
        else if (c == L'\n')
          out += L"<lf>";
        else
          out.push_back(c);
      }
      for (size_t pos = 0;;) {
        const size_t cr = out.find(L"<cr>", pos);
        const size_t lf = out.find(L"<lf>", pos);
        if (cr != std::wstring::npos && lf != std::wstring::npos && lf == cr + 4) {
          out.replace(cr, 8, L"<crlf>");
          pos = cr + 6;
          continue;
        }
        break;
      }
      return out;
    }
  } catch (...) {
  }
  try {
    _variant_t n;
    n.ChangeType(VT_R8, &v);
    std::wostringstream os;
    os << std::fixed << std::setprecision(6) << n.dblVal;
    return os.str();
  } catch (...) {
  }
  return L"?";
}

std::wstring BuildConnStr(const wchar_t* provider, const std::wstring& path, const std::wstring* password) {
  std::wstring s = L"Provider=";
  s += provider;
  s += L";Data Source=";
  s += path;
  if (password && !password->empty()) {
    s += L";Jet OLEDB:Database Password=";
    s += *password;
  }
  return s;
}

std::wstring Repeat(wchar_t ch, int n) {
  if (n <= 0) return {};
  return std::wstring(static_cast<size_t>(n), ch);
}

std::wstring FitCell(const std::wstring& s, int width) {
  if (width <= 0) return {};
  if (static_cast<int>(s.size()) <= width) return s;
  if (width <= 3) return s.substr(0, static_cast<size_t>(width));
  return s.substr(0, static_cast<size_t>(width - 3)) + L"...";
}

std::wstring PadCell(const std::wstring& s, int width, bool right_align) {
  const std::wstring f = FitCell(s, width);
  const int pad = std::max(0, width - static_cast<int>(f.size()));
  if (right_align) return Repeat(L' ', pad) + f;
  return f + Repeat(L' ', pad);
}

void FitWidthsToLine(std::vector<int>& widths, int line_width) {
  if (widths.empty()) return;
  const int sep = static_cast<int>(widths.size() - 1) * 3;  // " | "
  auto sum = [&]() {
    int s = 0;
    for (int w : widths) s += w;
    return s + sep;
  };
  int total = sum();
  while (total > line_width) {
    size_t idx = 0;
    for (size_t i = 1; i < widths.size(); ++i) {
      if (widths[i] > widths[idx]) idx = i;
    }
    if (widths[idx] <= 6) break;
    widths[idx]--;
    total = sum();
  }
}

void PrintHeader(const std::vector<std::wstring>& headers, const std::vector<int>& widths) {
  std::wostringstream h;
  std::wostringstream u;
  for (size_t i = 0; i < headers.size(); ++i) {
    if (i) {
      h << L" | ";
      u << L"-+-";
    }
    h << PadCell(headers[i], widths[i], false);
    u << Repeat(L'-', widths[i]);
  }
  PrintUtf8WideLine(h.str());
  PrintUtf8WideLine(u.str());
}

bool IsNumericDataType(long ado_type) {
  switch (ado_type) {
    case 2:   // SmallInt
    case 3:   // Integer
    case 4:   // Single
    case 5:   // Double
    case 6:   // Currency
    case 14:  // Decimal
    case 16:  // TinyInt
    case 17:  // UnsignedTinyInt
    case 18:  // UnsignedSmallInt
    case 19:  // UnsignedInt
    case 20:  // BigInt
    case 21:  // UnsignedBigInt
    case 131: // Numeric
    case 139: // VarNumeric
      return true;
    default:
      return false;
  }
}

}  // namespace

struct AdoSession::Impl {
  AdoDb::_ConnectionPtr conn;
  bool txn_open = false;
};

AdoSession::AdoSession() : impl_(std::make_unique<Impl>()) {}
AdoSession::~AdoSession() { Disconnect(); }

void AdoSession::Disconnect() {
  if (!impl_->conn) return;
  try {
    if (impl_->txn_open) {
      impl_->conn->RollbackTrans();
      impl_->txn_open = false;
    }
    if (impl_->conn->GetState() != kAdStateClosed) impl_->conn->Close();
  } catch (...) {
  }
  impl_->conn = nullptr;
}

bool AdoSession::InTransaction() const { return impl_->txn_open; }

bool AdoSession::Connect(const std::wstring& db_path, const std::wstring* password, std::wstring& error) {
  error.clear();
  Disconnect();

  const wchar_t* ace[] = {L"Microsoft.ACE.OLEDB.16.0", L"Microsoft.ACE.OLEDB.12.0"};
  _com_error last_err(0);
  for (const wchar_t* prov : ace) {
    try {
      impl_->conn.CreateInstance(__uuidof(AdoDb::Connection));
      if (!impl_->conn) {
        error = L"Failed to create ADO Connection object.";
        return false;
      }
      std::wstring cs = BuildConnStr(prov, db_path, password);
      impl_->conn->Open(_bstr_t(cs.c_str()), L"", L"", kAdConnectUnspecified);
      impl_->txn_open = false;
      return true;
    } catch (_com_error& e) {
      last_err = e;
      impl_->conn = nullptr;
    }
  }

#ifndef _WIN64
  if (db_path.size() >= 4) {
    std::wstring ext = db_path.substr(db_path.size() - 4);
    for (auto& c : ext) c = static_cast<wchar_t>(std::towlower(c));
    if (ext == L".mdb") {
      try {
        impl_->conn.CreateInstance(__uuidof(AdoDb::Connection));
        std::wstring cs = BuildConnStr(L"Microsoft.Jet.OLEDB.4.0", db_path, password);
        impl_->conn->Open(_bstr_t(cs.c_str()), L"", L"", kAdConnectUnspecified);
        impl_->txn_open = false;
        return true;
      } catch (_com_error& e) {
        last_err = e;
        impl_->conn = nullptr;
      }
    }
  }
#endif

  error = L"Could not connect (ACE / provider). Install Microsoft Access Database Engine (ACE) matching this "
          L"executable's bitness.\nDetails: ";
  error += FormatComError(last_err);
  return false;
}

bool AdoSession::CreateDatabase(const std::wstring& db_path, std::wstring& error) {
  error.clear();
  const wchar_t* ace[] = {L"Microsoft.ACE.OLEDB.16.0", L"Microsoft.ACE.OLEDB.12.0"};
  _com_error last_err(0);
  for (const wchar_t* prov : ace) {
    try {
      AdoX::_CatalogPtr cat(__uuidof(AdoX::Catalog));
      std::wstring cs = BuildConnStr(prov, db_path, nullptr);
      cat->Create(_bstr_t(cs.c_str()));
      return true;
    } catch (_com_error& e) {
      last_err = e;
    }
  }
#ifndef _WIN64
  if (db_path.size() >= 4) {
    std::wstring ext = db_path.substr(db_path.size() - 4);
    for (auto& c : ext) c = static_cast<wchar_t>(std::towlower(c));
    if (ext == L".mdb") {
      try {
        AdoX::_CatalogPtr cat(__uuidof(AdoX::Catalog));
        std::wstring cs = BuildConnStr(L"Microsoft.Jet.OLEDB.4.0", db_path, nullptr);
        cat->Create(_bstr_t(cs.c_str()));
        return true;
      } catch (_com_error& e) {
        last_err = e;
      }
    }
  }
#endif
  error = L"Failed to create database: ";
  error += FormatComError(last_err);
  return false;
}

void AdoSession::Commit(std::wstring& err) {
  err.clear();
  if (!impl_->conn) {
    err = L"Not connected.";
    return;
  }
  try {
    if (impl_->txn_open) {
      impl_->conn->CommitTrans();
      impl_->txn_open = false;
    }
  } catch (_com_error& e) {
    err = FormatComError(e);
  }
}

void AdoSession::Rollback(std::wstring& err) {
  err.clear();
  if (!impl_->conn) {
    err = L"Not connected.";
    return;
  }
  try {
    if (impl_->txn_open) {
      impl_->conn->RollbackTrans();
      impl_->txn_open = false;
    }
  } catch (_com_error& e) {
    err = FormatComError(e);
  }
}

void AdoSession::OnExitSession() {
  std::wstring ignored;
  Rollback(ignored);
}

void AdoSession::EnsureTxnForDml(const std::wstring& sql, std::wstring& err) {
  err.clear();
  std::wstring t = FirstTokenLower(sql);
  if (t != L"insert" && t != L"update" && t != L"delete") return;
  if (!impl_->conn) {
    err = L"Not connected.";
    return;
  }
  if (impl_->txn_open) return;
  try {
    impl_->conn->BeginTrans();
    impl_->txn_open = true;
  } catch (_com_error& e) {
    err = FormatComError(e);
  }
}

static void PrintRecordset(AdoDb::_RecordsetPtr rs, const DisplaySettings& settings) {
  if (!rs || rs->GetState() == kAdStateClosed) return;

  const int effective_width = GetConsoleWidthOrDefault(120);
  const int linesize = std::min(std::clamp(settings.linesize, 20, 500), effective_width);

  auto fields = rs->GetFields();
  const long n = fields->GetCount();
  std::vector<std::wstring> names;
  std::vector<int> widths;
  std::vector<bool> right_align;
  names.reserve(static_cast<size_t>(n));
  widths.reserve(static_cast<size_t>(n));
  right_align.reserve(static_cast<size_t>(n));
  for (long i = 0; i < n; ++i) {
    auto f = fields->Item[_variant_t(i)];
    std::wstring nm = static_cast<const wchar_t*>(f->GetName());
    names.push_back(nm);
    int w = static_cast<int>(nm.size());
    long def_size = 0;
    try {
      def_size = f->GetDefinedSize();
    } catch (...) {
      def_size = 0;
    }
    if (def_size > 0) w = std::max(w, std::min(40, static_cast<int>(def_size)));
    w = std::clamp(w, 6, 40);
    widths.push_back(w);
    right_align.push_back(IsNumericDataType(static_cast<long>(f->GetType())));
  }

  std::vector<long> visible;
  visible.reserve(static_cast<size_t>(n));
  if (settings.widthclamp) {
    int budget = linesize;
    for (long i = 0; i < n; ++i) {
      int need = static_cast<int>(widths[static_cast<size_t>(i)]);
      if (!visible.empty()) need += 3;
      if (need > budget) break;
      visible.push_back(i);
      budget -= need;
    }
  } else {
    for (long i = 0; i < n; ++i) visible.push_back(i);
  }
  if (!visible.empty()) {
    std::vector<int> vw;
    vw.reserve(visible.size());
    for (long idx : visible) vw.push_back(widths[static_cast<size_t>(idx)]);
    FitWidthsToLine(vw, linesize);
    for (size_t i = 0; i < visible.size(); ++i) widths[static_cast<size_t>(visible[i])] = vw[i];
  }

  {
    std::vector<std::wstring> headers;
    std::vector<int> hw;
    headers.reserve(visible.size());
    hw.reserve(visible.size());
    for (long idx : visible) {
      headers.push_back(names[static_cast<size_t>(idx)]);
      hw.push_back(widths[static_cast<size_t>(idx)]);
    }
    PrintHeader(headers, hw);
  }

  long row_in_page = 0;
  long total_since_more = 0;
  while (!rs->GetADO_EOF()) {
    std::wostringstream line;
    for (size_t vi = 0; vi < visible.size(); ++vi) {
      if (vi) line << L" | ";
      long i = visible[vi];
      auto f = fields->Item[_variant_t(i)];
      _variant_t v = f->GetValue();
      std::wstring cell = VariantToText(v);
      line << PadCell(cell, widths[static_cast<size_t>(i)], right_align[static_cast<size_t>(i)]);
    }
    PrintUtf8WideLine(line.str());

    rs->MoveNext();
    row_in_page++;
    total_since_more++;
    if (settings.pagesize > 0 && row_in_page >= settings.pagesize) {
      row_in_page = 0;
      PrintUtf8("\n");
    }
    if (settings.maxrows > 0 && total_since_more >= settings.maxrows) {
      PrintUtf8("--More--\n");
      std::wstring dummy;
      bool eof = false;
      bool ovf = false;
      std::wstring err;
      ReadConsoleLineWide(dummy, eof, ovf, err);
      if (!dummy.empty() && (dummy[0] == L'q' || dummy[0] == L'Q')) break;
      total_since_more = 0;
    }
  }
}

bool AdoSession::ExecuteInteractive(const std::wstring& sql_in, DisplaySettings& settings, std::wstring& err) {
  err.clear();
  if (!impl_->conn) {
    err = L"Not connected.";
    return false;
  }
  std::wstring sql = sql_in;
  Trim(sql);
  if (sql.empty()) return true;

  std::wstring ft = FirstTokenLower(sql);
  const bool is_select = (ft == L"select") || (ft == L"with");

  try {
    if (is_select) {
      _variant_t affected;
      AdoDb::_RecordsetPtr rs = impl_->conn->Execute(_bstr_t(sql.c_str()), &affected, kAdCmdText);
      PrintRecordset(rs, settings);
      return true;
    }

    EnsureTxnForDml(sql, err);
    if (!err.empty()) return false;

    _variant_t ra;
    impl_->conn->Execute(_bstr_t(sql.c_str()), &ra, kAdCmdText);
    return true;
  } catch (_com_error& e) {
    err = FormatComError(e);
    return false;
  }
}

bool AdoSession::RunTableList(std::wstring& err) {
  err.clear();
  if (!impl_->conn) {
    err = L"Not connected.";
    return false;
  }
  try {
    AdoDb::_RecordsetPtr rs = impl_->conn->OpenSchema(kAdSchemaTables);
    int count = 0;
    std::vector<std::pair<std::wstring, std::wstring>> rows;
    rows.reserve(64);
    while (!rs->GetADO_EOF()) {
      _variant_t name = rs->Fields->GetItem(L"TABLE_NAME")->GetValue();
      _variant_t type = rs->Fields->GetItem(L"TABLE_TYPE")->GetValue();
      std::wstring n = MetaStr(name);
      if (n.empty()) {
        rs->MoveNext();
        continue;
      }
      if (LooksLikeMsys(n)) {
        rs->MoveNext();
        continue;
      }
      std::wstring t = MetaStr(type);
      rows.emplace_back(n, t);
      count++;
      if (count >= 255) {
        PrintUtf8("Error: metadata enumeration limit (255) exceeded.\n");
        return false;
      }
      rs->MoveNext();
    }
    const int cw = GetConsoleWidthOrDefault(120);
    int w1 = 10, w2 = 10;
    for (const auto& r : rows) {
      w1 = std::max(w1, static_cast<int>(r.first.size()));
      w2 = std::max(w2, static_cast<int>(r.second.size()));
    }
    w1 = std::clamp(w1, 8, 60);
    w2 = std::clamp(w2, 8, 30);
    std::vector<int> ws{w1, w2};
    FitWidthsToLine(ws, cw);
    PrintHeader({L"TABLE_NAME", L"TABLE_TYPE"}, ws);
    for (const auto& r : rows) {
      PrintUtf8WideLine(PadCell(r.first, ws[0], false) + L" | " + PadCell(r.second, ws[1], false));
    }
    return true;
  } catch (_com_error& e) {
    err = FormatComError(e);
    return false;
  }
}

bool AdoSession::RunQueryList(std::wstring& err) {
  err.clear();
  if (!impl_->conn) {
    err = L"Not connected.";
    return false;
  }
  try {
    AdoDb::_RecordsetPtr rs = impl_->conn->OpenSchema(kAdSchemaViews);
    int count = 0;
    std::vector<std::wstring> rows;
    rows.reserve(64);
    while (!rs->GetADO_EOF()) {
      _variant_t name = rs->Fields->GetItem(L"TABLE_NAME")->GetValue();
      std::wstring n = MetaStr(name);
      if (n.empty()) {
        rs->MoveNext();
        continue;
      }
      rows.push_back(n);
      count++;
      if (count >= 255) {
        PrintUtf8("Error: metadata enumeration limit (255) exceeded.\n");
        return false;
      }
      rs->MoveNext();
    }
    const int cw = GetConsoleWidthOrDefault(120);
    int w = 10;
    for (const auto& r : rows) w = std::max(w, static_cast<int>(r.size()));
    w = std::clamp(w, 8, 100);
    std::vector<int> ws{w};
    FitWidthsToLine(ws, cw);
    PrintHeader({L"QUERY_NAME"}, ws);
    for (const auto& r : rows) PrintUtf8WideLine(PadCell(r, ws[0], false));
    return true;
  } catch (_com_error& e) {
    err = FormatComError(e);
    return false;
  }
}

namespace {

bool DumpAdoxColumns(AdoX::ColumnsPtr cols, std::wstring& err) {
  err.clear();
  if (!cols) {
    err = L"No column metadata.";
    return false;
  }
  struct Row {
    std::wstring name;
    std::wstring type;
    std::wstring size;
  };
  std::vector<Row> rows;
  rows.reserve(64);
  int cnt = 0;
  const long c = cols->GetCount();
  for (long i = 0; i < c; ++i) {
    AdoX::_ColumnPtr col = cols->Item[_variant_t(i)];
    std::wstring cn = static_cast<const wchar_t*>(col->GetName());
    const long dtype = static_cast<long>(col->GetType());
    long dsize = 0;
    try {
      dsize = col->GetDefinedSize();
    } catch (...) {
      dsize = 0;
    }
    rows.push_back({cn, std::to_wstring(dtype), std::to_wstring(dsize)});
    cnt++;
    if (cnt >= 255) {
      PrintUtf8("Error: metadata enumeration limit (255) exceeded.\n");
      return false;
    }
  }
  const int cw = GetConsoleWidthOrDefault(120);
  int w1 = 8, w2 = 8, w3 = 8;
  for (const auto& r : rows) {
    w1 = std::max(w1, static_cast<int>(r.name.size()));
    w2 = std::max(w2, static_cast<int>(r.type.size()));
    w3 = std::max(w3, static_cast<int>(r.size.size()));
  }
  w1 = std::clamp(w1, 8, 60);
  w2 = std::clamp(w2, 8, 16);
  w3 = std::clamp(w3, 8, 16);
  std::vector<int> ws{w1, w2, w3};
  FitWidthsToLine(ws, cw);
  PrintHeader({L"NAME", L"TYPE", L"DEFINED_SIZE"}, ws);
  for (const auto& r : rows) {
    PrintUtf8WideLine(PadCell(r.name, ws[0], false) + L" | " + PadCell(r.type, ws[1], true) + L" | " +
                      PadCell(r.size, ws[2], true));
  }
  return true;
}

}  // namespace

bool AdoSession::RunDescribe(const std::wstring& name, std::wstring& err) {
  err.clear();
  if (!impl_->conn) {
    err = L"Not connected.";
    return false;
  }
  try {
    AdoX::_CatalogPtr cat(__uuidof(AdoX::Catalog));
    cat->PutActiveConnection(_variant_t(static_cast<IDispatch*>(impl_->conn)));

    try {
      AdoX::_TablePtr tbl = cat->Tables->GetItem(_variant_t(name.c_str()));
      return DumpAdoxColumns(tbl->GetColumns(), err);
    } catch (...) {
      // Fallback: query column schema for view/query names.
      AdoDb::_RecordsetPtr rs = impl_->conn->OpenSchema(kAdSchemaColumns);
      std::vector<std::pair<std::wstring, std::wstring>> rows;
      rows.reserve(64);
      int cnt = 0;
      while (!rs->GetADO_EOF()) {
        std::wstring tn = MetaStr(rs->Fields->GetItem(L"TABLE_NAME")->GetValue());
        if (!tn.empty() && _wcsicmp(tn.c_str(), name.c_str()) == 0) {
          std::wstring cn = MetaStr(rs->Fields->GetItem(L"COLUMN_NAME")->GetValue());
          std::wstring dt = MetaStr(rs->Fields->GetItem(L"DATA_TYPE")->GetValue());
          rows.emplace_back(cn, dt);
          cnt++;
          if (cnt >= 255) {
            PrintUtf8("Error: metadata enumeration limit (255) exceeded.\n");
            return false;
          }
        }
        rs->MoveNext();
      }
      if (cnt == 0) {
        err = L"Object not found.";
        return false;
      }
      const int cw = GetConsoleWidthOrDefault(120);
      int w1 = 11, w2 = 9;
      for (const auto& r : rows) {
        w1 = std::max(w1, static_cast<int>(r.first.size()));
        w2 = std::max(w2, static_cast<int>(r.second.size()));
      }
      w1 = std::clamp(w1, 8, 60);
      w2 = std::clamp(w2, 8, 24);
      std::vector<int> ws{w1, w2};
      FitWidthsToLine(ws, cw);
      PrintHeader({L"COLUMN_NAME", L"DATA_TYPE"}, ws);
      for (const auto& r : rows) {
        PrintUtf8WideLine(PadCell(r.first, ws[0], false) + L" | " + PadCell(r.second, ws[1], true));
      }
      return true;
    }
  } catch (_com_error& e) {
    err = FormatComError(e);
    return false;
  }
}

bool AdoSession::RunShowSql(const std::wstring& name, std::wstring& err) {
  err.clear();
  if (!impl_->conn) {
    err = L"Not connected.";
    return false;
  }
  try {
    AdoDb::_RecordsetPtr rs = impl_->conn->OpenSchema(kAdSchemaViews);
    while (!rs->GetADO_EOF()) {
      std::wstring n = MetaStr(rs->Fields->GetItem(L"TABLE_NAME")->GetValue());
      if (!n.empty() && _wcsicmp(n.c_str(), name.c_str()) == 0) {
        std::wstring def = MetaStr(rs->Fields->GetItem(L"VIEW_DEFINITION")->GetValue());
        if (def.empty()) def = L"<definition unavailable>";
        PrintUtf8WideLine(def);
        return true;
      }
      rs->MoveNext();
    }
    err = L"Query definition not found.";
    return false;
  } catch (_com_error& e) {
    err = FormatComError(e);
    return false;
  }
}

}  // namespace adosql
