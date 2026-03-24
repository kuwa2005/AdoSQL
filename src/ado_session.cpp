#include "ado_session.hpp"

#include "console.hpp"

#include <algorithm>
#include <cwctype>
#include <iomanip>
#include <sstream>
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
    if (impl_->conn->GetState() != adStateClosed) impl_->conn->Close();
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
      impl_->conn->Open(_bstr_t(cs.c_str()), L"", L"", adConnectUnspecified);
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
        impl_->conn->Open(_bstr_t(cs.c_str()), L"", L"", adConnectUnspecified);
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
  if (!rs || rs->GetState() == adStateClosed) return;

  const int linesize = std::clamp(settings.linesize, 20, 500);

  auto fields = rs->GetFields();
  const long n = fields->GetCount();
  std::vector<std::wstring> names;
  std::vector<long> widths;
  names.reserve(static_cast<size_t>(n));
  widths.reserve(static_cast<size_t>(n));
  for (long i = 0; i < n; ++i) {
    auto f = fields->Item[_variant_t(i)];
    std::wstring nm = static_cast<const wchar_t*>(f->GetName());
    names.push_back(nm);
    long w = static_cast<long>(nm.size());
    if (w < 8) w = 8;
    widths.push_back(w);
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

  {
    std::wostringstream hdr;
    for (size_t vi = 0; vi < visible.size(); ++vi) {
      if (vi) hdr << L" | ";
      long i = visible[vi];
      std::wstring col = names[static_cast<size_t>(i)];
      if (static_cast<int>(col.size()) > widths[static_cast<size_t>(i)])
        col.resize(static_cast<size_t>(widths[static_cast<size_t>(i)]));
      hdr << col;
    }
    std::wstring hs = hdr.str();
    if (static_cast<int>(hs.size()) > linesize) hs.resize(static_cast<size_t>(linesize));
    PrintUtf8WideLine(hs);
    PrintUtf8WideLine(std::wstring(std::min<size_t>(static_cast<size_t>(linesize), 40), L'-'));
  }

  long row_in_page = 0;
  long total_since_more = 0;
  while (!rs->EndOfFile) {
    std::wostringstream line;
    for (size_t vi = 0; vi < visible.size(); ++vi) {
      if (vi) line << L" | ";
      long i = visible[vi];
      auto f = fields->Item[_variant_t(i)];
      _variant_t v = f->GetValue();
      std::wstring cell = VariantToText(v);
      if (static_cast<int>(cell.size()) > widths[static_cast<size_t>(i)])
        cell.resize(static_cast<size_t>(widths[static_cast<size_t>(i)]));
      line << cell;
    }
    std::wstring ls = line.str();
    if (static_cast<int>(ls.size()) > linesize) ls.resize(static_cast<size_t>(linesize));
    PrintUtf8WideLine(ls);

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
      AdoDb::_RecordsetPtr rs = impl_->conn->Execute(_bstr_t(sql.c_str()), &affected, adCmdText);
      PrintRecordset(rs, settings);
      return true;
    }

    EnsureTxnForDml(sql, err);
    if (!err.empty()) return false;

    _variant_t ra;
    impl_->conn->Execute(_bstr_t(sql.c_str()), &ra, adCmdText);
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
    AdoDb::_RecordsetPtr rs = impl_->conn->OpenSchema(adSchemaTables);
    int count = 0;
    PrintUtf8("TABLE_NAME\tTABLE_TYPE\n");
    while (!rs->EndOfFile) {
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
      PrintUtf8WideLine(n + L"\t" + t);
      count++;
      if (count >= 255) {
        PrintUtf8("Error: metadata enumeration limit (255) exceeded.\n");
        return false;
      }
      rs->MoveNext();
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
    AdoDb::_RecordsetPtr rs = impl_->conn->OpenSchema(adSchemaViews);
    int count = 0;
    PrintUtf8("VIEW_NAME\n");
    while (!rs->EndOfFile) {
      _variant_t name = rs->Fields->GetItem(L"TABLE_NAME")->GetValue();
      std::wstring n = MetaStr(name);
      if (n.empty()) {
        rs->MoveNext();
        continue;
      }
      PrintUtf8WideLine(n);
      count++;
      if (count >= 255) {
        PrintUtf8("Error: metadata enumeration limit (255) exceeded.\n");
        return false;
      }
      rs->MoveNext();
    }
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
  PrintUtf8("NAME\tTYPE\tDEFINED_SIZE\n");
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
    std::wostringstream os;
    os << cn << L'\t' << dtype << L'\t' << dsize;
    PrintUtf8WideLine(os.str());
    cnt++;
    if (cnt >= 255) {
      PrintUtf8("Error: metadata enumeration limit (255) exceeded.\n");
      return false;
    }
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
      try {
        AdoX::_ViewPtr vw = cat->Views->GetItem(_variant_t(name.c_str()));
        return DumpAdoxColumns(vw->GetColumns(), err);
      } catch (_com_error& e2) {
        err = FormatComError(e2);
        return false;
      }
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
    AdoX::_CatalogPtr cat(__uuidof(AdoX::Catalog));
    cat->PutActiveConnection(_variant_t(static_cast<IDispatch*>(impl_->conn)));
    AdoX::_ViewPtr v = cat->Views->GetItem(_variant_t(name.c_str()));
    AdoX::_CommandPtr cmd = v->GetCommand();
    std::wstring sql = static_cast<const wchar_t*>(cmd->GetCommandText());
    PrintUtf8WideLine(sql);
    return true;
  } catch (_com_error& e) {
    err = FormatComError(e);
    return false;
  }
}

}  // namespace adosql
