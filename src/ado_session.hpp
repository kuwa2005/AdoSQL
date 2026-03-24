#pragma once

#include "settings.hpp"

#include <memory>
#include <string>

namespace adosql {

class AdoSession {
 public:
  AdoSession();
  ~AdoSession();
  AdoSession(const AdoSession&) = delete;
  AdoSession& operator=(const AdoSession&) = delete;

  bool Connect(const std::wstring& db_path, const std::wstring* password, std::wstring& error);
  bool CreateDatabase(const std::wstring& db_path, std::wstring& error);
  void Disconnect();

  void Commit(std::wstring& err);
  void Rollback(std::wstring& err);

  bool ExecuteInteractive(const std::wstring& sql, DisplaySettings& settings, std::wstring& err);

  bool RunTableList(std::wstring& err);
  bool RunQueryList(std::wstring& err);
  bool RunDescribe(const std::wstring& name, std::wstring& err);
  bool RunShowSql(const std::wstring& name, std::wstring& err);

  void OnExitSession();

 private:
  struct Impl;
  std::unique_ptr<Impl> impl_;

  void EnsureTxnForDml(const std::wstring& sql, std::wstring& err);
  bool InTransaction() const;
};

}  // namespace adosql
