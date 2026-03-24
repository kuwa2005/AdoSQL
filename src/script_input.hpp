#pragma once

#include <fstream>
#include <string>
#include <vector>

namespace adosql {

class ScriptStack {
 public:
  bool Push(const std::wstring& path, std::wstring& err);
  void Pop();
  bool ReadLine(std::wstring& out, bool& eof, bool& overflowed, std::wstring& err);
  bool InScript() const { return !levels_.empty(); }

 private:
  struct Level {
    std::wstring path;
    std::ifstream in;
    bool strip_bom = true;
  };
  std::vector<Level> levels_;
};

}  // namespace adosql
