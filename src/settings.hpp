#pragma once

#include <cstddef>

namespace adosql {

struct DisplaySettings {
  int pagesize = 24;    // 5..500
  int linesize = 79;    // 20..500
  int maxrows = 0;      // 0 = unlimited
  bool widthclamp = false;
};

}  // namespace adosql
