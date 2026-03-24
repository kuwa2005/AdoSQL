#pragma once

namespace adosql {

class AdoSession;
struct DisplaySettings;

void RunRepl(AdoSession& session, DisplaySettings& settings);

}  // namespace adosql
