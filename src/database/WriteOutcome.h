#pragma once

#include "DatabaseManager.h"   // for DVE::WriteResult

namespace DVE {

// How MainWindow should react to the result of an automatic, load-time
// saveFile()/tryWriteFile(). Keeps the policy (clear vs. keep the dirty flag,
// and which message to surface) in one pure, unit-tested place. DATAVIEWER-3.
enum class LoadSavePolicy {
    Saved,         // WriteResult::Success         -> clear dirty, no message
    RetryOffline,  // WriteResult::OfflineReadOnly -> keep dirty, info "offline"
    RetryConflict, // Version/RowDeleted           -> keep dirty, warn "changed by another user"
    RetryError     // UniqueViolation/OtherError   -> keep dirty, warn generic failure
};

LoadSavePolicy classifyLoadSaveResult(WriteResult r);

} // namespace DVE
