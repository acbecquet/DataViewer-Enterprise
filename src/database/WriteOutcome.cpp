#include "WriteOutcome.h"

namespace DVE {

LoadSavePolicy classifyLoadSaveResult(WriteResult r)
{
    switch (r) {
    case WriteResult::Success:         return LoadSavePolicy::Saved;
    case WriteResult::OfflineReadOnly: return LoadSavePolicy::RetryOffline;
    case WriteResult::VersionMismatch: return LoadSavePolicy::RetryConflict;
    case WriteResult::RowDeleted:      return LoadSavePolicy::RetryConflict;
    case WriteResult::UniqueViolation: return LoadSavePolicy::RetryError;
    case WriteResult::OtherError:      return LoadSavePolicy::RetryError;
    }
    return LoadSavePolicy::RetryError;  // defensive; unreachable for a valid enum
}

} // namespace DVE
