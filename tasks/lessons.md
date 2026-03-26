# Lessons Learned

## 2026-03-19: Sensory layout should be horizontal, not vertical
- **Mistake:** Set sensory panel splitter to `Qt::Vertical` (cards top, chart bottom) matching the TPM layout
- **Correction:** Sensory mode uses a **horizontal** layout — cards on the left, radar chart on the right. This is intentionally different from TPM mode (table top, plot bottom).
- **Rule:** Don't assume new UI modes should mirror existing layout orientations. When the user says "place data and plot in the space where table and chart are," the spatial arrangement can still differ.
