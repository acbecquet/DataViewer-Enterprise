# DataViewer Enterprise Changelog

## 2026-04-01 - Detailed Sensory Mode UI Polish & Fixes

### UI Layout
- **Unified question form**: Replaced two separate column containers with a single grid layout, eliminating the visual separator between left and right columns. The form now appears as one cohesive area.
- **Question numbering**: Added numbers 1-14 to all questions so the left-to-right, top-to-bottom reading order is immediately clear.
- **4-quadrant alignment**: The question form column split now aligns with the dual radar chart split below, creating a clean 4-quadrant visual layout.
- **Input width capping**: Combo boxes capped at 280px, line edits at 220px, spin boxes at 70px. Inputs no longer stretch to fill all available horizontal space.
- **Combined header row**: Merged the header fields (Test Title, Assessor, Tester, Media, Date) and sample navigation (prev/next, Add Sample, Remove) into a single row. Header input fields narrowed to ~90px to fit everything.
- **Margin fixes**: Added 8px left/right padding inside the question grid for label readability. Equalized top/bottom margins on the header row.
- **Comments border**: Added a solid outline (QFrame::Box) to the Comments text edit for visual clarity.

### Data & Charts
- **Inverted radar charts**: Radar plot normalization is now inverted so that a score of 1 (best) maps to the outermost ring (9) and the worst score maps to the innermost ring (1). Good results now fill the plot area instead of appearing mostly empty.
- **Mouthpiece/Draw Resistance**: Changed from a free-text QLineEdit to a 5-option dropdown combo box with descriptive answers:
  1. Very easy pull. Good design overall
  2. Draw resistance fine, mouthpiece needs improvement
  3. Mouthpiece fine, draw resistance too small
  4. Mouthpiece fine, draw resistance too high
  5. Mouthpiece and draw resistance made it very hard to puff

### Files Modified
- `src/ui/DetailedSensoryPanel.cpp` - Form layout, header/nav merge, mouthpiece combo
- `src/ui/DetailedSensoryPanel.h` - Changed `m_mouthpieceEdit` (QLineEdit) to `m_mouthpieceCombo` (QComboBox)
- `src/pipeline/DetailedSensoryData.h` - Added `kMouthpieceOptions`, inverted `normalizeToRadar()`
