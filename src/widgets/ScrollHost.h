#pragma once

#include <QScrollArea>

namespace DVE {

// Universal scroll wrapper used across MainWindow's central pages, side docks,
// the ribbon group row, and dialog content. It is the structural guarantee of
// the v2.7.0 responsive overhaul: any region wrapped in a ScrollHost can never
// be clipped without a scrollbar appearing, in either direction, the instant
// content exceeds the viewport.
//
// Behaviour (set once in the ctor):
//   - setWidgetResizable(true)  -> content expands to fill the viewport when
//     there is room; ScrollHost scrolls (not clips) when there is not.
//   - both scrollbar policies = Qt::ScrollBarAsNeeded -> invisible when content
//     fits, present the moment it overflows.
//   - QFrame::NoFrame + transparent background + zero viewport margins -> a
//     visual no-op until it actually scrolls (no border, no background fill,
//     no inset), so wrapping a region never changes its standard-size look.
//
// Use the static factory at call sites for a clean read:
//     m_centralStack->addWidget(ScrollHost::wrap(m_centralSplitter));
// The optional second argument restricts scrolling to one axis (the ribbon
// group row passes Qt::Horizontal so the single fixed-height row never grows a
// vertical scrollbar).
//
// Caveat -- QSplitter content: a QSplitter inside a widgetResizable scroll area
// reports its sizeHint as the sum of its children's hints, so the splitter is
// driven by the ScrollHost rather than the reverse. This is fine: at small
// sizes the ScrollHost scrolls the whole splitter; at normal sizes the splitter
// fills the viewport and its handles work as usual. Give splitter children a
// sensible minimumHeight/Width (already done for the TPM plot) so the overflow
// point is meaningful.
class ScrollHost : public QScrollArea {
    Q_OBJECT
public:
    explicit ScrollHost(QWidget* parent = nullptr);

    // Convenience factory: construct a ScrollHost, take ownership of `content`
    // via setWidget(), and return the host. `scroll` selects which axes scroll
    // as-needed; an axis not in `scroll` is set to Qt::ScrollBarAlwaysOff. The
    // host's parent is left null so the caller can re-parent it by adding it to
    // a layout / stacked widget / dock, exactly like a plain `new` widget.
    static ScrollHost* wrap(QWidget* content,
                            Qt::Orientations scroll = Qt::Horizontal | Qt::Vertical);

    // True when that direction's scrollbar is currently shown. Used by the
    // --ui-stress harness to assert the fits-or-scrolls guarantee per region.
    bool scrollbarActive(Qt::Orientation o) const;

    // True when either direction has a non-zero scrollable range.
    bool contentOverflows() const;
};

} // namespace DVE
