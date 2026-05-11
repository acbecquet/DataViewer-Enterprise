#pragma once
#include "ReportLayout.h"
#include <QString>
#include <QRectF>

namespace DVE {

// Abstract command. apply()/undo() mutate the supplied ReportLayout
// in-place. description() is for tooltip / future undo history UI.
class LayoutCommand {
public:
    virtual ~LayoutCommand() = default;
    virtual void apply(ReportLayout&) = 0;
    virtual void undo(ReportLayout&)  = 0;
    virtual QString description() const = 0;
};

// Geometry change (move OR resize — same shape: old/new QRectF).
// Targets one of: cover title/subtitle, divider title, content/cumulative
// title/table/radar/propertiesBox.rect.
class RectCommand : public LayoutCommand {
public:
    RectCommand(const QString& slideKey,
                const QString& elementId,
                const QRectF&  oldRect,
                const QRectF&  newRect);
    void apply(ReportLayout&) override;
    void undo(ReportLayout&)  override;
    QString description() const override;
private:
    void setRect(ReportLayout&, const QRectF&) const;
    QString m_slideKey;
    QString m_elementId;
    QRectF  m_old;
    QRectF  m_new;
};

// Font-size change. Targets cover/divider title, cover subtitle, content/
// cumulative title/table/propertiesBox font.
class FontSizeCommand : public LayoutCommand {
public:
    FontSizeCommand(const QString& slideKey,
                    const QString& elementId,
                    int oldPt,
                    int newPt);
    void apply(ReportLayout&) override;
    void undo(ReportLayout&)  override;
    QString description() const override;
private:
    void setFontPt(ReportLayout&, int) const;
    QString m_slideKey;
    QString m_elementId;
    int     m_old;
    int     m_new;
};

// Table sort change. Mutates ReportLayout::tableSort.
// Note: TableSort has no "enabled" flag in this codebase — an empty
// column string means "insertion order" (i.e. sort disabled).
class SortCommand : public LayoutCommand {
public:
    SortCommand(const QString& oldColumn, Qt::SortOrder oldOrder,
                const QString& newColumn, Qt::SortOrder newOrder);
    void apply(ReportLayout&) override;
    void undo(ReportLayout&)  override;
    QString description() const override { return QStringLiteral("Sort column"); }
private:
    QString       m_oldCol, m_newCol;
    Qt::SortOrder m_oldOrd, m_newOrd;
};

} // namespace DVE
