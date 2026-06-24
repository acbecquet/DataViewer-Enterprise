#pragma once
#include "../pipeline/ReportData.h"
#include "../pipeline/NotesStory.h"
#include <QWidget>
#include <QSet>
#include <QHash>
class QVBoxLayout; class QScrollArea;

namespace DVE {

class NotesStoryPanel : public QWidget {
    Q_OBJECT
public:
    explicit NotesStoryPanel(QWidget* parent = nullptr);

    void setSample(const SampleResult& sample, const QSet<int>& excluded, bool perRowRegime);
    void clear();

public slots:
    void highlightRow(int dataRowIndex);

signals:
    void cellEdited(int dataRowIndex, int col, const QString& text);
    void noteActivated(int dataRowIndex);

private:
    QWidget*     buildNoteCard(const SampleResult& s, int rowIndex, bool excluded, bool perRowRegime);
    QWidget*     buildSummaryBar(const SampleResult& s, const StorySummary& sum, int segIndex, bool perRowRegime);
    void         rebuild();

    SampleResult m_sample;
    QSet<int>    m_excluded;
    bool         m_perRowRegime = false;
    QScrollArea* m_scroll  = nullptr;
    QWidget*     m_content = nullptr;
    QVBoxLayout* m_vbox    = nullptr;
    QHash<int, QWidget*> m_cardByRow;
};

} // namespace DVE
