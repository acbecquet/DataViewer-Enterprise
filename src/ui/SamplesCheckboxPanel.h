#pragma once
#include <QWidget>
#include <QHash>
#include <QSet>
#include <QVector>
#include "reporting/IReportSource.h"

class QCheckBox;

namespace DVE {

class SamplesCheckboxPanel : public QWidget {
    Q_OBJECT
public:
    SamplesCheckboxPanel(const QVector<SampleRef>& refs, QWidget* parent = nullptr);

    QSet<QString> excludedSampleIds() const;

signals:
    void sampleToggled(const QString& sampleId, bool included);

private:
    QHash<QString, QCheckBox*> m_boxes;
};

} // namespace DVE
