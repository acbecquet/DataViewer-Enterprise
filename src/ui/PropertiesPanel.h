#pragma once
#include <QWidget>
#include <QRectF>

class QDoubleSpinBox;
class QLabel;
class QPushButton;

namespace DVE {

class PropertiesPanel : public QWidget {
    Q_OBJECT
public:
    PropertiesPanel(QWidget* parent = nullptr);

    // Populate controls with the selected canvas item's id and rect (in inches).
    void setSelectedItem(const QString& elementId, const QRectF& rectInches);

    // Empty controls and disable inputs.
    void clearSelection();

signals:
    // Emitted when any of the four spinboxes commits a value (editingFinished).
    void rectEdited(const QString& elementId, const QRectF& newRectInches);

    // Emitted when the user clicks Bring Forward / Send Backward.
    void bringForwardClicked(const QString& elementId);
    void sendBackwardClicked(const QString& elementId);

private:
    void setControlsEnabled(bool on);
    void emitRectIfReady();

    QString m_currentId;                                // empty when no selection
    QDoubleSpinBox *m_x = nullptr, *m_y = nullptr;
    QDoubleSpinBox *m_w = nullptr, *m_h = nullptr;
    QLabel*         m_label = nullptr;
    QPushButton*    m_forward = nullptr;
    QPushButton*    m_backward = nullptr;
};

} // namespace DVE
