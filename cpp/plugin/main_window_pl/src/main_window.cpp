#include <main_window_pl/main_window.h>
#include <main_window_pl/ui_main_window.h>
#include <handler/handler.h>

class ConverterMainWindow::PrivateData {
public:
	Ui::ConverterWindow ui;
	Handler handler_;
};

ConverterMainWindow::ConverterMainWindow(QMainWindow* parent) : QMainWindow(parent), pData_(std::make_unique<PrivateData>()) {
	pData_->ui.setupUi(this);
	setConnections();
}

ConverterMainWindow::~ConverterMainWindow() {}

void ConverterMainWindow::setConnections() {
	connect(pData_->ui.pbLoadOldWord, &QPushButton::clicked, this, &ConverterMainWindow::onLoadOldWord);
	connect(pData_->ui.pbLoadNewWord, &QPushButton::clicked, this, &ConverterMainWindow::onLoadNewWord);
	connect(pData_->ui.pbCompareAndExport, &QPushButton::clicked, this, &ConverterMainWindow::onCompareAndExport);
	connect(pData_->ui.pbClean, &QPushButton::clicked, this, &ConverterMainWindow::onClean);
}

void ConverterMainWindow::onLoadOldWord() {

}

void ConverterMainWindow::onLoadNewWord() {

}

void ConverterMainWindow::onCompareAndExport() {

}

void ConverterMainWindow::onClean() {

}
