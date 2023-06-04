#include <main_window_pl/main_window.h>
#include <main_window_pl/ui_main_window.h>

class ConverterMainWindow::PrivateData {

public:
	Ui::ConverterWindow ui;
};

ConverterMainWindow::ConverterMainWindow(QMainWindow* parent) : QMainWindow(parent), m_pData(std::make_unique<PrivateData>()) {
	m_pData->ui.setupUi(this);
}

ConverterMainWindow::~ConverterMainWindow() {}
