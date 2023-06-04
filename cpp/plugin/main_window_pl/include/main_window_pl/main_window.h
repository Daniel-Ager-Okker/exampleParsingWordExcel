#pragma once

#include <QMainWindow>
#include <memory>

class ConverterMainWindow : public QMainWindow {
	Q_OBJECT
public:
	explicit ConverterMainWindow(QMainWindow* parent = nullptr);

	~ConverterMainWindow();

private:
	class PrivateData;
	std::unique_ptr<PrivateData> m_pData;
};