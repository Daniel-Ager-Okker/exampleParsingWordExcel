#pragma once

#include <QMainWindow>
#include <memory>

class ConverterMainWindow : public QMainWindow {
	Q_OBJECT
public:
	explicit ConverterMainWindow(QMainWindow* parent = nullptr);

	~ConverterMainWindow();

private:
	void setConnections();

private slots:
	void onLoadOldWord();
	void onLoadNewWord();
	void onCompareAndExport();
	void onClean();


private:
	class PrivateData;
	std::unique_ptr<PrivateData> pData_;
};