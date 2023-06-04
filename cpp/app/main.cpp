#include <QApplication>
#include <main_window_pl/main_window.h>

int main(int argc, char** argv) {
	QApplication app(argc, argv);
	ConverterMainWindow MarineMainWind;
	MarineMainWind.show();
	return app.exec();
	return 0;
}