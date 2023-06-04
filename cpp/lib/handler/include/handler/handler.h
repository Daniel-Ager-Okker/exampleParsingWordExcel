#include <memory>

class Handler {
public:
	Handler();
	~Handler();

    void setPathToOldWord(const std::string& path);
    void setPathToNewWord(const std::string& path);
    void parseOldWord();
    void exportOldWordData();
    void exportNewWordData();

private:
	class PrivateData;
	std::unique_ptr<PrivateData> pData_;
};