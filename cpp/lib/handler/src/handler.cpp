#include <unordered_map>
#include <vector>
#include <string>

#include <handler/handler.h>

class Handler::PrivateData {
public:
    std::string pathToOldWord_;
    std::string pathToNewWord_;

    std::unordered_map<std::string, std::vector<std::string>> dataOldDoc_;
    std::unordered_map<std::string, std::vector<std::string>> dataNewDoc_;

    std::unordered_map<std::string, std::string> dataOldHandled_;
    std::unordered_map<std::string, std::string> dataNewHandled_;

    std::unordered_map<std::string, std::vector<std::string>> mistakedWithDiffRevInNew_;
    std::unordered_map<std::string, std::vector<std::string>> mistakedWithDiffRevInOld_;

    std::unordered_map<std::string, std::vector<std::string>> names_ = {
        { "Not changed", std::vector<std::string>() },
        { "Mistaked", std::vector<std::string>() },
        { "Changed", std::vector<std::string>() },
        { "New", std::vector<std::string>() },
        { "Deleted", std::vector<std::string>() }
    };

    std::string pathToPDBExcel_;
    std::unordered_map<std::string, std::string> pdbExcelData_;

    std::vector<std::string> inPDB_;
    std::vector<std::string> notInPDB_;
};

Handler::Handler() : pData_(std::make_unique<PrivateData>()) {}

Handler::~Handler() {}


void Handler::setPathToOldWord(const std::string& path) {
    pData_->pathToOldWord_ = path;
}

void Handler::setPathToNewWord(const std::string& path) {
    pData_->pathToNewWord_ = path;
}

void Handler::parseOldWord() {

}

void Handler::exportOldWordData() {

}

void Handler::exportNewWordData() {

}
