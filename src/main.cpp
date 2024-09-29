#include <iostream>
#include "DocxParser.cpp" 

//expands the ~ directory since cpp doesn't do it like shell
std::string expand_home_directory(const std::string& path) {
    if (!path.empty() && path[0] == '~') {
        const char* home = getenv("HOME"); 
        if (home) {
            return std::string(home) + path.substr(1); 
        }
    }
    return path; 
}

int main() {
    //TODO: Hardcoded for now but should ask for user input, could handle this on the cloud as a future addition 
    std::string docx_file = expand_home_directory("~/workspace/FileFormatConverter/example.docx");  
    std::string output_dir = expand_home_directory("~/workspace/FileFormatConverter/outdir");

    if (unzip_docx(docx_file, output_dir)) {
        std::cout << "DOCX file successfully unzipped!" << std::endl;
    } else {
        std::cerr << "Failed to unzip DOCX file!" << std::endl;
    }

    return 0;
}
