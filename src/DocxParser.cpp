#include <zip.h>
#include <iostream>
#include <string>
#include <vector>
#include <fstream>

//uses libzip to extract contents of the docx file 
bool unzip_docx(const std::string& docx_path, const std::string& output_dir) {
    int err; 
    zip *zip_file = zip_open(docx_path.c_str(), 0, &err); 
    if (!zip_file) {
        std::cerr << "Failed to open DOCX file: " << docx_path << std::endl;
        return false;
    }

    //extract all files from the docx archive 
}