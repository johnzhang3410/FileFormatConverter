#ifndef DOCXPARSER_H
#define DOCXPARSER_H

#include <string>

bool create_directories(const std::string &dir);
bool unzip_docx(const std::string &docx_path, const std::string &output_dir);

#endif
