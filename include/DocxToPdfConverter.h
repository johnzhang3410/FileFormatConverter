#ifndef DOCXTOPDFCONVERTER_H
#define DOCXTOPDFCONVERTER_H

#include <string>

// pass in by const reference to save memory space
void generatePDF(const std::string &docxDir, const std::string &outputPdfPath);

#endif
