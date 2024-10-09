#include <iostream>
#include "DocxParser.h"
#include "DocxToPdfConverter.h"

// expands the ~ directory since cpp doesn't do it like shell
std::string expand_home_directory(const std::string &path)
{
    if (!path.empty() && path[0] == '~')
    {
        const char *home = getenv("HOME");
        if (home)
        {
            return std::string(home) + path.substr(1);
        }
    }
    return path;
}

int main()
{
    std::string base_dir; 

    //preprocessor directives so only necessary code gets compiled
    #ifdef __APPLE__
        // macOS: Include Desktop in the path
        base_dir = "~/Desktop/workspace/FileFormatConverter";
    #elif defined(__linux__)
        // Ubuntu/Linux: Do not include Desktop in the path
        base_dir = "~/workspace/FileFormatConverter";
    #else
        #error "Unsupported operating system"
    #endif

    // TODO: Hardcoded for now but should ask for user input, could handle this on the cloud as a future addition
    std::string docx_file = expand_home_directory(base_dir + "/example.docx");
    std::string output_dir = expand_home_directory(base_dir + "/outdir");
    std::string output_pdf = base_dir;

    if (unzip_docx(docx_file, output_dir))
    {
        std::cout << "DOCX file successfully unzipped!" << std::endl;
        generatePDF(output_dir, output_pdf);
        std::cout << "PDF file successfully generated!" << std::endl;
    }
    else
    {
        std::cerr << "Failed to unzip DOCX file!" << std::endl;
    }

    return 0;
}
