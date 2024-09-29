#include <zip.h>
#include <iostream>
#include <string>
#include <vector>
#include <fstream>
#include <errno.h>
#include <sys/stat.h>

bool create_directory(const std::string& dir) {
    if (mkdir(dir.c_str(), 0777) == -1) {
        if (errno != EEXIST) {
            std::cerr << "Failed to create directory: " << dir << std::endl;
            return false;
        }
    }
    return true;
}

//uses libzip to extract contents of the docx file 
bool unzip_docx(const std::string& docx_path, const std::string& output_dir) {
    int err; 
    zip *zip_archive = zip_open(docx_path.c_str(), 0, &err); 
    if (!zip_archive) {
        std::cerr << "Failed to open DOCX file: " << docx_path << std::endl;
        return false;
    }

    // Create output directory if it doesn't exist
    create_directory(output_dir);

    //extract all files from the docx archive, and write to output_dir
    for (zip_int64_t i = 0; i < zip_get_num_entries(zip_archive, 0); ++i) {
        const char* file_name = zip_get_name(zip_archive, i, 0); 
        std::string output_file_path = output_dir + "/" + file_name;

        zip_file* zf = zip_fopen_index(zip_archive, i, 0);
        if (!zf) {
            std::cerr << "Failed to open file in ZIP: " << file_name << std::endl;
            continue;
        }

        std::ofstream out_file(output_file_path, std::ios::binary);
        char buffer[1024];
        zip_int64_t bytes_read;
        while ((bytes_read = zip_fread(zf, buffer, sizeof(buffer))) > 0) {
            out_file.write(buffer, bytes_read);
        }

        zip_fclose(zf);
        out_file.close();
    }

    zip_close(zip_archive);
    return true;        
}