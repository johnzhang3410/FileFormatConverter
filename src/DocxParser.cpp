#include "DocxParser.h"
#include <zip.h>
#include <iostream>
#include <fstream>
#include <sys/stat.h>
#include <errno.h>

// Function to recursively create directories
bool create_directories(const std::string &path)
{
    if (path.empty())
    {
        return true;
    }

    // Check if the path already exists
    struct stat info;
    if (stat(path.c_str(), &info) == 0)
    {
        if (S_ISDIR(info.st_mode))
        {
            return true; // Directory already exists
        }
        else
        {
            std::cerr << "Path exists but is not a directory: " << path << std::endl;
            return false;
        }
    }

    // Recursively create parent directories
    size_t pos = path.find_last_of("/\\");
    if (pos != std::string::npos)
    {
        std::string parent = path.substr(0, pos);
        if (!create_directories(parent))
        {
            return false;
        }
    }

    // Create the directory
    if (mkdir(path.c_str(), 0777) == -1)
    {
        if (errno != EEXIST)
        {
            std::cerr << "Failed to create directory: " << path << " Error: " << strerror(errno) << std::endl;
            return false;
        }
    }
    return true;
}

// Uses libzip to extract contents of the DOCX file
bool unzip_docx(const std::string &docx_path, const std::string &output_dir)
{
    int err;
    zip *zip_archive = zip_open(docx_path.c_str(), ZIP_RDONLY, &err);
    if (!zip_archive)
    {
        std::cerr << "Failed to open DOCX file: " << docx_path << std::endl;
        return false;
    }

    // Create output directory if it doesn't exist
    if (!create_directories(output_dir))
    {
        std::cerr << "Failed to create output directory: " << output_dir << std::endl;
        zip_close(zip_archive);
        return false;
    }

    // Extract all files from the DOCX archive
    zip_int64_t num_entries = zip_get_num_entries(zip_archive, 0);
    for (zip_int64_t i = 0; i < num_entries; ++i)
    {
        const char *file_name = zip_get_name(zip_archive, i, 0);
        if (!file_name)
        {
            std::cerr << "Failed to get name for entry " << i << std::endl;
            continue;
        }

        std::string output_file_path = output_dir + "/" + file_name;

        // Check if the entry is a directory
        zip_stat_t sb;
        if (zip_stat_index(zip_archive, i, 0, &sb) == 0)
        {
            std::string name(file_name);
            if (!name.empty() && name.back() == '/')
            {
                // It's a directory, create it
                if (!create_directories(output_file_path))
                {
                    std::cerr << "Failed to create directory: " << output_file_path << std::endl;
                    continue;
                }
                continue; // Move to the next entry
            }
        }

        // Ensure parent directories exist
        size_t last_slash = output_file_path.find_last_of("/\\");
        if (last_slash != std::string::npos)
        {
            std::string dir = output_file_path.substr(0, last_slash);
            if (!create_directories(dir))
            {
                std::cerr << "Failed to create directory: " << dir << std::endl;
                continue;
            }
        }

        // Open the file inside the ZIP archive
        zip_file *zf = zip_fopen_index(zip_archive, i, 0);
        if (!zf)
        {
            std::cerr << "Failed to open file in ZIP: " << file_name << std::endl;
            continue;
        }

        // Open the output file
        std::ofstream out_file(output_file_path, std::ios::binary);
        if (!out_file.is_open())
        {
            std::cerr << "Failed to open output file: " << output_file_path << std::endl;
            zip_fclose(zf);
            continue;
        }

        // Read from ZIP and write to the output file
        char buffer[4096];
        zip_int64_t bytes_read;
        while ((bytes_read = zip_fread(zf, buffer, sizeof(buffer))) > 0)
        {
            out_file.write(buffer, bytes_read);
        }

        if (bytes_read < 0)
        {
            std::cerr << "Error reading from ZIP file: " << file_name << std::endl;
        }

        zip_fclose(zf);
        out_file.close();
    }

    zip_close(zip_archive);
    return true;
}
