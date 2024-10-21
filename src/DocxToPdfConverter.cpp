#include "DocxToPdfConverter.h"
#include <hpdf.h>
#include <tinyxml2.h>
#include <iostream>
#include <fstream>
#include <sstream>
#include <iomanip>

using namespace tinyxml2;

// Struct Definitions
struct TextFragment {
    std::string text;
    HPDF_Font font;
    int fontSize;
    float r, g, b; // Color components
};

struct TableCell {
    std::vector<TextFragment> textFragments;
};

struct Table {
    std::vector<std::vector<TableCell>> rows; // Each row contains multiple cells
};

// Function to Render Text with Wrapping
void renderTextWithWrapping(HPDF_Doc pdf, HPDF_Page &page, const std::string &text,
                            float &cursorX, float &cursorY, float pageWidth, float fontSize,
                            HPDF_Font font, float leftMargin, float rightMargin)
{
    size_t pos = 0;
    size_t len = text.length();

    while (pos < len)
    {
        // Find the next whitespace or non-whitespace run
        size_t nextPos = pos;
        bool isSpace = isspace(text[pos]);

        while (nextPos < len && isspace(text[nextPos]) == isSpace)
            nextPos++;

        // Extract the token (word or spaces)
        std::string token = text.substr(pos, nextPos - pos);
        float tokenWidth = HPDF_Page_TextWidth(page, token.c_str());

        // If word doesn't fit on the current line
        if (cursorX + tokenWidth > pageWidth - rightMargin && !isSpace)
        {
            cursorY -= fontSize + 2.0f;
            cursorX = leftMargin;

            // Handle page break
            if (cursorY < 50)
            {
                page = HPDF_AddPage(pdf);
                HPDF_Page_SetSize(page, HPDF_PAGE_SIZE_A4, HPDF_PAGE_PORTRAIT);
                cursorY = HPDF_Page_GetHeight(page) - 50;
            }
        }

        // Render the token
        HPDF_Page_BeginText(page);
        HPDF_Page_SetFontAndSize(page, font, fontSize);
        HPDF_Page_MoveTextPos(page, cursorX, cursorY);
        HPDF_Page_ShowText(page, token.c_str());
        HPDF_Page_EndText(page);

        cursorX += tokenWidth;

        pos = nextPos;
    }
}

// Function to Parse a Table Element and Populate the Table Structure
Table parseTable(XMLElement *tblElement, HPDF_Font defaultFont, HPDF_Font boldFont,
                 HPDF_Font italicFont, HPDF_Font boldItalicFont)
{
    Table table;

    for (XMLElement *tr = tblElement->FirstChildElement("w:tr"); tr; tr = tr->NextSiblingElement("w:tr"))
    {
        std::vector<TableCell> row;

        for (XMLElement *tc = tr->FirstChildElement("w:tc"); tc; tc = tc->NextSiblingElement("w:tc"))
        {
            TableCell cell;

            // Iterate over paragraphs within the cell
            for (XMLElement *para = tc->FirstChildElement("w:p"); para; para = para->NextSiblingElement("w:p"))
            {
                for (XMLElement *run = para->FirstChildElement("w:r"); run; run = run->NextSiblingElement("w:r"))
                {
                    XMLElement *rPr = run->FirstChildElement("w:rPr");
                    bool isBold = false;
                    bool isItalic = false;
                    std::string color = "000000";
                    int fontSize = 12;
                    HPDF_Font font = defaultFont;

                    if (rPr)
                    {
                        if (rPr->FirstChildElement("w:b"))
                        {
                            isBold = true;
                        }
                        if (rPr->FirstChildElement("w:i"))
                        {
                            isItalic = true;
                        }
                        XMLElement *colorElement = rPr->FirstChildElement("w:color");
                        if (colorElement && colorElement->Attribute("w:val"))
                        {
                            color = colorElement->Attribute("w:val");
                        }
                        XMLElement *szElement = rPr->FirstChildElement("w:sz");
                        if (szElement && szElement->Attribute("w:val"))
                        {
                            fontSize = std::stoi(szElement->Attribute("w:val")) / 2;
                        }
                    }

                    if (isBold && isItalic)
                    {
                        font = boldItalicFont;
                    }
                    else if (isBold)
                    {
                        font = boldFont;
                    }
                    else if (isItalic)
                    {
                        font = italicFont;
                    }

                    // Convert color to RGB
                    float r = 0, g = 0, b = 0;
                    if (color.length() == 6)
                    {
                        std::stringstream ss;
                        ss << std::hex << color;
                        unsigned int rgb;
                        ss >> rgb;
                        r = ((rgb >> 16) & 0xFF) / 255.0f;
                        g = ((rgb >> 8) & 0xFF) / 255.0f;
                        b = (rgb & 0xFF) / 255.0f;
                    }

                    // Iterate over child elements within the run
                    for (XMLElement *child = run->FirstChildElement(); child; child = child->NextSiblingElement())
                    {
                        if (strcmp(child->Name(), "w:t") == 0)
                        {
                            // Text element
                            if (child->GetText())
                            {
                                const char* spacePreserve = child->Attribute("xml:space");
                                std::string textContent = child->GetText();

                                std::string text;
                                if (spacePreserve && strcmp(spacePreserve, "preserve") == 0)
                                {
                                    text = textContent;
                                }
                                else
                                {
                                    text = textContent;
                                }

                                // Create a TextFragment and add to the cell's textFragments
                                TextFragment fragment;
                                fragment.text = text;
                                fragment.font = font;
                                fragment.fontSize = fontSize;
                                fragment.r = r;
                                fragment.g = g;
                                fragment.b = b;
                                cell.textFragments.push_back(fragment);
                            }
                        }
                        else if (strcmp(child->Name(), "w:br") == 0)
                        {
                            // Line break within table cell
                            // You can handle line breaks within table cells if needed
                            continue;
                        }
                    }
                }
            }

            row.push_back(cell);
        }
        // Push back the row after all cells in the row have been processed
        table.rows.push_back(row);
    }

    return table;
}

// Function to Render a Table Using libharu
void renderTable(HPDF_Doc pdf, HPDF_Page &page, const Table &table,
                 float &cursorX, float &cursorY, float pageWidth,
                 float leftMargin, float rightMargin)
{
    if (table.rows.empty()) return;

    // Table properties
    float tableStartX = leftMargin;
    float tableWidth = pageWidth - leftMargin - rightMargin;
    size_t numCols = table.rows[0].size();

    // Define column widths (evenly distributed)
    std::vector<float> colWidths(numCols, tableWidth / numCols);

    // Define row height (fixed for simplicity)
    float rowHeight = 20.0f;

    // Draw table borders and content
    for (const auto &row : table.rows)
    {
        // Draw horizontal line for the row
        HPDF_Page_SetRGBStroke(page, 0, 0, 0); // Black color for borders
        HPDF_Page_SetLineWidth(page, 0.5);
        HPDF_Page_MoveTo(page, tableStartX, cursorY);
        HPDF_Page_LineTo(page, tableStartX + tableWidth, cursorY);
        HPDF_Page_Stroke(page);

        float cellX = tableStartX;

        for (size_t i = 0; i < row.size(); ++i)
        {
            const TableCell &cell = row[i];

            // Draw vertical line for the cell
            HPDF_Page_MoveTo(page, cellX, cursorY);
            HPDF_Page_LineTo(page, cellX, cursorY - rowHeight);
            HPDF_Page_Stroke(page);

            // Render cell content
            float textCursorX = cellX + 5; // Padding from left
            float textCursorY = cursorY - 15; // Adjust as needed

            // Render text fragments within the cell
            for (const auto &fragment : cell.textFragments)
            {
                // Set font and color
                HPDF_Page_SetFontAndSize(page, fragment.font, fragment.fontSize);
                HPDF_Page_SetRGBFill(page, fragment.r, fragment.g, fragment.b);

                renderTextWithWrapping(pdf, page, fragment.text, textCursorX, textCursorY,
                                       pageWidth, fragment.fontSize, fragment.font, leftMargin, rightMargin);
            }

            cellX += colWidths[i];
        }

        // Draw the last vertical line of the row
        HPDF_Page_MoveTo(page, tableStartX + tableWidth, cursorY);
        HPDF_Page_LineTo(page, tableStartX + tableWidth, cursorY - rowHeight);
        HPDF_Page_Stroke(page);

        // Move cursorY for the next row
        cursorY -= rowHeight;

        // Handle page break
        if (cursorY < 50)
        {
            page = HPDF_AddPage(pdf);
            HPDF_Page_SetSize(page, HPDF_PAGE_SIZE_A4, HPDF_PAGE_PORTRAIT);
            cursorY = HPDF_Page_GetHeight(page) - 50;
        }
    }

    // After table, adjust cursorY
    cursorY -= 10.0f; // Space after table
}

// Function to Process Elements (Paragraphs and Tables)
void processElement(XMLElement *element, HPDF_Doc pdf, HPDF_Page &page,
                    float &cursorX, float &cursorY, float pageWidth,
                    HPDF_Font defaultFont, HPDF_Font boldFont, HPDF_Font italicFont,
                    HPDF_Font boldItalicFont, float leftMargin, float rightMargin)
{
    std::string elemName = element->Name();

    if (elemName == "w:p")
    {
        // Handle paragraph
        int fontSize = 12;

        // For each run in the paragraph
        for (XMLElement *run = element->FirstChildElement("w:r"); run; run = run->NextSiblingElement("w:r"))
        {
            XMLElement *rPr = run->FirstChildElement("w:rPr");
            bool isBold = false;
            bool isItalic = false;
            std::string color = "000000";

            if (rPr)
            {
                if (rPr->FirstChildElement("w:b"))
                {
                    isBold = true;
                }
                if (rPr->FirstChildElement("w:i"))
                {
                    isItalic = true;
                }
                XMLElement *colorElement = rPr->FirstChildElement("w:color");
                if (colorElement && colorElement->Attribute("w:val"))
                {
                    color = colorElement->Attribute("w:val");
                }
                XMLElement *szElement = rPr->FirstChildElement("w:sz");
                if (szElement && szElement->Attribute("w:val"))
                {
                    fontSize = std::stoi(szElement->Attribute("w:val")) / 2;
                }
            }

            // Iterate over child elements within the run
            for (XMLElement *child = run->FirstChildElement(); child; child = child->NextSiblingElement())
            {
                std::string text;

                if (strcmp(child->Name(), "w:t") == 0)
                {
                    // Text element
                    if (child->GetText())
                    {
                        text = child->GetText();
                    }
                }
                else if (strcmp(child->Name(), "w:tab") == 0)
                {
                    // Tab element
                    const float tabWidth = 40.0f;
                    cursorX += tabWidth;
                    continue;
                }
                else if (strcmp(child->Name(), "w:br") == 0)
                {
                    // Line break element
                    cursorY -= fontSize + 2.0f;
                    cursorX = leftMargin;

                    // Check if we need a new page after adjusting cursorY
                    if (cursorY < 50)
                    {
                        page = HPDF_AddPage(pdf);
                        HPDF_Page_SetSize(page, HPDF_PAGE_SIZE_A4, HPDF_PAGE_PORTRAIT);
                        cursorY = HPDF_Page_GetHeight(page) - 50;
                    }

                    continue;
                }
                else
                {
                    continue;
                }

                // Set font and size
                HPDF_Font font = defaultFont;
                if (isBold && isItalic)
                {
                    font = boldItalicFont;
                }
                else if (isBold)
                {
                    font = boldFont;
                }
                else if (isItalic)
                {
                    font = italicFont;
                }
                HPDF_Page_SetFontAndSize(page, font, fontSize);

                // Convert color to RGB
                unsigned int r = 0, g = 0, b = 0;
                if (color.length() == 6)
                {
                    std::stringstream ss;
                    ss << std::hex << color;
                    unsigned int rgb;
                    ss >> rgb;
                    r = (rgb >> 16) & 0xFF;
                    g = (rgb >> 8) & 0xFF;
                    b = rgb & 0xFF;
                }
                HPDF_Page_SetRGBFill(page, r / 255.0f, g / 255.0f, b / 255.0f);

                if (!text.empty())
                {
                    renderTextWithWrapping(pdf, page, text, cursorX, cursorY,
                                           pageWidth, fontSize, font, leftMargin, rightMargin);
                }
            }
        }

        float lineHeight = fontSize + 2.0f;

        // Move to next line after paragraph
        cursorY -= lineHeight;
        cursorX = leftMargin;

        // Handle page break after the paragraph
        if (cursorY < 50)
        {
            page = HPDF_AddPage(pdf);
            HPDF_Page_SetSize(page, HPDF_PAGE_SIZE_A4, HPDF_PAGE_PORTRAIT);
            cursorY = HPDF_Page_GetHeight(page) - 50;
        }
    }
    else if (elemName == "w:tbl")
    {
        // Handle table
        Table table = parseTable(element, defaultFont, boldFont, italicFont, boldItalicFont);
        renderTable(pdf, page, table, cursorX, cursorY, pageWidth, leftMargin, rightMargin);
    }
    else
    {
        // Handle other elements if necessary
        
    }
}

// Generates PDF from the parsed DOCX content
void generatePDF(const std::string &docxDir, const std::string &outputPdfPath)
{
    HPDF_Doc pdf = HPDF_New(NULL, NULL);

    if (!pdf)
    {
        std::cerr << "Failed to create PDF object." << std::endl;
        return;
    }

    HPDF_UseUTFEncodings(pdf);
    HPDF_SetCurrentEncoder(pdf, "UTF-8");

    // Load fonts
    const char *fontPath = "../fonts/dejavu-fonts-ttf/ttf/";
    const char *fontName = HPDF_LoadTTFontFromFile(pdf, (std::string(fontPath) + "DejaVuSans.ttf").c_str(), HPDF_TRUE);
    if (!fontName)
    {
        std::cerr << "Failed to load TrueType font DejaVuSans.ttf." << std::endl;
        HPDF_Free(pdf);
        return;
    }
    const char *boldFontName = HPDF_LoadTTFontFromFile(pdf, (std::string(fontPath) + "DejaVuSans-Bold.ttf").c_str(), HPDF_TRUE);
    const char *italicFontName = HPDF_LoadTTFontFromFile(pdf, (std::string(fontPath) + "DejaVuSans-Oblique.ttf").c_str(), HPDF_TRUE);
    const char *boldItalicFontName = HPDF_LoadTTFontFromFile(pdf, (std::string(fontPath) + "DejaVuSansCondensed-BoldOblique.ttf").c_str(), HPDF_TRUE);

    // Get HPDF_Font objects
    HPDF_Font defaultFont = HPDF_GetFont(pdf, fontName, "UTF-8");
    HPDF_Font boldFont = HPDF_GetFont(pdf, boldFontName, "UTF-8");
    HPDF_Font italicFont = HPDF_GetFont(pdf, italicFontName, "UTF-8");
    HPDF_Font boldItalicFont = HPDF_GetFont(pdf, boldItalicFontName, "UTF-8");

    if (!defaultFont || !boldFont || !italicFont || !boldItalicFont)
    {
        std::cerr << "Failed to get one or more HPDF_Font objects." << std::endl;
        HPDF_Free(pdf);
        return;
    }

    // Create a new page and set its size
    HPDF_Page page = HPDF_AddPage(pdf);
    HPDF_Page_SetSize(page, HPDF_PAGE_SIZE_A4, HPDF_PAGE_PORTRAIT);

    // Setting page info and margins
    float leftMargin = 50.0f;
    float rightMargin = 50.0f;
    float cursorX = leftMargin;
    float cursorY = HPDF_Page_GetHeight(page) - 50;
    float pageWidth = HPDF_Page_GetWidth(page);

    // Try to load and parse the document.xml file
    std::string documentXmlPath = docxDir + "/word/document.xml";
    XMLDocument doc;
    if (doc.LoadFile(documentXmlPath.c_str()) != XML_SUCCESS)
    {
        std::cerr << "Failed to load " << documentXmlPath << std::endl;
        HPDF_Free(pdf);
        return;
    }

    XMLElement *root = doc.RootElement(); // <w:document>
    if (!root)
    {
        std::cerr << "No root element in document.xml." << std::endl;
        HPDF_Free(pdf);
        return;
    }

    XMLElement *body = root->FirstChildElement("w:body"); // Content in the XML file
    if (!body)
    {
        std::cerr << "No body element in document.xml." << std::endl;
        HPDF_Free(pdf);
        return;
    }

    // Iterate through all child elements of <w:body> in order
    for (XMLElement *element = body->FirstChildElement(); element; element = element->NextSiblingElement())
    {
        processElement(element, pdf, page, cursorX, cursorY, pageWidth,
                       defaultFont, boldFont, italicFont, boldItalicFont, leftMargin, rightMargin);
    }

    if (HPDF_SaveToFile(pdf, outputPdfPath.c_str()) != HPDF_OK)
    {
        std::cerr << "Failed to save PDF to " << outputPdfPath << std::endl;
    }
    else
    {
        std::cout << "PDF saved successfully to " << outputPdfPath << std::endl;
    }

    HPDF_Free(pdf);
}
