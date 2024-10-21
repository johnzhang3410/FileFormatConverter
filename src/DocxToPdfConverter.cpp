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
    size_t gridSpan = 1; // Default gridSpan is 1 (no colspan)
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
        bool isSpace = isspace(static_cast<unsigned char>(text[pos]));

        while (nextPos < len && isspace(static_cast<unsigned char>(text[nextPos])) == isSpace)
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

            // Check for gridSpan
            XMLElement *tcPr = tc->FirstChildElement("w:tcPr");
            if (tcPr)
            {
                XMLElement *gridSpan = tcPr->FirstChildElement("w:gridSpan");
                if (gridSpan && gridSpan->Attribute("w:val"))
                {
                    cell.gridSpan = std::stoul(gridSpan->Attribute("w:val"));
                }
            }

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
                            // Handle if necessary
                            TextFragment fragment;
                            fragment.text = "\n";
                            fragment.font = font;
                            fragment.fontSize = fontSize;
                            fragment.r = r;
                            fragment.g = g;
                            fragment.b = b;
                            cell.textFragments.push_back(fragment);
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

float calculateTextHeight(HPDF_Doc pdf, HPDF_Page page, const std::string &text,
                          float fontSize, HPDF_Font font, float cellWidth)
{
    size_t pos = 0;
    size_t len = text.length();
    int lines = 1; // Start with at least one line
    float cursorX = 0.0f;

    while (pos < len)
    {
        // Handle line breaks
        if (text[pos] == '\n')
        {
            lines++;
            cursorX = 0.0f;
            pos++;
            continue;
        }

        // Find the next whitespace or non-whitespace run
        size_t nextPos = pos;
        bool isSpace = isspace(static_cast<unsigned char>(text[pos]));

        while (nextPos < len && isspace(static_cast<unsigned char>(text[nextPos])) == isSpace && text[nextPos] != '\n')
            nextPos++;

        // Extract the token
        std::string token = text.substr(pos, nextPos - pos);
        float tokenWidth = HPDF_Page_TextWidth(page, token.c_str());

        // If word doesn't fit on the current line
        if (cursorX + tokenWidth > cellWidth && !isSpace)
        {
            lines++;
            cursorX = 0.0f;
        }

        cursorX += tokenWidth;
        pos = nextPos;
    }

    // Calculate total height based on number of lines
    float lineHeight = fontSize + 2.0f; // Adjust as necessary
    return lines * lineHeight;
}

float calculateCellHeight(HPDF_Doc pdf, HPDF_Page page, const TableCell &cell, float cellWidth)
{
    float totalHeight = 0.0f;

    for (const auto &fragment : cell.textFragments)
    {
        // Set font and size
        HPDF_Page_SetFontAndSize(page, fragment.font, fragment.fontSize);

        // Handle line breaks
        std::string text = fragment.text;
        std::replace(text.begin(), text.end(), '\n', ' ');

        // Calculate height needed for this fragment
        float fragmentHeight = calculateTextHeight(pdf, page, text,
                                                   fragment.fontSize, fragment.font, cellWidth - 10); // Subtract padding
        totalHeight += fragmentHeight;
    }

    // Ensure minimum cell height
    float defaultLineHeight = 20.0f; // Adjust as needed
    return std::max(totalHeight, defaultLineHeight);
}

void renderTextInCell(HPDF_Doc pdf, HPDF_Page &page, const std::string &text,
                      float &cursorX, float &cursorY, float cellWidth,
                      float fontSize, HPDF_Font font, float bottomY)
{
    size_t pos = 0;
    size_t len = text.length();
    float initialX = cursorX; // Save initial X position

    while (pos < len)
    {
        // Handle line breaks
        if (text[pos] == '\n')
        {
            cursorY -= fontSize + 2.0f;
            cursorX = initialX; // Reset to left edge of cell
            pos++;

            // Stop rendering if we exceed the bottom of the cell
            if (cursorY < bottomY)
            {
                break;
            }
            continue;
        }

        // Find the next whitespace or non-whitespace run
        size_t nextPos = pos;
        bool isSpace = isspace(static_cast<unsigned char>(text[pos]));

        while (nextPos < len && isspace(static_cast<unsigned char>(text[nextPos])) == isSpace && text[nextPos] != '\n')
            nextPos++;

        // Extract the token
        std::string token = text.substr(pos, nextPos - pos);
        float tokenWidth = HPDF_Page_TextWidth(page, token.c_str());

        // If word doesn't fit on the current line
        if ((cursorX - initialX) + tokenWidth > cellWidth && !isSpace)
        {
            cursorY -= fontSize + 2.0f;
            cursorX = initialX; // Reset to left edge of cell

            // Stop rendering if we exceed the bottom of the cell
            if (cursorY < bottomY)
            {
                break;
            }
        }

        // Render the token
        HPDF_Page_BeginText(page);
        HPDF_Page_MoveTextPos(page, cursorX, cursorY);
        HPDF_Page_ShowText(page, token.c_str());
        HPDF_Page_EndText(page);

        cursorX += tokenWidth;
        pos = nextPos;
    }
}

void renderTable(HPDF_Doc pdf, HPDF_Page &page, const Table &table,
                 float &cursorX, float &cursorY, float pageWidth,
                 float leftMargin, float rightMargin)
{
    if (table.rows.empty()) return;

    // Table properties
    float tableStartX = leftMargin;
    float tableWidth = pageWidth - leftMargin - rightMargin;

    // Determine the maximum number of columns considering gridSpans
    size_t numCols = 0;
    for (const auto &row : table.rows)
    {
        size_t colCount = 0;
        for (const auto &cell : row)
        {
            colCount += cell.gridSpan;
        }
        numCols = std::max(numCols, colCount);
    }

    if (numCols == 0)
    {
        std::cerr << "Error: Table has zero columns." << std::endl;
        return;
    }

    // Define column widths (evenly distributed)
    std::vector<float> colWidths(numCols, tableWidth / numCols);

    // Iterate over each row
    for (const auto &row : table.rows)
    {
        // Calculate required height for the row
        float maxCellHeight = 0.0f;
        std::vector<float> cellHeights;
        std::vector<float> cellWidths;

        // First pass: calculate cell heights and widths
        size_t colIndex = 0;
        for (const auto &cell : row)
        {
            size_t span = cell.gridSpan;
            float cellWidth = 0.0f;
            for (size_t i = 0; i < span && (colIndex + i) < numCols; ++i)
            {
                cellWidth += colWidths[colIndex + i];
            }
            cellWidths.push_back(cellWidth);

            float cellHeight = calculateCellHeight(pdf, page, cell, cellWidth);
            cellHeights.push_back(cellHeight);
            maxCellHeight = std::max(maxCellHeight, cellHeight);

            colIndex += span;
        }

        // Handle page break if necessary
        if (cursorY - maxCellHeight < 50)
        {
            page = HPDF_AddPage(pdf);
            HPDF_Page_SetSize(page, HPDF_PAGE_SIZE_A4, HPDF_PAGE_PORTRAIT);
            cursorY = HPDF_Page_GetHeight(page) - 50;
        }

        // Draw horizontal line for the top of the row
        HPDF_Page_SetRGBStroke(page, 0, 0, 0); // Black color for borders
        HPDF_Page_SetLineWidth(page, 0.5);
        HPDF_Page_MoveTo(page, tableStartX, cursorY);
        HPDF_Page_LineTo(page, tableStartX + tableWidth, cursorY);
        HPDF_Page_Stroke(page);

        float cellX = tableStartX;
        colIndex = 0;

        // Draw vertical lines at cell boundaries
        std::vector<float> cellBoundaries = {cellX};

        for (const auto &cell : row)
        {
            size_t span = cell.gridSpan;
            float cellWidth = 0.0f;
            for (size_t i = 0; i < span && (colIndex + i) < numCols; ++i)
            {
                cellWidth += colWidths[colIndex + i];
            }

            cellX += cellWidth;
            cellBoundaries.push_back(cellX);
            colIndex += span;
        }

        // Draw vertical lines
        for (float x : cellBoundaries)
        {
            HPDF_Page_MoveTo(page, x, cursorY);
            HPDF_Page_LineTo(page, x, cursorY - maxCellHeight);
            HPDF_Page_Stroke(page);
        }

        // Draw vertical lines for any remaining columns
        while (colIndex < numCols)
        {
            cellX += colWidths[colIndex];
            HPDF_Page_MoveTo(page, cellX, cursorY);
            HPDF_Page_LineTo(page, cellX, cursorY - maxCellHeight);
            HPDF_Page_Stroke(page);
            colIndex++;
        }

        // Render cell content
        cellX = tableStartX;
        colIndex = 0;
        size_t cellIndex = 0;

        for (const auto &cell : row)
        {
            size_t span = cell.gridSpan;
            float cellWidth = cellWidths[cellIndex];

            // Calculate the bottom Y coordinate for the cell
            float cellBottomY = cursorY - maxCellHeight;

            // Render cell content
            float textCursorX = cellX + 5; // Padding from left
            float textCursorY = cursorY - 5; // Start from top of cell, adjust padding
            float bottomY = cellBottomY + 5; // Adjust padding

            for (const auto &fragment : cell.textFragments)
            {
                // Set font and color
                HPDF_Page_SetFontAndSize(page, fragment.font, fragment.fontSize);
                HPDF_Page_SetRGBFill(page, fragment.r, fragment.g, fragment.b);

                float availableWidth = cellWidth - 10; // Subtract padding

                // Render text within the cell
                float tempCursorX = textCursorX;
                float tempCursorY = textCursorY;

                renderTextInCell(pdf, page, fragment.text, tempCursorX, tempCursorY,
                                 availableWidth, fragment.fontSize, fragment.font, bottomY);

                textCursorY = tempCursorY; // Update textCursorY after rendering
            }

            cellX += cellWidth;
            colIndex += span;
            cellIndex++;
        }

        // Draw horizontal line for the bottom of the row
        HPDF_Page_MoveTo(page, tableStartX, cursorY - maxCellHeight);
        HPDF_Page_LineTo(page, tableStartX + tableWidth, cursorY - maxCellHeight);
        HPDF_Page_Stroke(page);

        // Move cursorY for the next row
        cursorY -= maxCellHeight;
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
                        const char* spacePreserve = child->Attribute("xml:space");
                        std::string textContent = child->GetText();

                        if (spacePreserve && strcmp(spacePreserve, "preserve") == 0)
                        {
                            text = textContent;
                        }
                        else
                        {
                            text = textContent;
                        }
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
                HPDF_Page_SetRGBFill(page, r, g, b);

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
    const char *boldItalicFontName = HPDF_LoadTTFontFromFile(pdf, (std::string(fontPath) + "DejaVuSans-BoldOblique.ttf").c_str(), HPDF_TRUE);

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
