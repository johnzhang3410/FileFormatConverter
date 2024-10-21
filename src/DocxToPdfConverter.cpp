#include "DocxToPdfConverter.h"
#include <hpdf.h>
#include <tinyxml2.h>
#include <iostream>
#include <fstream>

// Include any other necessary headers
#include <sstream>
#include <iomanip>

using namespace tinyxml2;

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

void renderTextWithWrapping(HPDF_Doc pdf, HPDF_Page &page, const std::string &text, float &cursorX, float &cursorY, float pageWidth, float fontSize, HPDF_Font font, float leftMargin, float rightMargin)
{
    size_t pos = 0;
    size_t len = text.length();

    while (pos < len)
    {
        // Find the next whitespace or non-whitespace run (so each token is either bunch of char or bunch of spaces)
        size_t nextPos = pos;
        bool isSpace = isspace(text[pos]);

        while (nextPos < len && isspace(text[nextPos]) == isSpace)
            nextPos++;

        // Extract the token (word or spaces)
        std::string token = text.substr(pos, nextPos - pos);
        float tokenWidth = HPDF_Page_TextWidth(page, token.c_str());

        // if word doesn't fit on the current line 
        if (cursorX + tokenWidth > pageWidth - rightMargin && !isSpace)
        {
            cursorY -= fontSize + 2.0f;
            cursorX = leftMargin;

            // handle page break
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

// Parse a table element and populate the Table structure
Table parseTable(XMLElement *tblElement, HPDF_Font defaultFont, HPDF_Font boldFont, HPDF_Font italicFont, HPDF_Font boldItalicFont) {
    Table table;

    for (XMLElement *tr = tblElement->FirstChildElement("w:tr"); tr; tr = tr->NextSiblingElement("w:tr")) {
        std::vector<TableCell> row;

        for (XMLElement *tc = tr->FirstChildElement("w:tc"); tc; tc = tc->NextSiblingElement("w:tc")) {
            TableCell cell;

            // Iterate over paragraphs within the cell
            for (XMLElement *para = tc->FirstChildElement("w:p"); para; para = para->NextSiblingElement("w:p")) {
                for (XMLElement *run = para->FirstChildElement("w:r"); run; run = run->NextSiblingElement("w:r")) {
                    XMLElement *rPr = run->FirstChildElement("w:rPr");
                    bool isBold = false;
                    bool isItalic = false;
                    std::string color = "000000";
                    int fontSize = 12;
                    HPDF_Font font = defaultFont;

                    if (rPr) {
                        if (rPr->FirstChildElement("w:b")) {
                            isBold = true;
                        }
                        if (rPr->FirstChildElement("w:i")) {
                            isItalic = true;
                        }
                        XMLElement *colorElement = rPr->FirstChildElement("w:color");
                        if (colorElement && colorElement->Attribute("w:val")) {
                            color = colorElement->Attribute("w:val");
                        }
                        XMLElement *szElement = rPr->FirstChildElement("w:sz");
                        if (szElement && szElement->Attribute("w:val")) {
                            fontSize = std::stoi(szElement->Attribute("w:val")) / 2;
                        }
                    }

                    if (isBold && isItalic) {
                        font = boldItalicFont;
                    }
                    else if (isBold) {
                        font = boldFont;
                    }
                    else if (isItalic) {
                        font = italicFont;
                    }

                    // Convert color to RGB
                    float r = 0, g = 0, b = 0;
                    if (color.length() == 6) {
                        std::stringstream ss;
                        ss << std::hex << color;
                        unsigned int rgb;
                        ss >> rgb;
                        r = ((rgb >> 16) & 0xFF) / 255.0f;
                        g = ((rgb >> 8) & 0xFF) / 255.0f;
                        b = (rgb & 0xFF) / 255.0f;
                    }

                    // Iterate over child elements within the run
                    for (XMLElement *child = run->FirstChildElement(); child; child = child->NextSiblingElement()) {
                        if (strcmp(child->Name(), "w:t") == 0) {
                            // Text element
                            if (child->GetText()) {
                                const char* spacePreserve = child->Attribute("xml:space");
                                std::string textContent = child->GetText();

                                std::string text;
                                if (spacePreserve && strcmp(spacePreserve, "preserve") == 0) {
                                    text = textContent;
                                }
                                else {
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
                        else if (strcmp(child->Name(), "w:drawing") == 0) {
                            // Handle images within table cells if needed
                            continue;
                        }
                        else if (strcmp(child->Name(), "w:br") == 0) {
                            // Line break within table cell
                            continue;
                        }
                    }
                }
            }

            row.push_back(cell);
        }
        //pushback row after all cells in the row have been processed 
        table.rows.push_back(row);
    }

    return table;
}

// Render a table using libharu
void renderTable(HPDF_Doc pdf, HPDF_Page &page, const Table &table,
                float &cursorX, float &cursorY, float pageWidth,
                float leftMargin, float rightMargin) {
    if (table.rows.empty()) return;

    //table properties
    float tableStartX = leftMargin;
    float tableStartY = cursorY;
    float tableWidth = pageWidth - leftMargin - rightMargin;
    size_t numCols = table.rows[0].size();

    // Define column widths (evenly distributed)
    std::vector<float> colWidths(numCols, tableWidth / numCols);

    // Define row height (fixed for simplicity)
    float rowHeight = 20.0f; 

    // Draw table borders and content
    for (const auto &row : table.rows) {
        // Draw horizontal line for the row
        HPDF_Page_SetRGBStroke(page, 0, 0, 0); // Black color for borders
        HPDF_Page_SetLineWidth(page, 0.5);
        HPDF_Page_MoveTo(page, tableStartX, cursorY);
        HPDF_Page_LineTo(page, tableStartX + tableWidth, cursorY);
        HPDF_Page_Stroke(page);

        float cellX = tableStartX;

        for (size_t i = 0; i < row.size(); ++i) {
            const TableCell &cell = row[i];

            // Draw vertical line for the cell
            HPDF_Page_MoveTo(page, cellX, cursorY);
            HPDF_Page_LineTo(page, cellX, cursorY - rowHeight);
            HPDF_Page_Stroke(page);

            // Render cell content
            float textCursorX = cellX + 5; // Padding from left
            float textCursorY = cursorY - 15; // Adjust as needed

            // Render text fragments within the cell
            for (const auto &fragment : cell.textFragments) {
                renderTextWithWrapping(pdf, page, fragment.text, textCursorX, textCursorY, pageWidth, fragment.fontSize, fragment.font, leftMargin, rightMargin);
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
        if (cursorY < 50) {
            page = HPDF_AddPage(pdf);
            HPDF_Page_SetSize(page, HPDF_PAGE_SIZE_A4, HPDF_PAGE_PORTRAIT);
            cursorY = HPDF_Page_GetHeight(page) - 50;
        }
    }

    // After table, adjust cursorY
    cursorY -= 10.0f; // Space after table
}

//generates pdf from the parsed docx content
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

    const char *fontName = HPDF_LoadTTFontFromFile(pdf, "../fonts/dejavu-fonts-ttf/ttf/DejaVuSans.ttf", HPDF_TRUE);
    if (!fontName) {
        std::cerr << "Failed to load TrueType font." << std::endl;
        HPDF_Free(pdf);
        return;
    }

    // Load bold font
    const char *boldFontName = HPDF_LoadTTFontFromFile(pdf, "../fonts/dejavu-fonts-ttf/ttf/DejaVuSans-Bold.ttf", HPDF_TRUE);
    HPDF_Font boldFont = HPDF_GetFont(pdf, boldFontName, "UTF-8");

    // Load italic font
    const char *italicFontName = HPDF_LoadTTFontFromFile(pdf, "../fonts/dejavu-fonts-ttf/ttf/DejaVuSans-Oblique.ttf", HPDF_TRUE);
    HPDF_Font italicFont = HPDF_GetFont(pdf, italicFontName, "UTF-8");

    // Load bold italic font
    const char *boldItalicFontName = HPDF_LoadTTFontFromFile(pdf, "../fonts/dejavu-fonts-ttf/ttf/DejaVuSansCondensed-BoldOblique.ttf", HPDF_TRUE);
    HPDF_Font boldItalicFont = HPDF_GetFont(pdf, boldItalicFontName, "UTF-8");

    // Get the HPDF_Font object using the font name
    HPDF_Font defaultFont = HPDF_GetFont(pdf, fontName, "UTF-8");
    if (!defaultFont) {
        std::cerr << "Failed to get HPDF_Font object." << std::endl;
        HPDF_Free(pdf);
        return;
    }

    // Create a new page and set its size
    HPDF_Page page = HPDF_AddPage(pdf);
    HPDF_Page_SetSize(page, HPDF_PAGE_SIZE_A4, HPDF_PAGE_PORTRAIT);

    //setting page info and margins
    float leftMargin = 50.0f;
    float rightMargin = 50.0f;
    float cursorX = leftMargin;
    float cursorY = HPDF_Page_GetHeight(page) - 50;
    float pageWidth = HPDF_Page_GetWidth(page);
    float maxTextWidth = pageWidth - leftMargin - rightMargin;

    //tries to load and parse the document.xml file
    std::string documentXmlPath = docxDir + "/word/document.xml";
    XMLDocument doc;
    if (doc.LoadFile(documentXmlPath.c_str()) != XML_SUCCESS)
    {
        std::cerr << "Failed to load " << documentXmlPath << std::endl;
        HPDF_Free(pdf);
        return;
    }

    XMLElement *root = doc.RootElement(); // <w:document>
    XMLElement *body = root->FirstChildElement("w:body"); //content in the xml file

    for (XMLElement *para = body->FirstChildElement("w:p"); para; para = para->NextSiblingElement("w:p"))
    {
        int fontSize = 12;
        //for each run in the same paragraph, I set attributes like bold, italic, color and write the stylized text to pdf
        for (XMLElement *run = para->FirstChildElement("w:r"); run; run = run->NextSiblingElement("w:r"))
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
                //retrieves the color value if it exists
                if (colorElement && colorElement->Attribute("w:val"))
                {
                    color = colorElement->Attribute("w:val");
                }
                //same thing for font value
                XMLElement *szElement = rPr->FirstChildElement("w:sz");
                if (szElement && szElement->Attribute("w:val"))
                {
                    fontSize = std::stoi(szElement->Attribute("w:val")) / 2;
                }
            }

            // Iterate over all child elements within the run since it can contain text, tabs and br 
            // without this loop I would only be able to process the first element of the run 
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
                    continue; // no text to process so move onto next element
                }
                else if (strcmp(child->Name(), "w:br") == 0)
                {
                    // Line break element
                    cursorY -= fontSize + 2.0f;
                    cursorX = 50;

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

                unsigned int r = 0, g = 0, b = 0;
                //convert hex to rgb by extracting out the values and normalize
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
                    renderTextWithWrapping(pdf, page, text, cursorX, cursorY, pageWidth, fontSize, font, leftMargin, rightMargin);
                }
            }
        }

        float lineHeight = fontSize + 2.0f; 

        // Move to next line after paragraph
        cursorY -= lineHeight;
        cursorX = leftMargin;

        //in case of page break after the paragraph
        if (cursorY < 50)
        {
            page = HPDF_AddPage(pdf);
            HPDF_Page_SetSize(page, HPDF_PAGE_SIZE_A4, HPDF_PAGE_PORTRAIT);
            cursorY = HPDF_Page_GetHeight(page) - 50;
        }
    }

    if (HPDF_SaveToFile(pdf, outputPdfPath.c_str()) != HPDF_OK) {
        std::cerr << "Failed to save PDF to " << outputPdfPath << std::endl;
    }

    HPDF_Free(pdf);
}