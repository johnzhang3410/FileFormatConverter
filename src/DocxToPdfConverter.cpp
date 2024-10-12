#include "DocxToPdfConverter.h"
#include <hpdf.h>
#include <tinyxml2.h>
#include <iostream>
#include <fstream>

// Include any other necessary headers
#include <sstream>
#include <iomanip>

using namespace tinyxml2;

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

//generates pdf from the parsed docx content
void generatePDF(const std::string &docxDir, const std::string &outputPdfPath)
{
    HPDF_Doc pdf = HPDF_New(NULL, NULL);

    if (!pdf)
    {
        std::cerr << "Failed to create PDF object." << std::endl;
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
    
    //sets font, left margin and right margin
    HPDF_Font defaultFont = HPDF_GetFont(pdf, "Helvetica", NULL);

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
                    font = HPDF_GetFont(pdf, "Helvetica-BoldOblique", NULL);
                }
                else if (isBold)
                {
                    font = HPDF_GetFont(pdf, "Helvetica-Bold", NULL);
                }
                else if (isItalic)
                {
                    font = HPDF_GetFont(pdf, "Helvetica-Oblique", NULL);
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