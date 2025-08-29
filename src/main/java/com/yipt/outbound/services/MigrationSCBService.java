package com.yipt.outbound.services;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Component;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
@Component
public class MigrationSCBService {

    private final String filePath="C:/Users/Toonch/Downloads/For_convert_open.xlsx";

    public void run() {
        try (FileInputStream fis = new FileInputStream(filePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            // Loop through all sheets
            for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
                XSSFSheet sheet = workbook.getSheetAt(s);

                // Loop through rows, convert column to HTML
                for (Row row : sheet) {
                    for (Cell cell : row) {
                        if (cell.getCellType() == CellType.STRING) {
//                            XSSFRichTextString richText = (XSSFRichTextString) cell.getRichStringCellValue();
//                            if (richText.numFormattingRuns() > 0) {
                                // Only convert cells with rich text formatting
                                String html = convertCellToHtml(cell);
                                cell.setCellValue(html);
//                            }
                        }
                    }
                }
            }

            // Save to a new file
            try (FileOutputStream fos = new FileOutputStream("C:/Users/Toonch/Downloads/For_convert_open_html.xlsx")) {
                workbook.write(fos);
            }

            System.out.println("Conversion complete. Saved to new file.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    private String convertCellToHtml(Cell cell) {
        if (cell == null) return "<p></p>";

        String text;
        switch (cell.getCellType()) {
            case STRING:
                text = ((XSSFRichTextString) cell.getRichStringCellValue()).getString();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    text = new SimpleDateFormat("yyyy-MM-dd").format(cell.getDateCellValue());
                } else {
                    text = String.valueOf(cell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                text = String.valueOf(cell.getBooleanCellValue());
                break;
            case FORMULA:
                text = cell.getCellFormula();
                break;
            default:
                text = "";
        }

        if (!(cell.getCellType() == CellType.STRING)) {
            return "<p>" + escapeHtml(text) + "</p>"; // wrap non-string too
        }

        XSSFRichTextString richText = (XSSFRichTextString) cell.getRichStringCellValue();
        StringBuilder html = new StringBuilder();

        int length = richText.length();
        int numRuns = richText.numFormattingRuns();

        if (numRuns == 0) {
            return "<p>" + escapeHtml(richText.getString()) + "</p>";
        }

        int currentIndex = 0;
        for (int i = 0; i < numRuns; i++) {
            int runIndex = richText.getIndexOfFormattingRun(i);

            if (runIndex > currentIndex) {
                html.append(escapeHtml(richText.getString().substring(currentIndex, runIndex)));
            }

            int nextIndex = (i + 1 < numRuns) ? richText.getIndexOfFormattingRun(i + 1) : length;
            XSSFFont font = richText.getFontOfFormattingRun(i);
            String textRun = richText.getString().substring(runIndex, nextIndex);

            html.append(applyFontHtml(textRun, font));
            currentIndex = nextIndex;
        }

        if (currentIndex < length) {
            html.append(escapeHtml(richText.getString().substring(currentIndex)));
        }

        return "<p>" + html.toString() + "</p>";
    }



    private String applyFontHtml(String text, XSSFFont font) {
        if (font == null) return escapeHtml(text);

        StringBuilder sb = new StringBuilder();

        // Open formatting tags
        if (font.getBold()) sb.append("<strong>");
        if (font.getItalic()) sb.append("<em>");
        if (font.getUnderline() != Font.U_NONE) sb.append("<u>");
        if (font.getStrikeout()) sb.append("<s>");

        // Span for font size/color
        sb.append("<span style=\"");
        sb.append("font-size:").append(font.getFontHeightInPoints()).append("pt;");
        if (font.getXSSFColor() != null) {
            byte[] rgb = font.getXSSFColor().getRGB();
            if (rgb != null && rgb.length == 3) {
                sb.append("color:#")
                  .append(String.format("%02X%02X%02X", rgb[0], rgb[1], rgb[2]))
                  .append(";");
            }
        }
        sb.append("\">");

        sb.append(escapeHtml(text));

        sb.append("</span>");

        // Close formatting tags in reverse order
        if (font.getStrikeout()) sb.append("</s>");
        if (font.getUnderline() != Font.U_NONE) sb.append("</u>");
        if (font.getItalic()) sb.append("</em>");
        if (font.getBold()) sb.append("</strong>");

        return sb.toString();
    }


    private String escapeHtml(String s) {
        return s.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("\r\n", "<br/>")   // Windows line break (Alt+Enter in Excel)
                .replace("\n", "<br/>")     // Unix line break
                .replace("\r", "<br/>");    // Old Mac line break
    }
}
