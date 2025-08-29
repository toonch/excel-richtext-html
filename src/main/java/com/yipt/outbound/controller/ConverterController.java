package com.yipt.outbound.controller;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
public class ConverterController {

    private static final Logger logger = LoggerFactory.getLogger(ConverterController.class);

    @PostMapping("/convert")
    public ResponseEntity<byte[]> convert(@RequestParam("file") MultipartFile file) {
        logger.info("Received file: {}, size={}", file.getOriginalFilename(), file.getSize());
        try (
            XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
            ByteArrayOutputStream bos = new ByteArrayOutputStream()
        ) {
            logger.info("Processing workbook with {} sheet(s)", workbook.getNumberOfSheets());
            for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
                XSSFSheet sheet = workbook.getSheetAt(s);
                logger.info("Processing sheet: {}", sheet.getSheetName());
                for (Row row : sheet) {
                    for (Cell cell : row) {
                        if (cell.getCellType() == CellType.STRING) {
                            String html = convertCellToHtml(cell);
                            cell.setCellValue(html);
                        }
                    }
                }
            }
            workbook.write(bos);
            logger.info("Workbook conversion completed for file: {}", file.getOriginalFilename());
            return ResponseEntity.ok()
                    .header("Content-Disposition", "attachment; filename=converted.xlsx")
                    .body(bos.toByteArray());
        } catch (Exception e) {
            logger.error("Error converting file: {} - {}", file.getOriginalFilename(), e.getMessage(), e);
            return ResponseEntity.status(500).body(null);
        }
    }

    private String convertCellToHtml(Cell cell) {
        if (cell == null || cell.getCellType() != CellType.STRING) {
            return "<p></p>";
        }

        XSSFRichTextString richText = (XSSFRichTextString) cell.getRichStringCellValue();
        String text = richText.getString();
        int length = text.length();
        int numRuns = richText.numFormattingRuns();

        StringBuilder html = new StringBuilder();
        int currentIndex = 0;

        if (numRuns == 0) {
            html.append(escapeHtml(text));
        } else {
            for (int i = 0; i < numRuns; i++) {
                int runIndex = richText.getIndexOfFormattingRun(i);
                if (runIndex < 0 || runIndex > length) continue;

                int nextIndex = (i + 1 < numRuns) ? richText.getIndexOfFormattingRun(i + 1) : length;
                nextIndex = Math.min(nextIndex, length);

                if (runIndex > currentIndex) {
                    html.append(escapeHtml(text.substring(currentIndex, runIndex)));
                }

                if (runIndex < nextIndex) {
                    String runText = text.substring(runIndex, nextIndex);
                    XSSFFont font = richText.getFontOfFormattingRun(i);

                    if (font != null && hasStyle(font)) {
                        html.append(applyFontHtml(runText, font));
                    } else {
                        html.append(escapeHtml(runText));
                    }
                }

                currentIndex = nextIndex;
            }

            if (currentIndex < length) {
                html.append(escapeHtml(text.substring(currentIndex)));
            }
        }

        return "<p>" + html.toString() + "</p>";
    }

    private boolean hasStyle(XSSFFont font) {
        return font.getBold() || font.getItalic() || font.getUnderline() != Font.U_NONE || font.getStrikeout()
               || (font.getFontHeightInPoints() != 11)
               || (font.getXSSFColor() != null);
    }

    private String applyFontHtml(String text, XSSFFont font) {
        if (font == null) return escapeHtml(text);

        StringBuilder sb = new StringBuilder();

        if (font.getBold()) sb.append("<strong>");
        if (font.getItalic()) sb.append("<em>");
        if (font.getUnderline() != Font.U_NONE) sb.append("<u>");
        if (font.getStrikeout()) sb.append("<s>");

        sb.append("<span style=\"");
        sb.append("font-size:").append(font.getFontHeightInPoints()).append("pt;");
        if (font.getXSSFColor() != null) {
            byte[] rgb = font.getXSSFColor().getRGB();
            if (rgb != null && rgb.length == 3) {
                sb.append("color:#").append(String.format("%02X%02X%02X", rgb[0], rgb[1], rgb[2])).append(";");
            }
        }
        sb.append("\">");

        sb.append(escapeHtml(text));

        sb.append("</span>");

        if (font.getStrikeout()) sb.append("</s>");
        if (font.getUnderline() != Font.U_NONE) sb.append("</u>");
        if (font.getItalic()) sb.append("</em>");
        if (font.getBold()) sb.append("</strong>");

        return sb.toString();
    }

    private String escapeHtml(String s) {
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\"", "&quot;")
                .replace("\r\n", "<br/>")
                .replace("\n", "<br/>")
                .replace("\r", "<br/>");
    }
}
