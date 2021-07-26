package Statistics;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Exporting statistic to Excel
 */
public class ExportStatistics {
    private static String[] columns = {"Excel filename", "VBA Source files", "Empty source files", "Macro source files", "Source files with Subs/Funcs", "Source files with ext refs", "Source files with credentials"};

    public void createSpreadsheet(SpreadsheetStatistics sStats, boolean omitEmpty) {
        Workbook workbook = new XSSFWorkbook();
        CreationHelper createHelper = workbook.getCreationHelper();
        Sheet sheet = workbook.createSheet("Scan results");

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        // Create cells
        for(int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }
        int rowNum = 1;
        for (SourceFileStatsArray ssArray:sStats.spreadsheetStats) {
            boolean writeRow = true;
            if (omitEmpty) {
                if (ssArray.getTotals().totalFiles == 0) {
                    writeRow = false;
                }
            }
            // Show spreadsheet Filename
            if (writeRow) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(ssArray.getSpreadsheetFile().getName());
                row.createCell(1).setCellValue(ssArray.getTotals().totalFiles);
                row.createCell(2).setCellValue((ssArray.getTotals().emptyFileCount));
                row.createCell(3).setCellValue(ssArray.getTotals().containsMacroFileCount);
                row.createCell(4).setCellValue(ssArray.getTotals().containsCodeFileCount);
                row.createCell(5).setCellValue(ssArray.getTotals().containsExtRefsFileCount);
                row.createCell(6).setCellValue(ssArray.getTotals().containsCredentialsFileCount);
            }
        }
        try {
            // Write the output to a file
            FileOutputStream fileOut = new FileOutputStream("poi-generated-file.xlsx");
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
        } catch (IOException e) {
            System.out.printf("Error during writing of generated spreadsheet:\n\t%s", e.getMessage());
        }
    }
}
