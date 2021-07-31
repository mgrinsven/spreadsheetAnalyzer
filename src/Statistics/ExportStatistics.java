package Statistics;

import org.apache.commons.math3.geometry.euclidean.threed.Rotation;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Exporting statistic to Excel
 */
public class ExportStatistics {
    private Font headerFont, totalsFont;
    private CellStyle headerCellStyle, totalsCellStyle;
    private final int filesColSize = 60*256;
    private final int sourceColSize = 35*256;
    private final int obsColSize = 25*256;

    public ExportStatistics() {
    }

    /**
     * Starts the export of the gathered data to excel
     * @param sStats    The gathered spreadsheet data
     * @param omitEmpty Wether to discard spreadsheets without VBA source files or without logic
     */
    public void createSpreadsheet(SpreadsheetStatistics sStats, boolean omitEmpty) {
        Workbook workbook = new XSSFWorkbook();
        CreationHelper createHelper = workbook.getCreationHelper();
        createStyles(workbook);
        createSummarySheet(workbook, sStats);
        createScanResultSheet(workbook, sStats, omitEmpty);
        createDetailedResults(workbook, sStats, omitEmpty, false);
        createDetailedResults(workbook, sStats, omitEmpty, true);
        writeSpreadhseet(workbook);
    }

    /**
     * Creates the styles used in the generated spreadsheets
     * @param workbook  The workbook used for  the styles
     */
    public void createStyles(Workbook workbook) {
        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        //headerFont.setColor(IndexedColors.RED.getIndex());
        // Create a CellStyle with the font
        headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
        headerCellStyle.setRotation((short) 45);

        //Total
        Font totalsFont = workbook.createFont();
        totalsFont.setBold(true);
        totalsCellStyle = workbook.createCellStyle();
        totalsCellStyle.setFont(totalsFont);
    }

    /**
     * Creates a sheet with scan results per spreadsheet
     * @param workbook
     * @param sStats
     * @param omitEmpty
     */
    public void createScanResultSheet(Workbook workbook, SpreadsheetStatistics sStats, boolean omitEmpty) {
        String[] columns = {"Excel filename", "VBA Source files", "Empty source files", "Macro source files", "Source files with Subs/Funcs", "Source files with ext refs", "Source files with credentials"};

        Sheet scanResults = workbook.createSheet("Scan results");
        scanResults.setColumnWidth(0, filesColSize);

        // Create a Row
        Row headerRow = scanResults.createRow(0);

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
                Row row = scanResults.createRow(rowNum++);
                row.createCell(0).setCellValue(ssArray.getSpreadsheetFile().getName());
                row.createCell(1).setCellValue(ssArray.getTotals().totalFiles);
                row.createCell(2).setCellValue((ssArray.getTotals().emptyFileCount));
                row.createCell(3).setCellValue(ssArray.getTotals().containsMacroFileCount);
                row.createCell(4).setCellValue(ssArray.getTotals().containsCodeFileCount);
                row.createCell(5).setCellValue(ssArray.getTotals().containsExtRefsFileCount);
                row.createCell(6).setCellValue(ssArray.getTotals().containsCredentialsFileCount);
            }
        }
        // calculate totals
        int startRow = 2;
        Row totalsRow = scanResults.createRow(rowNum++);
        totalsRow.setRowStyle(totalsCellStyle);
        totalsRow.createCell(0).setCellValue("Totals");
        totalsRow.createCell(1).setCellFormula(String.format("SUM(B%d:B%d)", startRow, rowNum-1));
        totalsRow.createCell(2).setCellFormula(String.format("SUM(C%d:C%d)", startRow, rowNum-1));
        totalsRow.createCell(3).setCellFormula(String.format("SUM(D%d:D%d)", startRow, rowNum-1));
        totalsRow.createCell(4).setCellFormula(String.format("SUM(E%d:E%d)", startRow, rowNum-1));
        totalsRow.createCell(5).setCellFormula(String.format("SUM(F%d:F%d)", startRow, rowNum-1));
        totalsRow.createCell(6).setCellFormula(String.format("SUM(G%d:G%d)", startRow, rowNum-1));

        scanResults.createFreezePane(0,1);
    }

    /**
     * Create a sheet with the summary of he gathered results
     * @param workbook
     * @param stats
     */
    private void createSummarySheet(Workbook workbook, SpreadsheetStatistics stats) {
        Sheet sheet = workbook.createSheet("Summary");

        String sumColsHeader[] = {"Total files scanned", "Total files with PQ", "Total files containing VBA", "Total VBA files detected", "Total VBA files containing code", "Total VBA files containing macros", "Total empty VBA files", "Total VBA files containing credentials"};
        SpreadsheetTotals totals = stats.getTotals();
        int numChars = 40;
        sheet.setColumnWidth(0, numChars*256);

        // Create header row
        Row headerRow = sheet.createRow(0);

        // Create summary rows
        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue(sumColsHeader[0]);
        row1.createCell(1).setCellValue(stats.getTotalFiles());

        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue(sumColsHeader[1]);
        row2.createCell(1).setCellValue(totals.totalPowerQuery);

        Row row3 = sheet.createRow(3);
        row3.createCell(0).setCellValue(sumColsHeader[2]);
        row3.createCell(1).setCellValue(totals.totalVBASpreadsheets);

        Row row4 = sheet.createRow(4);
        row4.createCell(0).setCellValue(sumColsHeader[3]);
        row4.createCell(1).setCellValue(totals.totalFiles);

        Row row5 = sheet.createRow(5);
        row5.createCell(0).setCellValue(sumColsHeader[4]);
        row5.createCell(1).setCellValue(totals.containsCodeFileCount);

        Row row6 = sheet.createRow(6);;
        row6.createCell(0).setCellValue(sumColsHeader[5]);
        row6.createCell(1).setCellValue(totals.containsMacroFileCount);

        Row row7 = sheet.createRow(7);;
        row7.createCell(0).setCellValue(sumColsHeader[6]);
        row7.createCell(1).setCellValue(totals.emptyFileCount);

        Row row8 = sheet.createRow(8);;
        row8.createCell(0).setCellValue(sumColsHeader[7]);
        row8.createCell(1).setCellValue(totals.containsCredentialsFileCount);
    }

    /**
     * Creates a sheet in the workbook with details per VBA source file, grouped by spreadsheet
     * @param workbook
     * @param stats
     * @param omitEmpty Boolean when true, spreadsheets with empty files are disgarded
     * @param showObs   Boolean when true it shows all individual observations.
     */
    public void createDetailedResults(Workbook workbook, SpreadsheetStatistics stats, boolean omitEmpty, boolean showObs) {
        String headers[] = {"Spreadsheet", "Source file", "Nr. of subs/func", "Nr of macros", "Nr of ext refs", "Nr of credential refs"};
        String headersObs[] = {"Spreadsheet", "Source file", "Observation", "Start line", "End line", "Details"};
        int headerLength = 0;
        Sheet sheet;

        if (showObs) {
            headerLength = headersObs.length;
        } else {
            headerLength = headers.length;
        }

        if (showObs) {
            sheet = workbook.createSheet("Source observations");
        } else {
            sheet = workbook.createSheet("Detailed results");
        }
        Row headerRow = sheet.createRow(0);
        sheet.createFreezePane(0,1);
        sheet.setColumnWidth(0, filesColSize);
        sheet.setColumnWidth(1, sourceColSize);
        if (showObs) {
            sheet.setColumnWidth(2, obsColSize);
        }
        // Create cells
        for(int i = 0; i < headerLength; i++) {
            Cell cell = headerRow.createCell(i);
            if (showObs) {
                cell.setCellValue(headersObs[i]);
            } else {
                cell.setCellValue(headers[i]);
            }
            cell.setCellStyle(headerCellStyle);
        }

        int rowNum = 1;
        for (SourceFileStatsArray ssArray:stats.spreadsheetStats) {
            boolean writeRow = true;
            if (omitEmpty) {
                if (ssArray.getTotals().totalFiles == 0) {
                    writeRow = false;
                }
            }
            // Show spreadsheet Filename
            if (writeRow) {
                boolean newSpreadsheet = true;
                for (SourceFileStats sfStats: ssArray.getSourceFileStats()) {
                    boolean newSourceFile = true;
                    Row row = sheet.createRow(rowNum++);
                    if (newSpreadsheet) {
                        row.createCell(0).setCellValue(ssArray.getSpreadsheetFile().getName());
                        newSpreadsheet = false;
                    }
                    if (showObs) {
                        for (Observations pObs: sfStats.getParserObservations().getObservationList()) {
                            if (newSourceFile) {
                                row.createCell(1).setCellValue(sfStats.getSourceFile().getName());
                                newSourceFile = false;
                            } else {
                                row = sheet.createRow(rowNum++);
                            }
                            row.createCell(2).setCellValue(pObs.getObservation());
                            row.createCell(3).setCellValue(pObs.startLine);
                            row.createCell(4).setCellValue(pObs.endLine);
                            row.createCell(5).setCellValue(pObs.subject);
                        }
                    } else {
                        row.createCell(1).setCellValue(sfStats.getSourceFile().getName());
                        row.createCell(2).setCellValue(sfStats.getTotals().containsCodeFileCount);
                        row.createCell(3).setCellValue(sfStats.getTotals().containsMacroFileCount);
                        row.createCell(4).setCellValue(sfStats.getTotals().containsExtRefsFileCount);
                        row.createCell(5).setCellValue(sfStats.getTotals().containsCredentialsFileCount);
                    }
                }
            }
        }
    }

    /**
     * Writes the workboos with sheets to an Excel file.
     * @param workbook
     */
    public void writeSpreadhseet (Workbook workbook) {
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
