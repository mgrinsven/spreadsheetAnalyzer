package Statistics;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;

/**
 * Exporting statistic to Excel
 */
public class ExportStatistics {
    //private Font headerFont, totalsFont;
    private XSSFCellStyle defaultCellStyle, headerCellStyle, totalsCellStyle, credentialsCellStyle, fileCellStyle, extrefCellStyle;
    private final int filesColSize = 60*256;
    private final int sourceColSize = 35*256;
    private final int checksumColSize = 42*256;
    private final int obsColSize = 25*256;

    public ExportStatistics() {
    }

    /**
     * Starts the export of the gathered data to excel
     * @param sStats    The gathered spreadsheet data
     * @param omitEmpty Wether to discard spreadsheets without VBA source files or without logic
     */
    public void createSpreadsheet(SpreadsheetStatistics sStats, boolean omitEmpty) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        //CreationHelper createHelper = workbook.getCreationHelper();
        createStyles(workbook);
        createSummarySheet(workbook, sStats);
        createScanResultSheet(workbook, sStats, omitEmpty);
        createDetailedResults(workbook, sStats, omitEmpty, false);
        createDetailedResults(workbook, sStats, omitEmpty, true);
        ArrayList<Integer> testResults = new ArrayList<Integer>(Arrays.asList(SourceFileStats.HAS_CREDENTIAL_REFS));
        createFilteredSheet(workbook, sStats, "Files with credentails", testResults );
        createChecksumOverview(workbook, sStats, "Checksum overview", testResults);
        writeSpreadhseet(workbook);
    }

    /**
     * Creates the styles used in the generated spreadsheets
     * @param workbook  The workbook used for  the styles
     */
    public void createStyles(XSSFWorkbook workbook) {
        byte[] bGrey = new byte[]{(byte) 235, (byte) 235, (byte) 235};
        byte[] bOrange = new byte[]{(byte)244, (byte)212, (byte)128};
        byte[] bRed = new byte[] {(byte)255, (byte)128, (byte)149};
        byte[] bBlue = new byte[] {(byte)128, (byte)191, (byte)255};

        XSSFColor grey = new XSSFColor(bGrey, null);
        XSSFColor orange = new XSSFColor(bOrange, null);
        XSSFColor red = new XSSFColor(bRed, null);
        XSSFColor blue = new XSSFColor(bBlue, null);

        // crrate defailt cellstyle
        defaultCellStyle = workbook.createCellStyle();

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        // Create a CellStyle with the font
        headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
        headerCellStyle.setRotation((short) 45);
        headerCellStyle.setFillForegroundColor(grey);
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);


        //Total
        Font totalsFont = workbook.createFont();
        totalsFont.setFontHeightInPoints((short) 14);
        totalsFont.setBold(true);
        totalsCellStyle = workbook.createCellStyle();
        totalsCellStyle.setFont(totalsFont);
        totalsCellStyle.setFillForegroundColor(grey);
        totalsCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // credentialsCellStyle
        credentialsCellStyle = workbook.createCellStyle();
        credentialsCellStyle.setFillForegroundColor(red);
        credentialsCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // fileCellStyle
        fileCellStyle = workbook.createCellStyle();
        //fileCellStyle.setFillBackgroundColor(new XSSFColor(new java.awt.Color(128, 191, 255)));
        fileCellStyle.setFillForegroundColor(blue);
        fileCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // extrefCellStyle
        extrefCellStyle = workbook.createCellStyle();
        extrefCellStyle.setFillForegroundColor(orange);
        //extrefCellStyle.setFillBackgroundColor(new XSSFColor(new java.awt.Color(255, 212, 128)));
        extrefCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    /**
     * Creates a sheet with scan results per spreadsheet
     * @param workbook
     * @param sStats
     * @param omitEmpty
     */
    public void createScanResultSheet(XSSFWorkbook workbook, SpreadsheetStatistics sStats, boolean omitEmpty) {
        String[] columns = {"Excel filename", "VBA Source files", "Empty source files", "Macro source files", "Source files with Subs/Funcs", "Source files with ext refs", "Source files with credentials"};

        Sheet scanResults = workbook.createSheet("Scan results");
        scanResults.setColumnWidth(0, filesColSize);

        // Create a Row
        Row headerRow = scanResults.createRow(0);

        fillHeader(columns, headerRow);

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
        //totalsRow.setRowStyle(totalsCellStyle);
        totalsRow.createCell(0).setCellValue("Totals");
        totalsRow.createCell(1).setCellFormula(String.format("SUM(B%d:B%d)", startRow, rowNum-1));
        totalsRow.createCell(2).setCellFormula(String.format("SUM(C%d:C%d)", startRow, rowNum-1));
        totalsRow.createCell(3).setCellFormula(String.format("SUM(D%d:D%d)", startRow, rowNum-1));
        totalsRow.createCell(4).setCellFormula(String.format("SUM(E%d:E%d)", startRow, rowNum-1));
        totalsRow.createCell(5).setCellFormula(String.format("SUM(F%d:F%d)", startRow, rowNum-1));
        totalsRow.createCell(6).setCellFormula(String.format("SUM(G%d:G%d)", startRow, rowNum-1));
        setRowStyle(totalsRow, 0, 6, totalsCellStyle);
        scanResults.createFreezePane(0,1);
    }

    /**
     * Create a sheet with the summary of the gathered results
     * @param workbook
     * @param stats
     */
    private void createSummarySheet(XSSFWorkbook workbook, SpreadsheetStatistics stats) {
        XSSFSheet sheet = workbook.createSheet("Summary");

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

        Row row6 = sheet.createRow(6);
        row6.createCell(0).setCellValue(sumColsHeader[5]);
        row6.createCell(1).setCellValue(totals.containsMacroFileCount);

        Row row7 = sheet.createRow(7);
        row7.createCell(0).setCellValue(sumColsHeader[6]);
        row7.createCell(1).setCellValue(totals.emptyFileCount);

        Row row8 = sheet.createRow(8);
        row8.createCell(0).setCellValue(sumColsHeader[7]);
        row8.createCell(1).setCellValue(totals.containsCredentialsFileCount);

        // Try to create a pie-chart
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 5, 2, 17, 20);

        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText("VBA code statistics");
        chart.setTitleOverlay(false);
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.RIGHT);

        String[] chartCat = {"Empty VBA source files", "VBA sources with macro's", "VBA sources with credentials", "VBA files with code"};
        Integer[] intData = { totals.emptyFileCount, totals.containsMacroFileCount, totals.containsCredentialsFileCount,
                              totals.containsCodeFileCount - totals.containsMacroFileCount - totals.containsCredentialsFileCount};

        XDDFCategoryDataSource cat = XDDFDataSourcesFactory.fromArray(chartCat);
        XDDFNumericalDataSource<Integer> val = XDDFDataSourcesFactory.fromArray(intData);

        XDDFChartData data = chart.createData(ChartTypes.PIE, null, null);
        data.setVaryColors(true);
        data.addSeries(cat, val);
        chart.plot(data);
    }

    /**
     * Creates a sheet in the workbook with details per VBA source file, grouped by spreadsheet
     * @param workbook
     * @param stats
     * @param omitEmpty Boolean when true, spreadsheets with empty files are disgarded
     * @param showObs   Boolean when true it shows all individual observations.
     */
    public void createDetailedResults(Workbook workbook, SpreadsheetStatistics stats, boolean omitEmpty, boolean showObs) {
        String headers[] = {"Spreadsheet", "Source file", "SHA1 checksum", "Nr. of subs/func", "Nr of macros", "Nr of ext refs", "Nr of credential refs"};
        String headersObs[] = {"Spreadsheet", "Source file", "SHA1 checksum", "Observation", "Start line", "End line", "Details"};
        Sheet sheet;

        if (showObs) {
            sheet = workbook.createSheet("Source observations");
            sheet.setColumnWidth(3, obsColSize);
            sheet.setColumnWidth(6, filesColSize);
        } else {
            sheet = workbook.createSheet("Detailed results");
        }
        Row headerRow = sheet.createRow(0);
        sheet.createFreezePane(0,1);
        sheet.setColumnWidth(0, filesColSize);
        sheet.setColumnWidth(1, sourceColSize);
        sheet.setColumnWidth(2, checksumColSize);

        if (showObs) {
            fillHeader(headersObs, headerRow);
        } else {
            fillHeader(headers, headerRow);
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
                                row.createCell(2).setCellValue(sfStats.getChecksum());
                                newSourceFile = false;
                            } else {
                                row = sheet.createRow(rowNum++);
                            }
                            row.createCell(3).setCellValue(pObs.getObservation());
                            row.createCell(4).setCellValue(pObs.startLine);
                            row.createCell(5).setCellValue(pObs.endLine);
                            row.createCell(6).setCellValue(pObs.subject);
                            setRowStyle(row, 3, 6, getStyle(pObs));
                        }
                    } else {
                        row.createCell(1).setCellValue(sfStats.getSourceFile().getName());
                        row.createCell(2).setCellValue(sfStats.getChecksum());
                        row.createCell(3).setCellValue(sfStats.getTotals().containsCodeFileCount);
                        row.createCell(4).setCellValue(sfStats.getTotals().containsMacroFileCount);
                        row.createCell(5).setCellValue(sfStats.getTotals().containsExtRefsFileCount);
                        row.createCell(6).setCellValue(sfStats.getTotals().containsCredentialsFileCount);
                    }
                }
            }
        }
    }

    /**
     *
     * @param workbook
     * @param stats
     * @param sheetName
     * @param parseResults
     */
    public void createFilteredSheet(Workbook workbook, SpreadsheetStatistics stats, String sheetName, ArrayList<Integer> parseResults) {
        String headers[] = {"Spreadsheet", "Source file", "SHA-1 Checksum", "Observation", "Start line", "End line", "Details"};
        Sheet sheet = workbook.createSheet(sheetName);
        Row headerRow = sheet.createRow(0);
        sheet.createFreezePane(0,1);
        sheet.setColumnWidth(0, filesColSize);
        sheet.setColumnWidth(1, sourceColSize);
        sheet.setColumnWidth(2, checksumColSize);
        sheet.setColumnWidth(3, obsColSize);
        sheet.setColumnWidth(6, filesColSize);

        fillHeader(headers, headerRow);

        int rowNum = 1;
        Row row = sheet.getRow(1);
        for (SourceFileStatsArray ssArray: stats.getSpreadsheetStats()) {
            boolean newSpreadsheet = true;
            for (SourceFileStats sfStats: ssArray.getSourceFileStats()) {
                boolean showSourceFile = false;
                boolean showSpreadsheet = false;
                boolean newSourceFile = true;
                for (Integer result: sfStats.getSourceFileStats()) {
                    for (Integer checkResult : parseResults) {
                        if (checkResult == result) {
                            showSpreadsheet = true;
                            showSourceFile = true;
                        }
                    }
                }
                if (newSpreadsheet && showSpreadsheet) {
                    row = sheet.createRow(rowNum++);
                    row.createCell(0).setCellValue(ssArray.getSpreadsheetFile().getName());
                }
                if (showSourceFile) {
                    for (Observations pObs: sfStats.getParserObservations().getObservationList()) {
                        if (newSourceFile) {
                            row.createCell(1).setCellValue(sfStats.getSourceFile().getName());
                            row.createCell(2).setCellValue(sfStats.getChecksum());
                            newSourceFile = false;
                            newSpreadsheet = false;
                        } else {
                            row = sheet.createRow(rowNum++);
                        }
                        row.createCell(3).setCellValue(pObs.getObservation());
                        row.createCell(4).setCellValue(pObs.startLine);
                        row.createCell(5).setCellValue(pObs.endLine);
                        row.createCell(6).setCellValue(pObs.subject);
                        setRowStyle(row, 3, 6, getStyle(pObs));
                    }
                }
            }
        }
    }

    public void createChecksumOverview(Workbook workbook, SpreadsheetStatistics stats, String sheetName, ArrayList<Integer> parseResults) {
        FileChecksums checksums = new FileChecksums();
        checksums.fillChecksums(stats, parseResults);

        String headers[] = {"Checksum", "Count", "Source"};
        Sheet sheet = workbook.createSheet(sheetName);

        sheet.setColumnWidth(0, checksumColSize);
        sheet.setColumnWidth(2, filesColSize+sourceColSize);

        fillHeader(headers, sheet.createRow(0));
        sheet.createFreezePane(0,1);

        int rowNum = 1;
        for (FileChecksums.ChecksumStruct cs: checksums.getChecksums()) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(cs.FileChecksum);
            row.createCell(1).setCellValue(cs.count);
            //row.createCell(2).setCellValue(cs.sourceFile.getAbsolutePath());
            row.createCell(2).setCellValue(String.format("%s/%s", cs.spreadsheetFile.getName(), cs.sourceFile.getName()));
        }
        Row row = sheet.createRow(rowNum++);
        row.createCell(0).setCellValue("Total files");
        row.createCell(1).setCellFormula(String.format("SUM(B1:B%d)", rowNum-1));
        setRowStyle(row, 0,1, totalsCellStyle);
        row = sheet.createRow(rowNum);
        row.createCell(0).setCellValue("Total unique checksums");
        row.createCell(1).setCellValue(checksums.getChecksums().size());
        setRowStyle(row, 0,1, totalsCellStyle);
    }

    private void fillHeader(String[] headers, Row headerRow) {
        // Create cells
        for(int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerCellStyle);
        }
    }

    /**
     * Writes the workbook with sheets to an Excel file.
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

    private void setRowStyle(Row row, int beginCell, int endCell, CellStyle cellStyle) {
        for (int i=beginCell; i <= endCell; i++) {
            row.getCell(i).setCellStyle(cellStyle);
        }
    }

     private CellStyle getStyle(Observations pObs) {
        CellStyle result = defaultCellStyle;
        if (pObs.getObservation().contains(Observations.VBA_HAS_PASSWORD)) {
            result = credentialsCellStyle;
        }
        if (pObs.getObservation().contains(Observations.VBA_HAS_USERID)) {
            result = credentialsCellStyle;
        }
        if (pObs.getObservation().contains(Observations.VBA_CREDENTIAL_ASSIGN)) {
             result = credentialsCellStyle;
         }
        if (pObs.getObservation().contains(Observations.VBA_HAS_USERID)) {
            result = credentialsCellStyle;
        }
        if (pObs.getObservation().contains(Observations.VBA_FILENAME_ASSIGN)) {
             result = fileCellStyle;
        }
        if (pObs.getObservation().contains(Observations.VBA_USES_EXTLIBS)) {
             result = extrefCellStyle;
        }
        return result;
    }
}
