package Statistics;

import java.util.ArrayList;

public class SpreadsheetStatistics {
    protected ArrayList<SourceFileStatsArray> spreadsheetStats;
    protected SpreadsheetTotals totals;
    protected boolean totalsCalculated = false;

    public SpreadsheetStatistics() {
        spreadsheetStats = new ArrayList<SourceFileStatsArray>();
        totals = new SpreadsheetTotals();
    }

    public void addSpreadsheet(SourceFileStatsArray sfsArray) {
        spreadsheetStats.add(sfsArray);
    }

    private void calculateTotals() {
        for (SourceFileStatsArray sfsArray:spreadsheetStats) {
            totals.emptyFileCount += sfsArray.getTotals().emptyFileCount;
            totals.containsCodeFileCount += sfsArray.getTotals().containsCodeFileCount;
            totals.containsCredentialsFileCount += sfsArray.getTotals().containsCredentialsFileCount;
            totals.containsMacroFileCount += sfsArray.getTotals().containsMacroFileCount;
            totals.containsExtRefsFileCount += sfsArray.getTotals().containsExtRefsFileCount;
            totals.totalFiles += sfsArray.getTotals().totalFiles;
            if (sfsArray.hasPowerQuery()) {
                totals.totalPowerQuery++;
            }
            if (sfsArray.hasOtherDMU()) {
                totals.totalOtherDMU++;
            }
            if (sfsArray.getTotals().totalFiles > 0) {
                totals.totalVBASpreadsheets++;
            }
            if (sfsArray.hasError()) {
                totals.totalErrors++;
            }
        }
        totalsCalculated = true;
    }

    public SpreadsheetTotals getTotals() {
        if (!totalsCalculated) {
            calculateTotals();
        }
        return totals;
    }

    public int getTotalFiles() {
        return spreadsheetStats.size();
    }

    public ArrayList<SourceFileStatsArray> getSpreadsheetStats() {
        return spreadsheetStats;
    }
}
