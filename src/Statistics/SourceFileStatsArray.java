package Statistics;

import extractPQ.PowerQueryExtractor;

import java.io.File;
import java.util.ArrayList;

public class SourceFileStatsArray {
    private SpreadsheetFileTotals totals = new SpreadsheetFileTotals();
    boolean totalsCalculated = false;
    File spreadsheetFile;
    ArrayList<SourceFileStats> sfStats;

    // PQ related information
    int containsPowerQuery = PowerQueryExtractor.PQ_OK;
    String dataMashUpURI = "";

    boolean containsError = false;

    public SourceFileStatsArray(File spreadsheet) {
        spreadsheetFile = spreadsheet;
        sfStats = new ArrayList<SourceFileStats>();
    }

    public ArrayList<SourceFileStats> getSourceFileStats() {
        return sfStats;
    }

    public File getSpreadsheetFile() {
        return spreadsheetFile;
    }

    public SourceFileStats addSourceFile(File sf) {
        SourceFileStats result = new SourceFileStats(sf);
        sfStats.add(result);
        return result;
    }

    public int getSourceFileCount() {
        return sfStats.size();
    }

    protected void calculateTotals() {
        for (SourceFileStats sfs:sfStats) {
            ArrayList<Integer> results = sfs.getSourceFileStats();
            for (int result:results) {
                switch (result) {
                    case SourceFileStats.IS_EMPTY_FILE:
                        totals.emptyFileCount++;
                        break;
                    case SourceFileStats.HAS_SUBS_AND_FUNC:
                        totals.containsCodeFileCount++;
                        break;
                    case SourceFileStats.HAS_MACROS:
                        totals.containsMacroFileCount++;
                        break;
                    case SourceFileStats.HAS_CREDENTIAL_REFS:
                        totals.containsCredentialsFileCount++;
                        break;
                    case SourceFileStats.HAS_EXT_REFS:
                        totals.containsExtRefsFileCount++;
                        break;
                }
            }
        }
        totals.totalFiles = sfStats.size();
        totalsCalculated = true;
    }

    public SpreadsheetFileTotals getTotals() {
        if (!totalsCalculated) {
            calculateTotals();
        }
        return totals;
    }

    public void setDataMashupURI(String mashUpURI) {
        dataMashUpURI = mashUpURI;
    }
    public String getDataMashupURI() {
        return dataMashUpURI;
    }

    public void setPowerQuery(int containsPQ) {
        containsPowerQuery = containsPQ;
    }
    public boolean hasPowerQuery() {
        return containsPowerQuery == PowerQueryExtractor.PQ_FOUND;
    }

    public boolean hasOtherDMU() {
        if (containsPowerQuery == PowerQueryExtractor.PQ_OTHER_DMU) {
            if (dataMashUpURI.length() > 1) {
                return true;
            }
        }
        return false;
    }

    public void setError(boolean setError) {
        containsError = setError;
    }
    public boolean hasError() {
        return containsError;
    }

    @Override
    public String toString() {
        return "SourceFileStatsArray{" +
                "spreadsheetFile=" + spreadsheetFile +
                ",\n\t sfStats=" + sfStats.toString() +
                '}';
    }
}
