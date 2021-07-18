package Statistics;

import java.io.File;
import java.util.ArrayList;

public class SourceFileStatsArray {
    File spreadsheetFile;
    ArrayList<SourceFileStats> sfStats;

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

    @Override
    public String toString() {
        return "SourceFileStatsArray{" +
                "spreadsheetFile=" + spreadsheetFile +
                ",\n\t sfStats=" + sfStats.toString() +
                '}';
    }
}
