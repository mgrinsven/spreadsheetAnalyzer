package Statistics;

import java.io.File;
import java.util.ArrayList;

public class FileChecksums {
    public class ChecksumStruct {
        int count;
        String FileChecksum;
        File sourceFile;
        File spreadsheetFile;

        public ChecksumStruct(String fileChecksum, File spreadsheet, File source) {
            this.FileChecksum = fileChecksum;
            this.spreadsheetFile = spreadsheet;
            this.sourceFile = source;
            count = 1;
        }
        public int getCount() {
            return count;
        }

        public String getFileChecksum() {
            return FileChecksum;
        }

        public File getSourceFile() {
            return sourceFile;
        }

        public File getSpreadsheetFile() {
            return spreadsheetFile;
        }
    }

    protected ArrayList<ChecksumStruct> checksums = new ArrayList<ChecksumStruct>();

    public void addChecksum(String checksum, File spreadsheet, File source) {
        boolean match = false;
        for (ChecksumStruct cs: checksums) {
            if (cs.FileChecksum.equals(checksum)) {
                cs.count++;
                match = true;
            }
        }
        if (!match) {
            checksums.add(new ChecksumStruct(checksum, spreadsheet, source));
        }
    }

    public void fillChecksums(SpreadsheetStatistics stats, ArrayList<Integer> parseResults) {
        for (SourceFileStatsArray ssArray: stats.getSpreadsheetStats()) {
            for (SourceFileStats sfStats: ssArray.getSourceFileStats()) {
                for (Integer result: sfStats.getSourceFileStats()) {
                    for (Integer checkResult : parseResults) {
                        if (checkResult == result) {
                            addChecksum(sfStats.getChecksum(), ssArray.getSpreadsheetFile(), sfStats.getSourceFile());
                        }
                    }
                }
            }
        }
    }

    public ArrayList<ChecksumStruct> getChecksums() {
        return checksums;
    }
}
