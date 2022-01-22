package Statistics;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.security.DigestInputStream;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.ArrayList;

public class SourceFileStats {
    public static final int IS_EMPTY_FILE = 0;
    public static final int HAS_MACROS = 1;
    public static final int HAS_CREDENTIAL_REFS = 2;
    public static final int HAS_SUBS_AND_FUNC = 3;
    public static final int HAS_EXT_REFS = 4;
    public static final int HAS_DB_REFS = 5;

    File sourceFile;
    String checkum;
    SourceFileTotals totals;
    ParserObservations parserObservations;
    ArrayList<Integer> parseResults = new ArrayList<Integer>();

    public SourceFileStats(File sourceFile) {
        this.sourceFile = sourceFile;
        parserObservations = new ParserObservations();
        totals = new SourceFileTotals();
    }

    public File getSourceFile() {
        return sourceFile;
    }

    public String getChecksum() {
        return checkum;
    }

    public ParserObservations getObservations() {
        return parserObservations;
    }

    public ArrayList<Integer> getSourceFileStats() {
        if (parseResults.isEmpty()) {
            try {
                checkum = calcChecksum(sourceFile);
                //System.out.printf("SHA1: %s\n", sha1);
            } catch (IOException e) {
                e.printStackTrace();
            }

            totals.containsMacroFileCount = parserObservations.countMacros();
            totals.containsCodeFileCount = parserObservations.countCodeBlocks();
            totals.containsCredentialsFileCount = parserObservations.countCredentials();
            totals.containsExtRefsFileCount = parserObservations.countExternalLibRefs();
            totals.containsDatabaseCount = parserObservations.countDatabase();

            if (parserObservations.hasObservations()) {
                if (totals.containsMacroFileCount > 0) {
                    parseResults.add(HAS_MACROS);
                }
                if (totals.containsCredentialsFileCount > 0) {
                    parseResults.add(HAS_CREDENTIAL_REFS);
                }
                if (totals.containsCodeFileCount > 0) {
                    parseResults.add(HAS_SUBS_AND_FUNC);
                }
                if (totals.containsExtRefsFileCount > 0) {
                    parseResults.add(HAS_EXT_REFS);
                }
                if (totals.containsDatabaseCount > 0) {
                    parseResults.add(HAS_DB_REFS);
                }
            } else {
                parseResults.add(IS_EMPTY_FILE);
            }
        }
        return parseResults;
    }

    public SourceFileTotals getTotals() {
        return totals;
    }

    public ParserObservations getParserObservations() {
        return parserObservations;
    }

    private String calcChecksum(File file) throws IOException {
        MessageDigest md = null; //SHA, MD2, MD5, SHA-256, SHA-384...
        try {
            md = MessageDigest.getInstance("SHA-1");
        } catch (NoSuchAlgorithmException e) {
            e.printStackTrace();
        }
        // file hashing with DigestInputStream
        try (DigestInputStream dis = new DigestInputStream(new FileInputStream(file), md)) {
            while (dis.read() != -1) ; //empty loop to clear the data
            md = dis.getMessageDigest();
        }

        // bytes to hex
        StringBuilder result = new StringBuilder();
        for (byte b : md.digest()) {
            result.append(String.format("%02x", b));
        }
        return result.toString();
    }

    @Override
    public String toString() {
        return "\n\tSourceFileStats{" +
                "sourceFile=" + sourceFile +
                ", Observations=" + parserObservations +
                '}';
    }
}
