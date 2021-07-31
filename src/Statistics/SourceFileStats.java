package Statistics;

import java.io.File;
import java.util.ArrayList;

public class SourceFileStats {
    public static final int IS_EMPTY_FILE = 0;
    public static final int HAS_MACROS = 1;
    public static final int HAS_CREDENTIAL_REFS = 2;
    public static final int HAS_SUBS_AND_FUNC = 3;
    public static final int HAS_EXT_REFS = 4;

    File sourceFile;
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

    public ParserObservations getObservations() {
        return parserObservations;
    }

    public ArrayList<Integer> getSourceFileStats() {
        if (parseResults.isEmpty()) {
            totals.containsMacroFileCount = parserObservations.countMacros();
            totals.containsCodeFileCount = parserObservations.countCodeBlocks();
            totals.containsCredentialsFileCount = parserObservations.countCredentials();
            totals.containsExtRefsFileCount = parserObservations.countExternalLibRefs();

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

    @Override
    public String toString() {
        return "\n\tSourceFileStats{" +
                "sourceFile=" + sourceFile +
                ", Observations=" + parserObservations +
                '}';
    }
}
