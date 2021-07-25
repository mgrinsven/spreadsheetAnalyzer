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
    ParserObservations parserObservations;
    ArrayList<Integer> parseResults = new ArrayList<Integer>();

    public SourceFileStats(File sourceFile) {
        this.sourceFile = sourceFile;
        parserObservations = new ParserObservations();
    }

    public File getSourceFile() {
        return sourceFile;
    }

    public ParserObservations getObservations() {
        return parserObservations;
    }

    public ArrayList<Integer> getSourceFileStats() {
        if (parseResults.isEmpty()) {
            if (parserObservations.hasObservations()) {
                if (parserObservations.countMacros() > 0) {
                    parseResults.add(HAS_MACROS);
                }
                if (parserObservations.countCredentials() > 0) {
                    parseResults.add(HAS_CREDENTIAL_REFS);
                }
                if (parserObservations.countCodeBlocks() > 0) {
                    parseResults.add(HAS_SUBS_AND_FUNC);
                }
                if (parserObservations.countExternalLibRefs() > 0) {
                    parseResults.add(HAS_EXT_REFS);
                }
            } else {
                parseResults.add(IS_EMPTY_FILE);
            }
        }
        return parseResults;
    }

    @Override
    public String toString() {
        return "\n\tSourceFileStats{" +
                "sourceFile=" + sourceFile +
                ", Observations=" + parserObservations +
                '}';
    }
}
