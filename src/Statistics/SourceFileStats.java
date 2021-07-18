package Statistics;

import java.io.File;

public class SourceFileStats {
    File sourceFile;
    ParserObservations Observations;

    public SourceFileStats(File sourceFile) {
        this.sourceFile = sourceFile;
        Observations = new ParserObservations();
    }

    public File getSourceFile() {
        return sourceFile;
    }

    public ParserObservations getObservations() {
        return Observations;
    }

    @Override
    public String toString() {
        return "\n\tSourceFileStats{" +
                "sourceFile=" + sourceFile +
                ", Observations=" + Observations +
                '}';
    }
}
