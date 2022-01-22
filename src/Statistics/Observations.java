package Statistics;

public class Observations {
    public static final String VBA_HAS_NOOBSERVATIONS = "VBA_HAS_NOOBSERVATIONS";
    public static final String VBA_HAS_SUBSTMT = "VBA_HAS_SUBSTMT";
    public static final String VBA_HAS_MACROS = "VBA_HAS_MACROS";
    public static final String VBA_HAS_FUNCSTMT = "VBA_HAS_FUNCSTMT";
    public static final String VBA_HAS_USERID = "VBA_HAS_USERID";
    public static final String VBA_HAS_PASSWORD = "VBA_HAS_PASSWORD";
    public static final String VBA_USES_EXTLIBS = "VBA_USES_EXTLIBS";
    public static final String VBA_CREDENTIAL_ASSIGN = "VBA_CREDENTIAL_ASSIGN";
    public static final String VBA_FILENAME_ASSIGN = "VBA_FILENAME_ASSIGN";
    public static final String VBA_DB_ASSIGN = "VBA_DB_ASSIGN";

    String Observation;
    int startLine;
    int endLine;
    String subject;

    public Observations(String observation, int startLine, int endLine, String subject) {
        Observation = observation;
        this.startLine = startLine;
        this.endLine = endLine;
        this.subject = subject;
    }

    public String getObservation() {
        return Observation;
    }

    public int getStartLine() {
        return startLine;
    }

    public int getEndLine() {
        return endLine;
    }

    @Override
    public String toString() {
        return "Observations{" +
                "Observation='" + Observation + '\'' +
                ", startLine=" + startLine +
                ", endLine=" + endLine +
                ", subject='" + subject + '\'' +
                '}';
    }
}
