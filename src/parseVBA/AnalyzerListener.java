package parseVBA;

/**
 * AnalyzerListener extends vbaBaseListener and implements the aspects that we are interested in.
 * For now we are interested to see if a VBA file is not an empty file or a template file.
 * We are interested in only the VBA files that contain actual written code
 */
public class AnalyzerListener extends vbaBaseListener {
    private vbaParser parser;
    private boolean hasSubStmt = false;

    public AnalyzerListener(vbaParser parser) {
        this.parser=parser;
    }
     @Override public void enterSubStmt(vbaParser.SubStmtContext ctx) {
        System.out.println("Enter SubStmt");
        hasSubStmt = true;
    }
    @Override public void exitSubStmt(vbaParser.SubStmtContext ctx) {
        System.out.println("Exit SubStmt");
    }

    public boolean getSubStmt() {
        return hasSubStmt;
    }
}
