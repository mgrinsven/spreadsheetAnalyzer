package parseVBA;

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
