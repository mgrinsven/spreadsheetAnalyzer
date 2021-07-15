package parseVBA;

/**
 * AnalyzerListener extends vbaBaseListener and implements the aspects that we are interested in.
 * For now we are interested to see if a VBA file is not an empty file or a template file.
 * We are interested in only the VBA files that contain actual written code
 */
public class AnalyzerListener extends vbaBaseListener {
    private vbaParser parser;
//    private int nrSubStmts = 0;
//    private int nrMacroCmts = 0;
    ParserObservations observations;

    public AnalyzerListener(vbaParser parser, ParserObservations obs) {
        this.observations = obs;
        this.parser=parser;
    }
     @Override public void enterSubStmt(vbaParser.SubStmtContext ctx) {
        System.out.println("Enter SubStmt");
        observations.updateObservation(Observations.VBA_HAS_SUBSTMT);
//        nrSubStmts++;
    }

    @Override public void enterComment(vbaParser.CommentContext ctx) {
        if (ctx.getText().toLowerCase().contains("macro")) {
            System.out.println("This file contains macro inside a comment");
            observations.updateObservation(Observations.VBA_HAS_MACROS);
//            nrMacroCmts++;
        }
    }
/*    public int getNrSubStmts() {
        return nrSubStmts;
    }

    public int getNrMacroCmts() {
        return nrMacroCmts;
    }

 */
}
