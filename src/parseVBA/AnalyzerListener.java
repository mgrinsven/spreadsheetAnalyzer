package parseVBA;

import Statistics.Observations;
import Statistics.ParserObservations;

/**
 * AnalyzerListener extends vbaBaseListener and implements the aspects that we are interested in.
 * For now we are interested to see if a VBA file is not an empty file or a template file.
 * We are interested in only the VBA files that contain actual written code
 */
public class AnalyzerListener extends vbaBaseListener {
    private vbaParser parser;
    ParserObservations observations;

    public AnalyzerListener(vbaParser parser, ParserObservations obs) {
        this.observations = obs;
        this.parser=parser;
    }
     @Override public void enterSubStmt(vbaParser.SubStmtContext ctx) {
        System.out.println("Enter Sub Statement");
        observations.addObservation(Observations.VBA_HAS_SUBSTMT, ctx.getStart().getLine(), ctx.getStop().getLine(), ctx.ambiguousIdentifier().getText());
    }

    @Override public void enterComment(vbaParser.CommentContext ctx) {
        if (ctx.getText().toLowerCase().contains("macro")) {
            System.out.println("This file contains macro inside a comment");
            observations.addObservation(Observations.VBA_HAS_MACROS, ctx.getStart().getLine(), ctx.getStop().getLine(), "");
        }
    }
    @Override public void enterFunctionStmt(vbaParser.FunctionStmtContext ctx) {
        System.out.println("Enter Function Statement");
        observations.addObservation(Observations.VBA_HAS_FUNCSTMT, ctx.getStart().getLine(), ctx.getStop().getLine(), ctx.ambiguousIdentifier().getText());
    }

    @Override public void enterArgList(vbaParser.ArgListContext ctx) {
        System.out.printf("ArgList [ ");
        for (int i=0; i < ctx.arg().size(); i++) {
            System.out.printf("%s, ", ctx.arg(i).getText());
            if (ctx.arg(i).getText().toLowerCase().contains("passw")) {
                System.out.printf("PASSWORD FOUND");
                System.out.printf("\nLine of occurence: %d-%d", ctx.getStart().getLine(), ctx.getStop().getLine());
                observations.addObservation(Observations.VBA_HAS_PASSWORD, ctx.getStart().getLine(), ctx.getStop().getLine(), ctx.arg(i).getText());
            }
            if (ctx.arg(i).getText().toLowerCase().contains("user")) {
                System.out.printf("USER FOUND");
                observations.addObservation(Observations.VBA_HAS_USERID, ctx.getStart().getLine(), ctx.getStop().getLine(), ctx.arg(i).getText());
            }
        }
        System.out.printf("]\n");
    }

    @Override public void enterDeclareStmt(vbaParser.DeclareStmtContext ctx) {
        System.out.print("EnterDeclareStmt: [ ");
        String declareString = "[";
        for (int i=0; i < ctx.STRINGLITERAL().size(); i++) {
            System.out.printf("%s : ", ctx.STRINGLITERAL(i).getText());
            declareString = declareString + ctx.STRINGLITERAL(i).getText() + ", ";
        }
        System.out.print("]\n");
        observations.addObservation(Observations.VBA_USES_EXTLIBS, ctx.getStart().getLine(), ctx.getStop().getLine(), declareString);
    }

/*    @Override public void enterAmbiguousKeyword(vbaParser.AmbiguousKeywordContext ctx) {
        System.out.printf("enterAmbiguousKeyword: [ %s ]\n", ctx.getText());
    }
 */
}
