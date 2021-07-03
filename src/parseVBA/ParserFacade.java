package parseVBA;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;

import org.antlr.v4.runtime.tree.ParseTree;
import org.antlr.v4.runtime.ANTLRFileStream;
import org.antlr.v4.runtime.CommonTokenStream;
import org.antlr.v4.runtime.*;
import org.antlr.v4.runtime.tree.ParseTreeWalker;

public class ParserFacade {
	public ParserFacade() {}
	private AnalyzerListener extractor;

    public void parse(File file) throws IOException {
    	ANTLRFileStream fileStream = new ANTLRFileStream(file.getAbsolutePath());

    	vbaLexer lexer = new vbaLexer(fileStream);
        CommonTokenStream tokens = new CommonTokenStream(lexer);
        vbaParser parser = new vbaParser(tokens);

        //AnalyzerListener extractor = new AnalyzerListener(parser);
        extractor = new AnalyzerListener(parser);

        ParseTree tree = parser.module();
        ParseTreeWalker.DEFAULT.walk(extractor, tree);
    }

    AnalyzerListener getExtractor() {
        return extractor;
    }
    boolean hasSubRoutines() {
        return extractor.getSubStmt();
    }

 /*   private void explore(RuleContext ctx, int indentation, OutputStream os) {
    	String line = "";
    	String indent = "";
    	String identifier = "";
        String ruleName = vbaParser.ruleNames[ctx.getRuleIndex()];
        for (int i=0;i<indentation;i++) {
            indent += "  ";
        }

        if (ruleName.equals("ambiguousIdentifier") || ruleName.equals("ambiguousKeyword") || ruleName.equals("literal")
        		|| ruleName.equals("comment" ) || ruleName.equals("baseType" )) {
        	identifier = String.format(" : %s", ctx.getText());
        }
        
        line = String.format("%s%s%s\n", indent, ruleName, identifier);
        System.out.print(line);
        try {
        	os.write(line.getBytes());
        } catch (Exception e) {
        	e.getMessage();
        }
        for (int i=0;i<ctx.getChildCount();i++) {
            ParseTree element = ctx.getChild(i);
            if (element instanceof RuleContext) {
                explore((RuleContext)element, indentation + 1, os);
            }
        }
    }
  */
}
