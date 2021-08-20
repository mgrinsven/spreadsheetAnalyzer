package parseVBA;

import java.io.File;
import java.io.IOException;

import Statistics.ParserObservations;
import org.antlr.v4.runtime.tree.ParseTree;
import org.antlr.v4.runtime.ANTLRFileStream;
import org.antlr.v4.runtime.CommonTokenStream;
import org.antlr.v4.runtime.tree.ParseTreeWalker;

public class ParserFacade {
    private AnalyzerListener extractor;
    private ParserObservations observations;

	public ParserFacade(ParserObservations obs) {
	    this.observations = obs;
    }

    public void parse(File file) throws IOException {
    	ANTLRFileStream fileStream = new ANTLRFileStream(file.getAbsolutePath());

    	vbaLexer lexer = new vbaLexer(fileStream);
        CommonTokenStream tokens = new CommonTokenStream(lexer);
        vbaParser parser = new vbaParser(tokens);

        //AnalyzerListener extractor = new AnalyzerListener(parser);
        extractor = new AnalyzerListener(parser, observations, file.getName());

        ParseTree tree = parser.module();
        ParseTreeWalker.DEFAULT.walk(extractor, tree);
    }

    AnalyzerListener getExtractor() {
        return extractor;
    }

    public ParserObservations getObservations() {
	    return observations;
    }
}
