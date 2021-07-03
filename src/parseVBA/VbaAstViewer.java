package parseVBA;

import java.io.File;
import java.io.IOException;
import java.util.Arrays;

import javax.swing.JFrame;
import javax.swing.JPanel;

import org.antlr.v4.gui.TreeViewer;
import org.antlr.v4.runtime.ANTLRFileStream;
import org.antlr.v4.runtime.CommonTokenStream;
import org.antlr.v4.runtime.tree.ParseTree;

public class VbaAstViewer {
	public void showTree(vbaParser parser, ParseTree tree) {
        //show AST in GUI
		buildTree(parser, tree);
	}
	public void showTree(File file) {
		try {
			ANTLRFileStream fileStream = new ANTLRFileStream(file.getAbsolutePath());
	    	vbaLexer lexer = new vbaLexer(fileStream);
	        CommonTokenStream tokens = new CommonTokenStream(lexer);
	        vbaParser parser = new vbaParser(tokens);
	        ParseTree tree = parser.attributeStmt();//ifBlockStmt();
	        buildTree(parser, tree);
		} catch (IOException e){
			System.out.println(e.getMessage());
		}
		
	}
	
	private void buildTree(vbaParser parser, ParseTree tree) {
        JFrame frame = new JFrame("Antlr AST");
        JPanel panel = new JPanel();
        TreeViewer viewer = new TreeViewer(Arrays.asList( parser.getRuleNames()),tree);
        viewer.setScale(1); // Scale a little
        panel.add(viewer);
        frame.add(panel);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.pack();
        frame.setVisible(true);
	}
}
