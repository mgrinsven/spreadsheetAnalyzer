package spreadsheetAnalyzer;

import java.io.File;
import java.util.ArrayList;

import extractPQ.PowerQueryExtractor;
import parseVBA.ParseVBA;

public class RecursiveFileAnalyzer {
	private static final boolean DEBUG = true;
	private ArrayList<File> fList;

	public RecursiveFileAnalyzer() {
		// TODO Auto-generated constructor stub
	}

	public void Analyze(String sourceDir, String targetDir) {
		int totalFileCount = 0;
		int totalVBAFiles = 0;
		int totalPQFormulas = 0;
		int totalErrors = 0;
		
		VbaExtractor vbaExtractor = new VbaExtractor();
		PowerQueryExtractor pqExtractor = new PowerQueryExtractor();

		fList = new ArrayList<File>();
		recursiveFileList(sourceDir);

		for (File file:fList) {
			int result = vbaExtractor.ExtractVba(file, targetDir);
			switch (result) {
			case VbaExtractor.ERROR:
				totalErrors++;
				break;
			case VbaExtractor.VBA_EXTRACTED:
				totalVBAFiles++;
				AnalyzeVBAFiles(file, targetDir);
			case VbaExtractor.OK:
				System.out.println("Now we can start looking at PQ formulas");
				if (pqExtractor.ExtractPQ(file, targetDir)) {
					totalPQFormulas++;
				}
				break;
			}
			totalFileCount++;
		}
		System.out.println("===================================");
		System.out.printf("Total files scanned: %d \nFiles containing VBA: %d\nTotal files containing PQ: %d\nTotal errors: %d\n", totalFileCount, totalVBAFiles, totalPQFormulas,totalErrors);

	}

	/**
	 * recursiveFileList recursively creates a list of files
	 * @param dir Start directory from where the files are added
	 */
	private void recursiveFileList(String dir) {
		File[] files = new File(dir).listFiles();
		for (File file:files) {
			if (file.isDirectory()) {
				System.out.println("Subdir: "+file.getAbsolutePath());
				recursiveFileList(file.getAbsolutePath());
			} else {
				// Perhaps add a file filter also (only xls / xlsx / xlsm)
				fList.add(file);
				System.out.printf("File: %s, Total count %d\n", file, fList.size());
			}
		}
	}

	private void AnalyzeVBAFiles(File sourceFile, String targetDir) {
		ParseVBA vbaParser = new ParseVBA();
		// Construct the directory the VBA extractor put the VBA code
		String sourceDir = targetDir+"/"+sourceFile.getName();
		System.out.printf("Scanning VBA files in %s\n", sourceDir);

		// Create a list of VBA souce files inside that directory
		File[] files = new File(sourceDir).listFiles();
		// Loop through all the files and let the ANTLR parser decide if the file
		// contains VBA code we want to analyze further
		for (File file:files) {
			if (vbaParser.containsCode(file)) {
				// The file contains code, so we want to put it somewhere separate

				//TODO copy file to directory and rename if same name exist
			}
		}

	}
}