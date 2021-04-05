package spreadsheetAnalyzer;

import java.io.File;
import java.util.ArrayList;

import extractPQ.ExtractQDEFF;
import extractPQ.ParseDataMashUp;

public class RecursiveFileAnalyzer {
	private static final boolean DEBUG = true;

	public RecursiveFileAnalyzer() {
		// TODO Auto-generated constructor stub
	}

	public void Analyze(String sourceDir, String targetDir) {
		int totalFileCount = 0;
		int totalVBAFiles = 0;
		int totalPQFormulas = 0;
		int totalErrors = 0;
		
		VbaExtractor vbaExtractor = new VbaExtractor();
		ArrayList<File> fList = recursiveFileList(sourceDir);
		
		for (File file:fList) {
			int result = vbaExtractor.ExtractVba(file, targetDir);
			switch (result) {
			case VbaExtractor.ERROR:
				totalErrors++;
				break;
			case VbaExtractor.VBA_EXTRACTED:
				totalVBAFiles++;
			case VbaExtractor.OK:
				System.out.println("Now we can start looking at PQ formulas");
				if (ExtractPQ(file, targetDir)) {
					totalPQFormulas++;
				}
				break;
			}
			totalFileCount++;
		}
		System.out.println("===================================");
		System.out.printf("Total files scanned: %d \nFiles containing VBA: %d\nTotal files containing PQ: %d\nTotal errors: %d\n", totalFileCount, totalVBAFiles, totalPQFormulas,totalErrors);

	}

	private boolean ExtractPQ(File file, String targetDir) {
		boolean result = false;
		ExtractQDEFF qdeff = new ExtractQDEFF();
		String targetPath = String.format("%s/%s", targetDir, file.getName());
		File destFile = new File(targetPath+"/item1.xml");
		int qdeff_result = qdeff.extractPQFile(file, destFile);
		if (qdeff_result == ExtractQDEFF.OK) {
			ParseDataMashUp pdmu = new ParseDataMashUp();
			if (pdmu.getXPathValue(destFile, targetPath)) {
				result = pdmu.extractFormula(targetPath);
				result = pdmu.extractPermissions(targetPath);
				result = pdmu.extractPermissionBinding(targetPath);
				result = pdmu.extractMetaData(targetPath);
			}
		}
		return result;
	}
	
	
	/**
	 * recursiveFileList recursively creates a list of files
	 * @param dir Start directory from where the files are added
	 */
	private ArrayList<File> recursiveFileList(String dir) {
		ArrayList<File> flist = new ArrayList<File>();
		File[] files = new File(dir).listFiles();
		for (File file:files) {
			if (file.isDirectory()) {
				System.out.println("Subdir: "+file.getAbsolutePath());
				recursiveFileList(file.getAbsolutePath());
			} else {
				// Perhaps add a file filter also (only xls / xlsx / xlsm)
				flist.add(file);
			}
		}
		return flist;
	}
}
