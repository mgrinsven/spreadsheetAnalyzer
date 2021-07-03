package spreadsheetAnalyzer;

import java.io.File;
import java.util.ArrayList;

import extractPQ.ExtractQDEFF;
import extractPQ.ParseDataMashUp;
import extractPQ.PowerQueryExtractor;

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

		//ArrayList<File> fList = recursiveFileList(sourceDir);
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
	//private ArrayList<File> recursiveFileList(String dir)
	private void recursiveFileList(String dir) {
		//ArrayList<File> flist = new ArrayList<File>();
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
	//	return flist;
	}
}