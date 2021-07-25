package spreadsheetAnalyzer;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;

import Statistics.SourceFileStatsArray;
import Statistics.SpreadsheetStatistics;
import Statistics.SpreadsheetTotals;
import extractPQ.PowerQueryExtractor;
//import Statistics.Observations;
import parseVBA.ParseVBA;

public class RecursiveFileAnalyzer {
	private static final boolean DEBUG = true;
	private ArrayList<File> fList;
	private AnalyzerStatistics stats;
	private SourceFileStatsArray spreadsheetStats;
	private SpreadsheetStatistics spreadsheetStatsArray;

	public RecursiveFileAnalyzer() {
		// TODO Auto-generated constructor stub
		stats = new AnalyzerStatistics();
		spreadsheetStatsArray = new SpreadsheetStatistics();
	}

	public void Analyze(String sourceDir, String targetDir, String parsedVbaDir) {
		VbaExtractor vbaExtractor = new VbaExtractor();
		PowerQueryExtractor pqExtractor = new PowerQueryExtractor();

		fList = new ArrayList<File>();
		recursiveFileList(sourceDir);

		for (File file:fList) {
			spreadsheetStats = new SourceFileStatsArray(file);
			spreadsheetStatsArray.addSpreadsheet(spreadsheetStats);
			int result = vbaExtractor.ExtractVba(file, targetDir);
			switch (result) {
			case VbaExtractor.ERROR:
				stats.totalErrors++;
				break;
			case VbaExtractor.VBA_EXTRACTED:
				stats.totalVBAFiles++;
				AnalyzeVBAFiles(file, targetDir, parsedVbaDir);
			case VbaExtractor.OK:
				System.out.println("Now we can start looking at PQ formulas");
				if (pqExtractor.ExtractPQ(file, targetDir)) {
					spreadsheetStats.setPowerQuery(true);
					stats.totalPQFormulas++;
				}
				break;
			}
			stats.totalFileCount++;
		}
		System.out.println("===================================");
		System.out.printf("Total files scanned: %d \nFiles containing VBA: %d\nTotal files containing PQ: %d\nTotal errors: %d\n", stats.totalFileCount, stats.totalVBAFiles, stats.totalPQFormulas, stats.totalErrors);
		System.out.printf("Total VBA files: %d\n", stats.totalExtractedVBAFiles);
		System.out.printf("Total VBA files containing code: %d\n", stats.totalVBACodeFiles);
		System.out.printf("Total VBA files containing recorded macro's: %d\n", stats.totalVBAMacroFiles);
		System.out.printf("Total empty VBA files: %d\n", stats.totalEmptyVBA);
		System.out.printf("Total VBA files with credentials: %d\n", stats.totalVBACredentials);
		//System.out.printf("\n\n\n%s", spreadsheetStats.toString());

		System.out.println("===================================");
		SpreadsheetTotals totals = spreadsheetStatsArray.getTotals();
		System.out.printf("Total files scanned: %d \nFiles containing VBA: %d\nTotal files containing PQ: %d\nTotal errors: %d\n", spreadsheetStatsArray.getTotalFiles(), totals.totalVBASpreadsheets, totals.totalPowerQuery, stats.totalErrors);
		System.out.printf("Total VBA files: %d\n", totals.totalFiles);
		System.out.printf("Total VBA files containing code: %d\n", totals.containsCodeFileCount);
		System.out.printf("Total VBA files containing recorded macro's: %d\n", totals.containsMacroFileCount);
		System.out.printf("Total empty VBA files: %d\n", totals.emptyFileCount);
		System.out.printf("Total VBA files with credentials: %d\n", totals.containsCredentialsFileCount);
		System.out.printf("\n\n\n%s", spreadsheetStats.toString());

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

	private boolean AnalyzeVBAFiles(File sourceFile, String targetDir, String parsedVbaDir) {
		boolean result = false;
		ParseVBA vbaParser = new ParseVBA();
		// Construct the directory the VBA extractor put the VBA code
		String sourceDir = targetDir+"/"+sourceFile.getName();
		System.out.printf("Scanning VBA files in %s\n", sourceDir);
		//spreadsheetStats = new SourceFileStatsArray(sourceFile);

		// Create a list of VBA souce files inside that directory
		File[] files = new File(sourceDir).listFiles();

		// Create output directories
		try {
			Files.createDirectories(Paths.get(parsedVbaDir));
			Files.createDirectories(Paths.get(parsedVbaDir + "/empty"));
			Files.createDirectories(Paths.get(parsedVbaDir + "/macro"));
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		// Loop through all the files and let the ANTLR parser decide if the file
		// contains VBA code we want to analyze further
		for (File file:files) {
			stats.totalExtractedVBAFiles++;

			if (vbaParser.hasObservations(file, spreadsheetStats.addSourceFile(file))) {
				if (vbaParser.getObservations().countMacros() > 0) {
					if (vbaParser.getObservations().countCodeBlocks() > 0) {
						stats.totalVBACodeFiles++;
						// The file contains code, so we want to put it somewhere separate
						//TODO copy file to directory and rename if same name exist
						try {
							// Create a filename consisting of the xls name and the name of the vba source file
							String targetFileName = String.format("%s-%s", sourceFile.getName(), file.getName());
							Files.copy(Paths.get(file.getAbsolutePath()), Paths.get(parsedVbaDir + "/" + targetFileName), StandardCopyOption.REPLACE_EXISTING);
						} catch (Exception e) {
							System.out.println(e.getMessage());
						}
						if (vbaParser.getObservations().countCredentials() > 0) {
							stats.totalVBACredentials++;
						}
						//There is at least one vba file that contains a Sub statement
						result = true;
					}
				} else {
					try {
						String targetFileName = String.format("%s-%s", sourceFile.getName(), file.getName());
						Files.copy(Paths.get(file.getAbsolutePath()), Paths.get(parsedVbaDir + "/macro/" + targetFileName), StandardCopyOption.REPLACE_EXISTING);
						stats.totalVBAMacroFiles++;
					} catch (Exception e) {
						System.out.println(e.getMessage());
					}
				}
			} else {
				// VBA file has no observations, meaning there is no code...
				try {
					String targetFileName = String.format("%s-%s", sourceFile.getName(), file.getName());
					Files.copy(Paths.get(file.getAbsolutePath()), Paths.get(parsedVbaDir + "/empty/" +targetFileName), StandardCopyOption.REPLACE_EXISTING);
					stats.totalEmptyVBA++;
				} catch (Exception e) {
					System.out.println(e.getMessage());
				}
			}
		}
		return result;
	}
}