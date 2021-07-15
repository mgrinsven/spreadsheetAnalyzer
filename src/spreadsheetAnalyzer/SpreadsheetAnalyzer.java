package spreadsheetAnalyzer;

public class SpreadsheetAnalyzer {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		if (args.length != 3) {
			System.out.println("This java program needs two paramaters");
			System.out.println("1. directory with excel sheets to be scanned");
			System.out.println("2. directory where extracted VBA macros will be saved");
			System.out.println("3. directpry where parsed VBA macros will be saved\n");
			System.out.println("So java -jar spreadsheetAnalyzer.jar <InputDir> <OutputDir> <parsedVbaDir");
			return;
		}
		
		RecursiveFileAnalyzer rfa = new RecursiveFileAnalyzer();
		rfa.Analyze(args[0], args[1], args[2]);
	}

}
