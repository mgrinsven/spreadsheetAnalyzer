package spreadsheetAnalyzer;

public class SpreadsheetAnalyzer {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		if (args.length != 2) {
			System.out.println("This java program needs two paramaters");
			System.out.println("1. directory with excel sheets to be scanned");
			System.out.println("2. directory where extracted VBA macros will be saved\n");
			System.out.println("So java -jar checkvba.jar <InputDir> <OutputDir>");
			return;
		}
		
		RecursiveFileAnalyzer rfa = new RecursiveFileAnalyzer();
		rfa.Analyze(args[0], args[1]);
	}

}
