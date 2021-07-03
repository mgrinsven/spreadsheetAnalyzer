package parseVBA;

import java.io.File;

public class ParseVBA {

	public static void main(String[] args) {
		String strFile="/home/menno/git_repos/sony/output/03_FY15_CC590020500020-Information Technology June 14.xlsm/Format_Sheet_for_Pixel_09052007.vba";
		//File file = new File(args[0]);
		File file = new File(strFile);
	}
	public boolean containsCode(File file) {
		ParserFacade facade = new ParserFacade();
		try {
			facade.parse(file);
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		if (facade.hasSubRoutines()) {
			return true;
		}
		return false;
	}
}
