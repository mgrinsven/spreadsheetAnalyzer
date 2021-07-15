package parseVBA;

import java.io.File;

public class ParseVBA {
	ParserObservations observations;

	public static void main(String[] args) {
		ParserObservations pObs = new ParserObservations();
		String strFile="/home/menno/git_repos/sony/output/03_FY15_CC590020500020-Information Technology June 14.xlsm/Format_Sheet_for_Pixel_09052007.vba";
		//String strFile="/home/menno/git_repos/sony/output/03_FY15_CC590020500020-Information Technology June 14.xlsm/Cover.vba";
		//File file = new File(args[0]);
		File file = new File(strFile);
		ParserFacade facade = new ParserFacade(pObs);
		try {
			facade.parse(file);
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		if (pObs.hasObservations()) {
			System.out.println("File has VBA code that contains at least one Sub statement");
		}
	}

	public boolean hasObservations(File file) {
		observations = new ParserObservations();
		ParserFacade facade = new ParserFacade(observations);
		try {
			facade.parse(file);
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		if (observations.hasObservations()) {
			for (Observations obs:observations.getObservationList()) {
				System.out.printf("Observation %d : %d\n", obs.Observation, obs.Count);
			}
			return true;
		}
		return false;
	}

	public ParserObservations getObservations() {
		return observations;
	}
}
