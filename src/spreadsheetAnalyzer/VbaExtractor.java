package spreadsheetAnalyzer;

import java.io.File;
import org.apache.poi.poifs.macros.VBAMacroReader;
import org.apache.poi.poifs.macros.VBAMacroExtractor;
import org.apache.poi.poifs.filesystem.NotOLE2FileException;

public class VbaExtractor {
	private VBAMacroExtractor vme = null;
	
	final static int ERROR = -1;
	final static int OK = 0;
	final static int NOT_VALID = 1;
	final static int VBA_EXTRACTED = 2;
	
	public VbaExtractor() {
		vme = new VBAMacroExtractor();
	}

	int ExtractVba(File sourceFile, String targetDir) {
		int result = OK;
		try {
			VBAMacroReader vmr = new VBAMacroReader(sourceFile);
			if (vmr.readMacros().size() > 0) {
				System.out.println(sourceFile.getAbsolutePath()+" contains macros");
				vme.extract(sourceFile, new File(targetDir+"/"+sourceFile.getName()));
				result = VBA_EXTRACTED;
			}
			vmr.close();
		} catch(IllegalArgumentException e) {
			System.out.printf("IllegalArgumentException: %s\n", e.getMessage());
			result = OK;
		} catch(NotOLE2FileException e) {
			System.out.println(sourceFile.getName()+" is not a valid office file");
			result = NOT_VALID;
		} catch(Exception e) {
			System.out.printf("Exception: %s\n", e.getMessage());
			result = ERROR;
		}
		return result;
	}
}
