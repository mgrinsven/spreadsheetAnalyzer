package extractPQ;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

public class ExtractQDEFF {
	private final boolean DEBUG=true;
	
	private static final String QDEFF_FILE = "customXml/item1.xml";
	private static final int BUFFER_SIZE = 4096;
	public static final int ERROR = -1;
	public static final int OK = 0;
	
	public int extractPQFile(File zipFile, File destFile) {
		int result = ERROR;
		try {
			ZipFile zf = new ZipFile(zipFile);
			
			ZipEntry ze = zf.getEntry(QDEFF_FILE);
			if (ze != null) {
				InputStream zis = zf.getInputStream(ze);
				if (DEBUG) System.out.printf("Trying to extract %s\n", zipFile.getName());
				File path = new File(destFile.getParent());
				if (!path.exists()) {
					if (DEBUG) System.out.printf("%s does not exist, creating\n", path.getPath());
					path.mkdir();
				}

				FileOutputStream fos = new FileOutputStream(destFile);
		
		        byte[] bytesIn = new byte[BUFFER_SIZE];
		        int read = 0;
		        while ((read = zis.read(bytesIn)) != -1) {
		            fos.write(bytesIn, 0, read);
		        }
		        fos.close();
		        result = OK;
			}
			zf.close();
		} catch (java.io.IOException ioe) {
			System.out.println(ioe.getMessage());
			result = ERROR;
		}
		return result;
	}
}
