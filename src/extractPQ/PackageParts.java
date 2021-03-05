package extractPQ;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.ByteBuffer;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

public class PackageParts {
	private static final String FORMULA_FILE = "Formulas/Section1.m"; 
	private ByteBuffer bbPackageParts = null;
	private static final int BUFFER_SIZE = 4096;
	
	public PackageParts(ByteBuffer packageparts) {
		bbPackageParts = packageparts;
	}
	
	public boolean extractFormula(File zipFile, File destFile) {
		boolean result = true;
		try {
			ZipFile zf = new ZipFile(zipFile);
			System.out.println("extractFormula");

			ZipEntry ze = zf.getEntry(FORMULA_FILE);
			if (ze != null) {
				System.out.println("Entry found");
				InputStream zis = zf.getInputStream(ze);
				FileOutputStream fos = new FileOutputStream(destFile);
		
		        byte[] bytesIn = new byte[BUFFER_SIZE];
		        int read = 0;
		        while ((read = zis.read(bytesIn)) != -1) {
		            fos.write(bytesIn, 0, read);
		        }
		        fos.close();
			}
			zf.close();
		} catch (java.io.IOException ioe) {
			System.out.println(ioe.getMessage());
			result = false;
		}
		return result;
	}
	
	public boolean writePackageParts(File destFile) {
		boolean result = true;
		try {
        	FileOutputStream fos = new FileOutputStream(destFile);
        	fos.write(bbPackageParts.array());
        	fos.close();
		} catch (IOException e) {
			System.out.println(e.getMessage());
			result = false;
		}
		return result;
	}
}
