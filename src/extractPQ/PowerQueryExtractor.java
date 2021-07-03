package extractPQ;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathFactory;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/**
 *
 */
public class PowerQueryExtractor {
    private static final String itemProps = "customXml/itemProps1.xml";
    private static final String DataMashUpURI = "http://schemas.microsoft.com/DataMashup";
    private static boolean DEBUG=true;
    private static final int BUFFER_SIZE = 4096;
    ExtractQDEFF qdeff = null;
    ParseDataMashUp pdmu = null;

    public PowerQueryExtractor() {
        qdeff = new ExtractQDEFF();
        pdmu = new ParseDataMashUp();
    }

    /**
     *
     * @param file
     * @param targetDir
     * @return
     */
    public boolean ExtractPQ(File file, String targetDir) {
        boolean result = false;
        if (hasPowerQuery(file)) {
            String targetPath = String.format("%s/%s", targetDir, file.getName());
            File destFile = new File(targetPath+"/item1.xml");
            int qdeff_result = qdeff.extractPQFile(file, destFile);
            if (qdeff_result == ExtractQDEFF.OK) {
                if (pdmu.getXPathValue(destFile, targetPath)) {
                    result = pdmu.extractFormula(targetPath);
                    result = pdmu.extractPermissions(targetPath);
                    result = pdmu.extractPermissionBinding(targetPath);
                    result = pdmu.extractMetaData(targetPath);
                }
            }
        }
        return result;
    }

    /**
     *
     * @param zipFile
     * @return
     */
    private boolean hasPowerQuery(File zipFile) {
        boolean result = false;
        File destFile = null;
        try {
            ZipFile zf = new ZipFile(zipFile);
            ZipEntry ze = zf.getEntry(itemProps);
            destFile = File.createTempFile("tmpItemProps", null);
            if (ze != null) {
                InputStream zis = zf.getInputStream(ze);
//                if (DEBUG) System.out.printf("Trying to extract %s\n", zipFile.getName());
                FileOutputStream fos = new FileOutputStream(destFile);
                byte[] bytesIn = new byte[BUFFER_SIZE];
                int read = 0;
                while ((read = zis.read(bytesIn)) != -1) {
                    fos.write(bytesIn, 0, read);
                }
                fos.close();
                result = parseSchemaRef(destFile);
            }
            zf.close();
        } catch (java.io.IOException ioe) {
            System.out.printf("Extracting failed : %s\n", ioe.getMessage());
            if (destFile != null) {
                destFile.deleteOnExit();
            }
            return false;
        }
        //return result;
        return true;
    }

    /**
     *  The itemProps[x].xml contains information about the payload inside item[x].xml. In order to
     *  determine what kind of payload is available in the item[x].xml, we take a look at the schemaRef.
     *  If the uri attribute of the schemaRef element contains the reference to the DataMashup structure,
     *  we can conclude that the item[x].xml contains the DataMashup data where PowerQuery formulas are stored.
     *
     * @param destFile
     * @return
     */
    private boolean parseSchemaRef(File destFile) {
        boolean result = false;
        DocumentBuilderFactory domFactory = DocumentBuilderFactory.newInstance();
        try {
            DocumentBuilder builder = domFactory.newDocumentBuilder();
            Document dDoc = builder.parse(new FileInputStream(destFile));
            XPath xPath = XPathFactory.newInstance().newXPath();
            Object objSchemaRef = xPath.evaluate("/*[local-name() = 'datastoreItem']/*[local-name() = 'schemaRefs']/*[local-name() = 'schemaRef']", dDoc, XPathConstants.NODE);
            String strURI = ((Element) objSchemaRef).getAttribute("ds:uri");
            System.out.printf("Datastore Schema is : %s\n", strURI);
            if (strURI.equals(DataMashUpURI)) {
                result = true;
                System.out.println("Datastore Schema references DataMashUp, so it contains PowerQuery.");
            }
        } catch (Exception e) {
            if (DEBUG) {e.printStackTrace();}
//            System.out.printf("Error in determining schemaRef: %s for file: %s\n", e.getMessage(), destFile.getName());
            result = false;
        }
        destFile.deleteOnExit();
        return result;
    }
}
