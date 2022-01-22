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
    public static final int PQ_ERROR = -1;
    public static final int PQ_OK = 0;
    public static final int PQ_FOUND = 1;
    public static final int PQ_OTHER_DMU = 2;
    private static final String itemProps = "customXml/itemProps1.xml";
    private static final String DataMashUpURI = "http://schemas.microsoft.com/DataMashup";
    private String strURI;
    private static boolean DEBUG=true;
    private static final int BUFFER_SIZE = 4096;
    ExtractQDEFF qdeff;
    ParseDataMashUp pdmu;

    public PowerQueryExtractor() {
        qdeff = new ExtractQDEFF();
        pdmu = new ParseDataMashUp();
    }

    /**
     *
     * @param file  File of the spreadsheet that needs analyzing
     * @param targetDir Directory where intermediate files need to be stored
     * @return  boolean indicating if the spreddsheet contains a PQ
     */
    public int ExtractPQ(File file, String targetDir) {
        int result = PQ_OK;
        int hasPQ = hasPowerQuery(file);
        if (hasPQ == PQ_FOUND) {
            String targetPath = String.format("%s/%s", targetDir, file.getName());
            File destFile = new File(targetPath+"/item1.xml");
            int qdeff_result = qdeff.extractPQFile(file, destFile);
            if (qdeff_result == ExtractQDEFF.OK) {
                if (pdmu.getXPathValue(destFile, targetPath)) {
                    pdmu.extractFormula(targetPath);
                    pdmu.extractPermissions(targetPath);
                    pdmu.extractPermissionBinding(targetPath);
                    pdmu.extractMetaData(targetPath);
                    result = PQ_FOUND;
                }
            }
        }
        if (hasPQ == PQ_OTHER_DMU) {
            result = PQ_OTHER_DMU;
        }
        return result;
    }

    /**
     *
     * @param zipFile   ZIP file containing the DataMashUp part that needs to be analyzed
     * @return boolean
     */
    private int hasPowerQuery(File zipFile) {
        int result = PQ_OK;
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
            result = PQ_ERROR;
        }
        //return result;
        return result;
    }

    /**
     *  The itemProps[x].xml contains information about the payload inside item[x].xml. In order to
     *  determine what kind of payload is available in the item[x].xml, we take a look at the schemaRef.
     *  If the uri attribute of the schemaRef element contains the reference to the DataMashup structure,
     *  we can conclude that the item[x].xml contains the DataMashup data where PowerQuery formulas are stored.
     *
     * @param destFile  File where the PQ source needs to be stored
     * @return  result
     */
    private int parseSchemaRef(File destFile) {
        int result = PQ_OK;
        DocumentBuilderFactory domFactory = DocumentBuilderFactory.newInstance();
        try {
            DocumentBuilder builder = domFactory.newDocumentBuilder();
            Document dDoc = builder.parse(new FileInputStream(destFile));
            XPath xPath = XPathFactory.newInstance().newXPath();
            Object objSchemaRef = xPath.evaluate("/*[local-name() = 'datastoreItem']/*[local-name() = 'schemaRefs']/*[local-name() = 'schemaRef']", dDoc, XPathConstants.NODE);
            strURI = ((Element) objSchemaRef).getAttribute("ds:uri");
            System.out.printf("Datastore Schema is : %s\n", strURI);
            if (strURI.equals(DataMashUpURI)) {
                result = PQ_FOUND;
                System.out.println("Datastore Schema references DataMashUp, so it contains PowerQuery.");
            } else {
                if (strURI.length() > 5) {
                    result = PQ_OTHER_DMU;
                }
            }
        } catch (Exception e) {
            if (DEBUG) {e.printStackTrace();}
//            System.out.printf("Error in determining schemaRef: %s for file: %s\n", e.getMessage(), destFile.getName());
            result = PQ_ERROR;
        }
        destFile.deleteOnExit();
        return result;
    }

    public String getDataMashUpURI() {
        return strURI;
    }
}
