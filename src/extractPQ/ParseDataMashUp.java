package extractPQ;

import java.io.File;
import java.util.Base64;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathFactory;

import org.w3c.dom.Document;

public class ParseDataMashUp {
	private DecodeQDEFF pqDecode = null;
	
	public boolean getXPathValue(File dmuFile, String targetDir) {
		boolean result = true;
		String dmuValue = new String();
        DocumentBuilderFactory domFactory = DocumentBuilderFactory.newInstance();
        try {
            DocumentBuilder builder = domFactory.newDocumentBuilder();
            Document dDoc = builder.parse(dmuFile);

            XPath xPath = XPathFactory.newInstance().newXPath();
            dmuValue = (String) xPath.evaluate("/DataMashup", dDoc, XPathConstants.STRING);
            System.out.println(dmuValue);
            System.out.println(base64Decode(dmuValue));
            pqDecode = new DecodeQDEFF();
            pqDecode.extractQDEFF(base64Decode(dmuValue), new File(targetDir+"/output.zip"));
            
        } catch (Exception e) {
            e.printStackTrace();
            result = false;
        }
        return result;
    }

	public boolean extractFormula(String targetDir) {
		boolean result = false;
		PackageParts  pParts = new PackageParts(pqDecode.getPackageParts());
		File inFile = new File(targetDir+"/packageParts.zip");
		result = pParts.writePackageParts(inFile);
		result = pParts.extractFormula(inFile, new File(targetDir+"/fomula.m"));
		return result;
	}
	
	private byte[] base64Decode(String value) {
		byte[] decodedValue = Base64.getDecoder().decode(value);
        return decodedValue;
    }
}
