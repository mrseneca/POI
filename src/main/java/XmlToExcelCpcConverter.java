import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

public class XmlToExcelCpcConverter {
    public static void main(String[] args) throws ParserConfigurationException, IOException, SAXException {
        File xmlFile = new File("f:\\CPC\\CPCSchemeXML202002\\cpc-scheme-A.xml");

        DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
        DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
        Document doc = dBuilder.parse(xmlFile);

        NodeList nList = doc.getElementsByTagName("classification-item");
        for (int i = 0; i < 1; i++) {
            Node node = nList.item(i);
            if (node.getNodeType() == Node.ELEMENT_NODE) {
                Element element  = (Element) node;
                System.out.println("===================================");

                System.out.println("===================================");
            }
        }
    }
}
