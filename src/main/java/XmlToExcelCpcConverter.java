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
        doc.getDocumentElement().normalize();

        NodeList nList = doc.getElementsByTagName("classification-item");
        for (int i = 0; i < nList.getLength(); i++) {
            Node node = nList.item(i);
            if (node.getNodeType() == Node.ELEMENT_NODE) {
                Element element  = (Element) node;
                System.out.println("===================================");
                String textContent = element.getElementsByTagName("classification-symbol").item(0).getTextContent();
                System.out.print(textContent+"~~~");
                System.out.print(element.getAttribute("level")+"~~~");

                Element title = (Element) element.getElementsByTagName("class-title").item(0);
                NodeList elementsByTagName = title.getElementsByTagName("title-part");
                for(int j=0; j<elementsByTagName.getLength();j++) {
                    Element item = (Element) elementsByTagName.item(j);
                    System.out.print(item.getTextContent());
                    if (elementsByTagName.item(j+1) != null) System.out.print(";");
                }

                Element warningsAndNotes = (Element) element.getElementsByTagName("notes-and-warnings").item(0);
                if (warningsAndNotes != null) {
                    Element warning = (Element) warningsAndNotes.getElementsByTagName("note").item(0);
                    System.out.println("+++++++++++++++++++++++++++++++++++");
                    System.out.println(warning.getTextContent());
                    System.out.println("");
                }
//                System.out.println(element.getElementsByTagName("class-title").item(0).getTextContent());
                System.out.println("===================================");
            }
        }
    }
}
