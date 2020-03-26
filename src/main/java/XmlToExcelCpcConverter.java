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
        File xmlFile = new File("f:\\CPC\\CPCSchemeXML202002\\cpc-scheme-A01B.xml");

        DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
        DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
        Document doc = dBuilder.parse(xmlFile);
        doc.getDocumentElement().normalize();
        String code = "";
        String title = "";
        String warning = "";
        String note = "";
        String level = "";
        String sortKey = "";
        String concordant = "";

        NodeList classificationItems = doc.getElementsByTagName("classification-item");
        for (int i = 0; i < 6; i++) {
            code = ""; title = ""; warning = ""; note = ""; level = ""; sortKey = ""; concordant = "";
            Element item = (Element) classificationItems.item(i);
            level = item.getAttribute("level");
            sortKey = item.getAttribute("sort-key");
            concordant = item.getAttribute("ipc-concordant");

            System.out.println("=======");
            System.out.println(i+"~"+sortKey+"~"+level+"~"+concordant);
            System.out.println("=======");

            NodeList itemChildren = classificationItems.item(i).getChildNodes();
            for (int j = 0; j<itemChildren.getLength(); j++){
                if (itemChildren.item(j).getNodeType() == Node.ELEMENT_NODE) {
                    if (itemChildren.item(j).getNodeName()=="classification-item") {
                        continue;
                    } else if (itemChildren.item(j).getNodeName()=="classification-symbol") {
                        code = itemChildren.item(j).getTextContent();
//                        System.out.println(j+"~"+itemChildren.item(j).getNodeName());
                    } else if (itemChildren.item(j).getNodeName()=="class-title") {
                        NodeList titleChildren = itemChildren.item(j).getChildNodes();
                        for (int k = 0; k < titleChildren.getLength(); k++) {
                            if(titleChildren.item(k).getNodeName()=="title-part") {
                                title = title + titleChildren.item(k).getTextContent() + (titleChildren.item(k+1)!=null ? ";" : "");
                            }
                        }
//                        title = itemChildren.item(j).getTextContent();;
//                        System.out.println(j+"~"+itemChildren.item(j).getNodeName());
                    } else if (itemChildren.item(j).getNodeName()=="notes-and-warnings") {
                        note = itemChildren.item(j).getTextContent();
//                        System.out.println(j+"~"+itemChildren.item(j).getNodeName());
                    }
                }
            }
            System.out.println(code+"*"+title+"*"+note);
        }
    }
}
