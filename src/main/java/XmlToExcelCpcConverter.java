import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
        import org.w3c.dom.Element;
        import org.w3c.dom.Node;
        import org.w3c.dom.NodeList;
        import org.xml.sax.SAXException;

import javax.swing.text.Style;
import javax.xml.parsers.DocumentBuilder;
        import javax.xml.parsers.DocumentBuilderFactory;
        import javax.xml.parsers.ParserConfigurationException;
        import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.Collections;
import java.util.stream.Stream;

public class XmlToExcelCpcConverter {
    private static Workbook workbook;
    private static int rowNum;

    private final static int CODE = 0;
    private final static int PARENT = 1;
    private final static int TITLE = 2;
    private final static int CONCORDANT = 3;
    private final static int SORTKEY = 4;
    private final static int LEVEL = 5;
    private final static int WARNING = 6;
    private final static int NOTE = 7;

    public static void main(String[] args) throws ParserConfigurationException, IOException, SAXException {
        String code;
        String title;
        String warning;
        String note;
        String level;
        String sortKey;
        String concordant;
        String parent;

        workbook = new XSSFWorkbook();

        CellStyle style = workbook.createCellStyle();
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        style.setFont(boldFont);
        style.setAlignment(HorizontalAlignment.CENTER);

        Sheet sheet = workbook.createSheet();
        rowNum = 0;
        Row row = sheet.createRow(rowNum++);

        Cell cell = row.createCell(CODE);
        cell.setCellValue("CODE");
        cell.setCellStyle(style);

        cell = row.createCell(PARENT);
        cell.setCellValue("PARENT");
        cell.setCellStyle(style);

        cell = row.createCell(TITLE);
        cell.setCellValue("TITLE");
        cell.setCellStyle(style);


        cell = row.createCell(SORTKEY);
        cell.setCellValue("SORTKEY");
        cell.setCellStyle(style);

        cell = row.createCell(CONCORDANT);
        cell.setCellValue("CONCORDANT");
        cell.setCellStyle(style);

        cell = row.createCell(LEVEL);
        cell.setCellValue("LEVEL");
        cell.setCellStyle(style);

        cell = row.createCell(WARNING);
        cell.setCellValue("WARNING");
        cell.setCellStyle(style);

        cell = row.createCell(NOTE);
        cell.setCellValue("NOTE");
        cell.setCellStyle(style);

        File folder = new File("f:\\CPCWork\\");

        File[] folderEntries = folder.listFiles();
        Arrays.sort(folderEntries);
        for (File entry : folderEntries)
        {
            File xmlFile = new File(String.valueOf(entry));

            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document doc = dBuilder.parse(xmlFile);
            doc.getDocumentElement().normalize();

            NodeList classificationItems = doc.getElementsByTagName("classification-item");
            for (int i = 0; i < classificationItems.getLength(); i++) {
                code = ""; title = ""; warning = ""; note = ""; level = ""; sortKey = ""; concordant = "";
                Element item = (Element) classificationItems.item(i);
                level = item.getAttribute("level");
                if (level.equals("3")||level.equals("6")) {
                    continue;
                }
                sortKey = item.getAttribute("sort-key");
                concordant = item.getAttribute("ipc-concordant");
                if(level.equals("4")||level.equals("7")) {
                    Element itemParent = (Element) item.getParentNode().getParentNode();
                    parent = itemParent.getAttribute("sort-key");
                } else {
                    Element itemParent = (Element) item.getParentNode();
                    parent = itemParent.getAttribute("sort-key");
                }


//                System.out.println("=======");
//                System.out.println("PARENT: "+parent);
//                System.out.println(i+"~"+sortKey+"~"+level+"~"+concordant);
//                System.out.println("=======");

                NodeList itemChildren = classificationItems.item(i).getChildNodes();
                for (int j = 0; j<itemChildren.getLength(); j++){
                    if (itemChildren.item(j).getNodeType() == Node.ELEMENT_NODE) {
                        if (itemChildren.item(j).getNodeName()=="classification-item") {
                            continue;
                        } else if (itemChildren.item(j).getNodeName()=="classification-symbol") {
                            code = code + itemChildren.item(j).getTextContent();
                        } else if (itemChildren.item(j).getNodeName()=="class-title") {
                            NodeList titleChildren = itemChildren.item(j).getChildNodes();
                            for (int k = 0; k < titleChildren.getLength(); k++) {
                                if(titleChildren.item(k).getNodeName()=="title-part") {
                                    NodeList titleChildrenSub = titleChildren.item(k).getChildNodes();
                                    for (int l = 0; l < titleChildrenSub.getLength(); l++){
                                        if (titleChildrenSub.item(l).getNodeName()=="text") {
                                            title = title + titleChildrenSub.item(l).getTextContent();
                                        } else if (titleChildrenSub.item(l).getNodeName()=="reference") {
                                            title = title + "(" + titleChildrenSub.item(l).getTextContent() + ")";
                                        } else if (titleChildrenSub.item(l).getNodeName()=="CPC-specific-text") {
                                            title = title + "{" + titleChildrenSub.item(l).getTextContent() + "}";
                                        }
                                    }
                                    title = title  + (titleChildren.item(k+1)!=null ? ";" : "");
                                }
                            }
                        } else if (itemChildren.item(j).getNodeName()=="notes-and-warnings") {
                            NodeList noteChildren = itemChildren.item(j).getChildNodes();
                            for (int k = 0; k < noteChildren.getLength(); k++) {
                                if (noteChildren.item(k).getNodeType()==Node.ELEMENT_NODE){
                                    Element noteChildrenElement = (Element) noteChildren.item(k);
                                    if (noteChildrenElement.getAttribute("type").equals("warning")){
                                        NodeList noteChildrenSub = noteChildrenElement.getChildNodes();
                                        for (int l = 0; l < noteChildrenSub.getLength(); l++) {
                                            warning = warning + (noteChildrenSub.getLength()>1 ? String.valueOf(l+1)+". " : "") + noteChildrenSub.item(l).getTextContent()+"\n";
                                        }
                                    } else {
                                        NodeList noteChildrenSub = noteChildrenElement.getChildNodes();
                                        for (int l = 0; l < noteChildrenSub.getLength(); l++) {
                                            note = note + (noteChildrenSub.getLength()>1 ? String.valueOf(l+1)+". " : "") + noteChildrenSub.item(l).getTextContent()+"\n";
                                        }
                                    }
                                }
                            }

//                        note = itemChildren.item(j).getTextContent();
//                        System.out.println(j+"~"+itemChildren.item(j).getNodeName());
                        }
                    }
                }

                row = sheet.createRow(rowNum++);

                cell = row.createCell(CODE);
                cell.setCellValue(code);

                cell = row.createCell(PARENT);
                cell.setCellValue(parent);

                cell = row.createCell(TITLE);
                cell.setCellValue(title);

                cell = row.createCell(SORTKEY);
                cell.setCellValue(sortKey);

                cell = row.createCell(CONCORDANT);
                cell.setCellValue(concordant);

                cell = row.createCell(LEVEL);
                cell.setCellValue(level);

                cell = row.createCell(WARNING);
                cell.setCellValue(warning);

                cell = row.createCell(NOTE);
                cell.setCellValue(note);

//                System.out.println(code+"\n"+title+"\n"+note+"\n"+warning);
            }
        }



        FileOutputStream fileOut = new FileOutputStream("F:\\CPC-C.xlsx");
        workbook.write(fileOut);
        workbook.close();
        fileOut.close();
    }
}
