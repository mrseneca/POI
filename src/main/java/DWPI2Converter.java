import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Iterator;

public class DWPI2Converter {
    public static void main(String[] args) throws IOException {
        int objectid = 59648;
        int classifierid = 59648;
        int classifierattributeid = 943;
        int classifieritemid = 2474020;

        FileWriter writer = new FileWriter("f:\\DWPI2.sql");

        FileInputStream stream = new FileInputStream(new File("f:\\Java\\JavaProjects\\POI\\src\\main\\java\\DWPI Masterlist 2020.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(stream);
        XSSFSheet sheet = workbook.getSheetAt(1);

        writer.write("insert into object (objectid,name,categoryid,isdeleted) values ("+objectid+",'DWPI - Masterlist Classes', 1, false);\n");
        writer.write("\n");

        writer.write("insert into classifier (classifierid,codeen,coderu,categoryid,status) values ("+classifierid+", 'DWPI Classes', 'DWPI Classes', 1,0);\n");
        writer.write("\n");

        String[] attributes = {"Scope Notes"};
        for (int i = 0; i < 1; i++) {
            writer.write("insert into classifierattribute (classifierattributeid,classifierid,type,name,nameen,attributeorder) values ("+(classifierattributeid+i)+", "+classifierid+", 0, '"+attributes[i]+"','"+attributes[i]+"',"+(i+1)+");\n");
        }
        writer.write("\n");

        Iterator<Row> rowIterator = sheet.iterator();
        rowIterator.next();

        Row row;
        while (rowIterator.hasNext()) {
            row = rowIterator.next();

            writer.write("insert into classifieritem (classifieritemid,classifierid,parentclassifieritemid,name,code) values ("+classifieritemid+", "+classifierid+", null, '"+row.getCell(1).getStringCellValue().replaceAll("'","''")+"', '"+row.getCell(0).getStringCellValue()+"');\n");
            for (int i = 0; i <1; i++) {
                writer.write("insert into classifieritemattributevalue (classifierattributeid,classifieritemid,textvalue) values ("+(classifierattributeid+i)+", "+classifieritemid+", "+((row.getCell(i+2) == null) ? "null" : "'"+row.getCell(i+2).getStringCellValue().replaceAll("'","''")+"'")+");\n");
            }
            writer.write("\n");

            classifieritemid++;
        }

        writer.close();
        stream.close();
    }
}
