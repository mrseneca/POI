import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;

public class RPConverter {
    public static void main(String[] args) throws IOException {
        int objectid = 59644;
        int classifierid = 59644;
        int classifierattributeid = 932;
        int classifieritemid = 2443948;

        FileWriter writer = new FileWriter("f:\\RP.sql");

        FileInputStream stream = new FileInputStream(new File("f:\\Структура РП.xls"));
        HSSFWorkbook workbook = new HSSFWorkbook(stream);
        HSSFSheet sheet = workbook.getSheetAt(0);

        writer.write("insert into object (objectid,name,categoryid,isdeleted) values ("+objectid+",'Справочник организационной структуры организации', 1, false);\n");
        writer.write("\n");

        writer.write("insert into classifier (classifierid,codeen,coderu,categoryid,status) values ("+classifierid+", 'ORG_FIPS', 'ORG_FIPS', 1,0);\n");
        writer.write("\n");

        String[] attributes_rus = {"Наименование организации"};
        for (int i = 0; i < 1; i++) {
            writer.write("insert into classifierattribute (classifierattributeid,classifierid,type,name,nameen,attributeorder) values ("+(classifierattributeid+i)+", "+classifierid+", 0, '"+attributes_rus[i]+"','"+attributes_rus[i]+"',"+(i+1)+");\n");
        }
        writer.write("\n");

        Iterator<Row> rowIterator = sheet.iterator();
        rowIterator.next();

        Row row;
        HashMap<String, Integer> parents = new HashMap();

        String parentId;

        while (rowIterator.hasNext()) {
            row = rowIterator.next();
            parents.put(row.getCell(0).getStringCellValue(), classifieritemid);

            if(row.getCell(3) ==null) {
                parentId = "null";
            } else {
                parentId = row.getCell(3).getStringCellValue();
            }

            writer.write("insert into classifieritem (classifieritemid,classifierid,parentclassifieritemid,name,code) values ("+classifieritemid+", "+classifierid+", "+parents.get(parentId)+", '"+row.getCell(1).getStringCellValue().replaceAll("'","''")+"', '"+row.getCell(0).getStringCellValue()+"');\n");
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
