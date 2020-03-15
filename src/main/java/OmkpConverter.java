import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;

public class OmkpConverter {
    public static void main(String[] args) throws IOException {
        int objectid = 59592;
        int classifierid = 59592;
        int classifierattributeid = 860;
        int classifieritemid = 2299372;

        FileWriter writer = new FileWriter("f:\\OMKP.sql");

        FileInputStream stream = new FileInputStream(new File("f:\\MKPO.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(stream);
        XSSFSheet sheet = workbook.getSheetAt(0);


        writer.write("insert into object (objectid,name,categoryid,isdeleted) values ("+objectid+",'Международная классификация промышленных образцов (МКПО)', 1, false);\n");
        writer.write("\n");

        writer.write("insert into classifier (classifierid,codeen,coderu,categoryid,status) values ("+classifierid+", 'МКПО', 'МКПО', 1,0);\n");
        writer.write("\n");

        String[] attributes_rus = {"Примечание","Тип", "Расшифровка"};
        for (int i = 0; i < 3; i++) {
            writer.write("insert into classifierattribute (classifierattributeid,classifierid,type,name,nameen,attributeorder) values ("+(classifierattributeid)+", "+classifierid+", 0, '"+attributes_rus[i]+"','"+attributes_rus[i]+"',"+(1)+");\n");
        }



        Iterator<Row> rowIterator = sheet.iterator();
        rowIterator.next();

        Row row;
        HashMap<String, Integer> parents = new HashMap();

        String parentId;
        int id = 0;

        while (rowIterator.hasNext()) {
            row = rowIterator.next();
            if (row.getCell(2).getNumericCellValue()<3) {
                parents.put(row.getCell(1).getStringCellValue(), classifieritemid);
            }

            if(row.getCell(6).getStringCellValue() == "") {
                parentId = "null";
            } else {
                parentId = row.getCell(6).getStringCellValue();
            }
            if (row.getCell(2).getNumericCellValue() == 3) {
                id = 4;
            } else {
                id = 1;
            }

            if(row.getCell(6).getStringCellValue() == "") {
                writer.write("insert into classifieritem (classifieritemid,classifierid,parentclassifieritemid,name,code) values ("+classifieritemid+", "+classifierid+", null, '"+row.getCell(3).getStringCellValue().replaceAll("'","''")+"', '"+row.getCell(id).getStringCellValue()+"');\n");
            } else {
                writer.write("insert into classifieritem (classifieritemid,classifierid,parentclassifieritemid,name,code) values ("+classifieritemid+", "+classifierid+", "+parents.get(parentId)+", '"+row.getCell(3).getStringCellValue().replaceAll("'","''")+"', '"+row.getCell(id).getStringCellValue()+"');\n");
            }
            for (int i = 0; i < 0; i++) {
                writer.write("insert into classifieritemattributevalue (classifierattributeid,classifieritemid,textvalue) values (" + (classifierattributeid) + ", " + classifieritemid + ", " + ((row.getCell(i+5) == null) ? "null" : "'" + row.getCell(i+5).getStringCellValue().replaceAll("'", "''") + "'") + ");\n");
            }
            writer.write("\n");

            classifieritemid++;
        }

        writer.close();
        stream.close();
    }
}
