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
        int objectid = 59595;
        int classifierid = 59595;
        int classifierattributeid = 862;
        int classifieritemid = 2322604;

        FileWriter writer = new FileWriter("/home/user/Downloads/Project/MKPO.sql");

        FileInputStream stream = new FileInputStream(new File("/home/user/Downloads/Project/MKPO2.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(stream);
        XSSFSheet sheet = workbook.getSheetAt(0);


        writer.write("insert into object (objectid,name,categoryid,isdeleted) values ("+objectid+",'Международная классификация промышленных образцов (МКПО)', 1, false);\n");
        writer.write("\n");

        writer.write("insert into classifier (classifierid,codeen,coderu,categoryid,status) values ("+classifierid+", 'МКПО', 'МКПО', 1,0);\n");
        writer.write("\n");

        String[] attributes_rus = {"Примечание","Тип", "Расшифровка"};
        for (int i = 0; i < 3; i++) {
            writer.write("insert into classifierattribute (classifierattributeid,classifierid,type,name,nameen,attributeorder) values ("+(classifierattributeid+i)+", "+classifierid+", 0, '"+attributes_rus[i]+"','"+attributes_rus[i]+"',"+(i+1)+");\n");
        }
        writer.write("\n");

        Iterator<Row> rowIterator = sheet.iterator();
        rowIterator.next();

        Row row;
        HashMap<String, Integer> parents = new HashMap();

        String parentId;
        int id = 0;

        while (rowIterator.hasNext()) {
            row = rowIterator.next();
            if (row.getCell(5).getNumericCellValue()<3) {
                parents.put(row.getCell(1).getStringCellValue(), classifieritemid);
            }

            if(row.getCell(7).getStringCellValue() == "") {
                parentId = "null";
            } else {
                parentId = row.getCell(7).getStringCellValue();
            }
            if (row.getCell(5).getNumericCellValue() == 3) {
                id = 3;
            } else {
                id = 1;
            }

            if(row.getCell(7).getStringCellValue() == "") {
                writer.write("insert into classifieritem (classifieritemid,classifierid,parentclassifieritemid,name,code) values ("+classifieritemid+", "+classifierid+", null, '"+row.getCell(2).getStringCellValue().replaceAll("'","''")+"', '"+row.getCell(id).getStringCellValue()+"');\n");
            } else {
                writer.write("insert into classifieritem (classifieritemid,classifierid,parentclassifieritemid,name,code) values ("+classifieritemid+", "+classifierid+", "+parents.get(parentId)+", '"+row.getCell(2).getStringCellValue().replaceAll("'","''")+"', '"+row.getCell(id).getStringCellValue()+"');\n");
            }

                writer.write("insert into classifieritemattributevalue (classifierattributeid,classifieritemid,textvalue) values (" +
                        (classifierattributeid) + ", " + classifieritemid + ", " + ((row.getCell(4) == null) ? "''" : "'" +
                        row.getCell(4).getStringCellValue().replaceAll("'", "''") + "'") + ");\n");
                writer.write("insert into classifieritemattributevalue (classifierattributeid,classifieritemid,textvalue) values (" +
                        (classifierattributeid+1) + ", " + classifieritemid + ", " + ((row.getCell(5) == null) ? "null" : "'" +
                        (int) row.getCell(5).getNumericCellValue() + "'") + ");\n");
                writer.write("insert into classifieritemattributevalue (classifierattributeid,classifieritemid,textvalue) values (" +
                        (classifierattributeid+2) + ", " + classifieritemid + ", " + ((row.getCell(6) == null) ? "null" : "'" +
                        row.getCell(6).getStringCellValue().replaceAll("'", "''") + "'") + ");\n");


            writer.write("\n");

            classifieritemid++;
        }

        writer.close();
        stream.close();
    }
}
