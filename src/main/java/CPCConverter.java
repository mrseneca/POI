<<<<<<< HEAD
import org.apache.poi.ss.usermodel.CellType;
=======
>>>>>>> master
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;

public class CPCConverter {
    public static void main(String[] args) throws IOException {
<<<<<<< HEAD
        int objectid = 59679;
        int classifierid = 59679;
        int classifierattributeid = 978;
        int classifieritemid = 2507407;
=======
        int objectid = 59647;
        int classifierid = 59647;
        int classifierattributeid = 938;
        int classifieritemid = 2447114;
>>>>>>> master

        FileWriter writer = new FileWriter("f:\\CPC.sql");


        writer.write("insert into object (objectid,name,categoryid,isdeleted) values ("+objectid+",'Cooperative Patent Classification', 1, false);\n");
        writer.write("\n");

<<<<<<< HEAD
        writer.write("insert into classifier (classifierid,codeen,coderu,categoryid,status) values ("+classifierid+", '(CPC)', '(CPC)', 1,0);\n");
        writer.write("\n");

        String[] attributes = {"SortCode","Level","Note","Warning"};
        int ctype;
        for (int i = 0; i < 4; i++) {
            if (i==1) {ctype=1;} else {ctype=0;}
            writer.write("insert into classifierattribute (classifierattributeid,classifierid,type,name,nameen,attributeorder) values ("+(classifierattributeid+i)+", "+classifierid+", "+ctype+", '"+attributes[i]+"','"+attributes[i]+"',"+(i+1)+");\n");
=======
        writer.write("insert into classifier (classifierid,codeen,coderu,categoryid,status) values ("+classifierid+", 'CPC', 'CPC', 1,0);\n");
        writer.write("\n");

        String[] attributes = {"Уровень","Примечание","Предупреждение"};
        for (int i = 0; i < 3; i++) {
            writer.write("insert into classifierattribute (classifierattributeid,classifierid,type,name,nameen,attributeorder) values ("+(classifierattributeid+i)+", "+classifierid+", 0, '"+attributes[i]+"','"+attributes[i]+"',"+(i+1)+");\n");
>>>>>>> master
        }
        writer.write("\n");
        writer.close();

        File folder = new File("f:\\CPCWork\\");

        File[] folderEntries = folder.listFiles();
        Arrays.sort(folderEntries);
        for (File entry : folderEntries) {
            writer = new FileWriter("f:\\"+String.valueOf(entry).substring(11,String.valueOf(entry).length())+".sql");
            FileInputStream stream = new FileInputStream(new File(String.valueOf(entry)));
            XSSFWorkbook workbook = new XSSFWorkbook(stream);
            XSSFSheet sheet = workbook.getSheetAt(0);

            Iterator<Row> rowIterator = sheet.iterator();
            rowIterator.next();

            Row row;
            HashMap<String, Integer> parents = new HashMap();
            String parentId;
<<<<<<< HEAD
            String cellText;
            String cellType;
=======
>>>>>>> master

            while (rowIterator.hasNext()) {
                row = rowIterator.next();

<<<<<<< HEAD
                parents.put(row.getCell(4).getStringCellValue(), classifieritemid);
=======
                parents.put(row.getCell(0).getStringCellValue(), classifieritemid);
>>>>>>> master
                if(row.getCell(1) ==null) {
                    parentId = "null";
                } else {
                    parentId = row.getCell(1).getStringCellValue();
                }


                writer.write("insert into classifieritem (classifieritemid,classifierid,parentclassifieritemid,name,code) values (" + classifieritemid + ", " + classifierid + ","+parents.get(parentId)+", '" + row.getCell(2).getStringCellValue().replaceAll("'", "''") + "', '" + row.getCell(0).getStringCellValue() + "');\n");
<<<<<<< HEAD
                for (int i = 0; i < 4; i++) {
                    if (i==1) {cellType="numbervalue";} else {cellType="textvalue";}
                    cellText = (row.getCell(i+4)==null) ? (i==1 ? "null" : "''") : (i != 1 ? "'"+row.getCell(i+4).getStringCellValue().replaceAll("'","''")+"'" : row.getCell(i+4).getStringCellValue().replaceAll("'","''"));
                    writer.write("insert into classifieritemattributevalue (classifierattributeid,classifieritemid,"+cellType+") values (" + (classifierattributeid + i) + ", " + classifieritemid + ", " + cellText + ");\n");
=======
                for (int i = 0; i < 3; i++) {
                    writer.write("insert into classifieritemattributevalue (classifierattributeid,classifieritemid,textvalue) values (" + (classifierattributeid + i) + ", " + classifieritemid + ", " + ((row.getCell(i + 5) == null) ? "null" : "'" + row.getCell(i + 5).getStringCellValue().replaceAll("'", "''") + "'") + ");\n");
>>>>>>> master
                }
                writer.write("\n");

                classifieritemid++;
            }
            stream.close();
            writer.close();
        }
    }
}
