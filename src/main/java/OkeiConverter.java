import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;

public class OkeiConverter {
    public static void main(String[] args) throws IOException {
        int objectid = 59591;
        int classifierid = 59591;
        int classifierattributeid = 853;
        int classifieritemid = 2298851;

        FileWriter writer = new FileWriter("/home/user/Downloads/Project/Okei.sql");

        FileInputStream stream = new FileInputStream(new File("/home/user/Downloads/Project/ОКЕИ.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(stream);
        XSSFSheet sheet = workbook.getSheetAt(0);

        writer.write("insert into object (objectid,name,categoryid,isdeleted) values ("+objectid+",'Общероссийский классификатор единиц измерения ОКЕИ (ОК 015-94 (МК 002-97))', 1, false);\n");
        writer.write("\n");

        writer.write("insert into classifier (classifierid,codeen,coderu,categoryid,status) values ("+classifierid+", 'ОК 015-94 (МК 002-97)', 'ОК 015-94 (МК 002-97)', 1,0);\n");
        writer.write("\n");

        String[] attributes = {"note","head","class"};
        String[] attributes_rus = {"Национальное обозначение","Международное обозначение","Национальное буквенное обозначение","Международное буквенное обозначение", "Код базового элемента", "Группа", "Подгруппа"};
        for (int i = 0; i < 7; i++) {
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

            if(row.getCell(11) ==null) {
                parentId = "null";
            } else {
                parentId = row.getCell(11).getStringCellValue();
            }

            writer.write("insert into classifieritem (classifieritemid,classifierid,parentclassifieritemid,name,code) values ("+classifieritemid+", "+classifierid+", "+parents.get(parentId)+", '"+row.getCell(1).getStringCellValue().replaceAll("'","''")+"', '"+row.getCell(0).getStringCellValue()+"');\n");
            for (int i = 0; i <7; i++) {
                writer.write("insert into classifieritemattributevalue (classifierattributeid,classifieritemid,textvalue) values ("+(classifierattributeid+i)+", "+classifieritemid+", "+((row.getCell(i+2) == null) ? "null" : "'"+row.getCell(i+2).getStringCellValue().replaceAll("'","''")+"'")+");\n");
            }
            writer.write("\n");

            classifieritemid++;
        }

        writer.close();
        stream.close();
    }
}
