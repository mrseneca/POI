import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;

import java.io.*;
import java.util.*;

public class ConverterVienna {
    public static void main(String[] args) throws IOException {
        int objectid = 59601;
        int classifierid = 59601;
        int classifierattributeid = 868;
        int classifieritemid = 2339945;

        FileWriter writer = new FileWriter("/home/user/Downloads/Project/viennaV3.sql");

        FileInputStream stream = new FileInputStream(new File("/home/user/Downloads/Project/viena.xls"));
        HSSFWorkbook workbook = new HSSFWorkbook(stream);
        HSSFSheet sheet = workbook.getSheetAt(0);

        writer.write("insert into object (objectid,name,categoryid,isdeleted) values ("+objectid+",'Венский классификатор изобразительных элементов (ВКЛ)', 1, false);\n");
        writer.write("\n");

        writer.write("insert into classifier (classifierid,codeen,coderu,categoryid,status) values ("+classifierid+", 'Венский классификатор изобразительных элементов (ВКЛ)', 'Венский классификатор изобразительных элементов (ВКЛ)', 1,0);\n");
        writer.write("\n");

        String[] attributes = {"note","head","class"};
        String[] attributes_rus = {"Примечание","Оглавление","Класс"};
        for (int i = 0; i < 3; i++) {
            writer.write("insert into classifierattribute (classifierattributeid,classifierid,type,name,nameen,attributeorder) values ("+(classifierattributeid+i)+", "+classifierid+", 0, '"+attributes_rus[i]+"','"+attributes[i]+"',"+(i+1)+");\n");
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

            if(row.getCell(1) ==null) {
                parentId = "null";
            } else {
                parentId = row.getCell(1).getStringCellValue();
            }

            writer.write("insert into classifieritem (classifieritemid,classifierid,parentclassifieritemid,name,code) values ("+classifieritemid+", "+classifierid+", "+parents.get(parentId)+", '"+row.getCell(2).getStringCellValue().replaceAll("'","''")+"', '"+row.getCell(0).getStringCellValue()+"');\n");
            for (int i = 0; i <3; i++) {
                writer.write("insert into classifieritemattributevalue (classifierattributeid,classifieritemid,textvalue) values ("+(classifierattributeid+i)+", "+classifieritemid+", "+((row.getCell(i+3) == null) ? "null" : "'"+row.getCell(i+3).getStringCellValue().replaceAll("'","''")+"'")+");\n");
            }
            writer.write("\n");

            classifieritemid++;
        }

        writer.close();
        stream.close();
    }
}
