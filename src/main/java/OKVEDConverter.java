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

public class OKVEDConverter {
    public static void main(String[] args) throws IOException {
        int objectid = 59646;
        int classifierid = 59646;
        int classifierattributeid = 935;
        int classifieritemid = 2444316;

        FileWriter writer = new FileWriter("f:\\OKVED.sql");

        FileInputStream stream = new FileInputStream(new File("f:\\data-9749-2020-03-02_ к загрузке.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(stream);
        XSSFSheet sheet = workbook.getSheetAt(0);

        writer.write("insert into object (objectid,name,categoryid,isdeleted) values ("+objectid+",'Общероссийский классификатор видов экономической деятельности ОКВЭД 2 (ОК 029-2014 (КДЕС Ред. 2))', 1558425, false);\n");
        writer.write("\n");

        writer.write("insert into classifier (classifierid,codeen,coderu,categoryid,status) values ("+classifierid+", 'ОКВЭД-2', 'ОКВЭД-2', 1558425,0);\n");
        writer.write("\n");

        String[] attributes_rus = {"Примечание", "Раздел", "Индекс"};
        for (int i = 0; i < 3; i++) {
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
            
            if(row.getCell(5) ==null) {
                parentId = "null";
            } else {
                parentId = row.getCell(5).getStringCellValue();
            }

            writer.write("insert into classifieritem (classifieritemid,classifierid,parentclassifieritemid,name,code) values ("+classifieritemid+", "+classifierid+", "+parents.get(parentId)+", '"+row.getCell(1).getStringCellValue().replaceAll("'","''")+"', '"+row.getCell(0).getStringCellValue()+"');\n");
            for (int i = 0; i <3; i++) {
                writer.write("insert into classifieritemattributevalue (classifierattributeid,classifieritemid,textvalue) values ("+(classifierattributeid+i)+", "+classifieritemid+", "+((row.getCell(i+2) == null) ? "null" : "'"+row.getCell(i+2).getStringCellValue().replaceAll("'","''")+"'")+");\n");
            }
            writer.write("\n");

            classifieritemid++;
        }

        writer.close();
        stream.close();
    }
}
