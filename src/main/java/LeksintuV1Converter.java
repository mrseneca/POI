import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;

public class LeksintuV1Converter {
    public static void main(String[] args) throws IOException {
        int objectid = 59592;
        int classifierid = 59592;
        int classifieritemid = 2299372;

        FileWriter writer = new FileWriter("/home/user/Downloads/Project/Leksintu1.sql");

        FileInputStream stream = new FileInputStream(new File("/home/user/Downloads/Project/Leksintu1.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(stream);
        XSSFSheet sheet = workbook.getSheetAt(0);


        writer.write("insert into object (objectid,name,categoryid,isdeleted) values ("+objectid+",'Указатель содержания классов товаров и услуг (ЛЕКСИНТУ Часть1)', 1, false);\n");
        writer.write("\n");

        writer.write("insert into classifier (classifierid,codeen,coderu,categoryid,status) values ("+classifierid+", 'Указатель содержания классов товаров и услуг (ЛЕКСИНТУ Часть1)', 'Указатель содержания классов товаров и услуг (ЛЕКСИНТУ Часть1)', 1,0);\n");
        writer.write("\n");


        Iterator<Row> rowIterator = sheet.iterator();
        rowIterator.next();

        Row row;
        HashMap<String, Integer> parents = new HashMap();

        String parentId;

        while (rowIterator.hasNext()) {
            row = rowIterator.next();
            parents.put(row.getCell(0).getStringCellValue(), classifieritemid);

            if(row.getCell(2) ==null) {
                parentId = "null";
            } else {
                parentId = row.getCell(2).getStringCellValue();
            }

            writer.write("insert into classifieritem (classifieritemid,classifierid,parentclassifieritemid,name,code) values ("+classifieritemid+", "+classifierid+", "+parents.get(parentId)+", '"+row.getCell(1).getStringCellValue().replaceAll("'","''")+"', '"+row.getCell(0).getStringCellValue()+"');\n");
            writer.write("\n");

            classifieritemid++;
        }

        writer.close();
        stream.close();
    }

}
