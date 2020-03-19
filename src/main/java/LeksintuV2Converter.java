import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;

public class LeksintuV2Converter {
    public static void main(String[] args) throws IOException {
        int objectid = 59602;
        int classifierid = 59602;
        int classifierattributeid = 871;
        int classifieritemid = 2341853;

        FileWriter writer = new FileWriter("/home/user/Downloads/Project/Leksintu2V2.sql");

        FileInputStream stream = new FileInputStream(new File("/home/user/Downloads/Project/Leksintu2.xls"));
        HSSFWorkbook workbook = new HSSFWorkbook(stream);
        HSSFSheet sheet = workbook.getSheetAt(1);

        writer.write("insert into object (objectid,name,categoryid,isdeleted) values ("+objectid+",'Предметный указатель товаров и услуг (ЛЕКСИНТУ Часть 2)', 1, false);\n");
        writer.write("\n");

        writer.write("insert into classifier (classifierid,codeen,coderu,categoryid,status) values ("+classifierid+", 'Предметный указатель товаров и услуг (ЛЕКСИНТУ Часть 2)', 'Предметный указатель товаров и услуг (ЛЕКСИНТУ Часть 2)', 1,0);\n");
        writer.write("\n");

        String[] attributes = {"note","head","class"};
        String[] attributes_rus = {"Основная группа","Дополнительная группа/подгруппа"};
        for (int i = 0; i < 2; i++) {
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
            if (row.getCell(0).getCellType() == CellType.STRING) {
                parents.put(row.getCell(0).getStringCellValue(), classifieritemid);
            } else if (row.getCell(0).getCellType() == CellType.NUMERIC) {
                parents.put(String.valueOf(row.getCell(0).getNumericCellValue()), classifieritemid);
            } else if (row.getCell(0).getCellType() == CellType.FORMULA) {
                parents.put(row.getCell(0).getStringCellValue(), classifieritemid);
            }

            if(row.getCell(4) ==null) {
                parentId = "null";
            } else {
                parentId = row.getCell(4).getStringCellValue();
            }

            writer.write("insert into classifieritem (classifieritemid,classifierid,parentclassifieritemid,name,code) values ("+
                    classifieritemid+", "+classifierid+", "+parents.get(parentId)+", '"+
                    (row.getCell(1).getCellType() == CellType.STRING ? row.getCell(1).getStringCellValue() :
                            (row.getCell(1).getCellType() == CellType.NUMERIC ? String.valueOf(row.getCell(1).getNumericCellValue()) :
                                    String.valueOf(row.getCell(1).getCachedFormulaResultType()))).replaceAll("'","''")+"', '"+
                    (row.getCell(0).getCellType() == CellType.STRING ? row.getCell(0).getStringCellValue() :
                            (row.getCell(0).getCellType() == CellType.NUMERIC ? String.valueOf(row.getCell(0).getNumericCellValue()) :
                                    row.getCell(0).getStringCellValue()))+"');\n");

            for (int i = 0; i <2; i++) {
                writer.write("insert into classifieritemattributevalue (classifierattributeid,classifieritemid,textvalue) values ("+(classifierattributeid+i)+", "+classifieritemid+", "+((row.getCell(i+2) == null) ? "null" : "'"+row.getCell(i+2).getStringCellValue().replaceAll("'","''")+"'")+");\n");
            }
            writer.write("\n");

            classifieritemid++;
        }

        writer.close();
        stream.close();
    }
}
