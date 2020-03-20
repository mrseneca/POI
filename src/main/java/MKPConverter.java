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

public class MKPConverter {
    public static void main(String[] args) throws IOException {
        int objectid = 59631;
        int classifierid = 59631;
        int classifierattributeid = 882;
        int classifieritemid = 2358165;

        FileWriter writer = new FileWriter("/home/user/Downloads/Project/MKP.sql");

        FileInputStream stream = new FileInputStream(new File("/home/user/Downloads/Project/ResultsMKPV2.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(stream);
        XSSFSheet sheet = workbook.getSheetAt(0);

        writer.write("insert into object (objectid,name,categoryid,isdeleted) values ("+objectid+",'Международная патентная классификация', 1, false);\n");
        writer.write("\n");

        writer.write("insert into classifier (classifierid,codeen,coderu,categoryid,status) values ("+classifierid+", 'МПК', 'МПК', 1,0);\n");
        writer.write("\n");

        String[] attributes_rus = {"Вид","Количество точек","Редакция","Версия","Назначение", "Уровень", "Заголовок","Примечание до","Примечание после","Содержание"};
        int ctype;
        for (int i = 0; i < 10; i++) {
            if (i==0||i==1||i==4||i==5) {ctype=1;} else {ctype=0;}
            writer.write("insert into classifierattribute (classifierattributeid,classifierid,type,name,nameen,attributeorder) values ("+(classifierattributeid+i)+", "+classifierid+", "+ctype+", '"+attributes_rus[i]+"','"+attributes_rus[i]+"',"+(i+1)+");\n");
        }
        writer.write("\n");

        Iterator<Row> rowIterator = sheet.iterator();
        rowIterator.next();

        Row row;
        HashMap<String, Integer> parents = new HashMap();

        String parentId;
        String cellText;
        String cellType;

        while (rowIterator.hasNext()) {
            row = rowIterator.next();
            parents.put(row.getCell(0).getStringCellValue(), classifieritemid);

            if(row.getCell(1) ==null) {
                parentId = "null";
            } else {
                parentId = row.getCell(1).getStringCellValue();
            }

            writer.write("insert into classifieritem (classifieritemid,classifierid,parentclassifieritemid,name,code) values ("+classifieritemid+", "+classifierid+", "+parents.get(parentId)+", '"+row.getCell(2).getStringCellValue().replaceAll("'","''")+"', '"+row.getCell(0).getStringCellValue()+"');\n");
            for (int i = 0; i <10; i++) {
                if (i==0||i==1||i==4||i==5) {cellType="numbervalue";} else {cellType="textvalue";}
//                cellText = (row.getCell(i+3).getCellType() == CellType.STRING) ? (row.getCell(i+3) == null) ? "null" : "'"+row.getCell(i+3).getStringCellValue().replaceAll("'","''")+"'" : (row.getCell(i+3) == null) ? "null" : "'"+String.valueOf(row.getCell(i+3).getNumericCellValue())+"'";
                cellText = (row.getCell(i+3)==null) ? "''" : (row.getCell(i+3).getCellType() == CellType.STRING ? "'"+row.getCell(i+3).getStringCellValue().replaceAll("'","''")+"'" : (row.getCell(i+3) == null) ? "null" : String.valueOf((int)row.getCell(i+3).getNumericCellValue()));
                writer.write("insert into classifieritemattributevalue (classifierattributeid,classifieritemid,"+cellType+") values ("+(classifierattributeid+i)+", "+classifieritemid+", "+cellText+");\n");
            }
            writer.write("\n");

            classifieritemid++;
        }

        writer.close();
        stream.close();
    }
}
