import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;

public class UpdateMKPO {
    public static void main(String[] args) throws IOException {
        FileWriter writer = new FileWriter("/home/user/Downloads/Project/UPDATE.sql");

        FileInputStream stream = new FileInputStream(new File("/home/user/Downloads/Project/Results.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(stream);
        XSSFSheet sheet = workbook.getSheetAt(0);

        Iterator<Row> rowIterator = sheet.iterator();
        Row row;

        while (rowIterator.hasNext()) {

            row = rowIterator.next();
            if (row.getCell(1) != null) {
                writer.write("update classifieritemattributevalue as cav set textvalue = '" + row.getCell(1).getStringCellValue() + "' from  public.classifieritem as ci WHERE ci.classifierid = 59603 and cav.classifierattributeid = 873 and ci.code = '" + row.getCell(0).getStringCellValue() + "' and ci.classifieritemid = cav.classifieritemid;");
                writer.write("\n");
            }
        }
        stream.close();
        writer.close();
    }
}
