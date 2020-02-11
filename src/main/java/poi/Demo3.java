package poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class Demo3 {
    public static void main(String[] args) throws IOException {
        Workbook workbook = new HSSFWorkbook();
        FileOutputStream fileOut = new FileOutputStream("c:\\poi\\poi创建row.xls");
        Sheet sheet = workbook.createSheet("创建的第一个sheet页");
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("张三");

        row.createCell(1).setCellValue(10);
        row.createCell(2).setCellValue(true);
        workbook.write(fileOut);
    }
}
