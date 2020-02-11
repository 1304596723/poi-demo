package poi;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Calendar;
import java.util.Date;

public class Demo4 {

    public static void main(String[] args) throws IOException {
         HSSFWorkbook wb = new HSSFWorkbook();
        OutputStream outputStream = new FileOutputStream("c:\\poi\\poi日期类型.xls");
        Sheet sheet =  wb.createSheet("第一个sheet页");
        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue("张三");
        row.createCell(1).setCellValue(true);
        Cell cell = row.createCell(2);
        HSSFCreationHelper creationHelper = wb.getCreationHelper();//创作助手（工具）
        HSSFDataFormat dataFormat = creationHelper.createDataFormat();
        HSSFCellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(dataFormat.getFormat("yyyy-MM-dd HH:mm:ss"));
        cell.setCellValue(Calendar.getInstance());
        cell.setCellStyle(cellStyle);

        Row row2 = sheet.createRow(1);
        row2.createCell(0).setCellValue("李四");
        row2.createCell(1).setCellValue(false);
        Cell cell1 = row2.createCell(2);
        cell1.setCellValue(new Date());
        cell1.setCellStyle(cellStyle);
        wb.write(outputStream);
        outputStream.close();
    }
}
