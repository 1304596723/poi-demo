package poi;

import org.apache.poi.hssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

public class Demo9 {

    public static void main(String[] args) throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet("第一个sheet页");
        HSSFRow row = sheet.createRow(2);
        row.setHeightInPoints((short) 30);
        createCell(wb,row,0,HSSFCellStyle.ALIGN_CENTER,HSSFCellStyle.VERTICAL_BOTTOM);
        FileOutputStream fos = new FileOutputStream("c:\\poi\\样式.xls");
        wb.write(fos);
        fos.close();
    }

    public static void createCell(HSSFWorkbook wb,HSSFRow row,int column,short  halign,short valign){
        HSSFCellStyle cellStyle = wb.createCellStyle();
        HSSFCell cell = row.createCell(column);
        cellStyle.setAlignment(halign);
        cellStyle.setVerticalAlignment(valign);
        cell.setCellValue(new HSSFRichTextString("zhangsan"));
        cell.setCellStyle(cellStyle);
    }
}
