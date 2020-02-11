package poi;

import javafx.scene.control.Cell;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.io.FileOutputStream;
import java.io.IOException;

public class Demo8 {
    public static void main(String[] args) throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet("第一个sheet页");
        HSSFRow row = sheet.createRow(2);
        row.setHeightInPoints((short) 30);
        HSSFCell cell = row.createCell(2);

        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setBorderBottom(CellStyle.BORDER_THICK);
        cellStyle.setBottomBorderColor(IndexedColors.RED.index);
        cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
        cellStyle.setLeftBorderColor(IndexedColors.BLACK.index);
        cellStyle.setBorderRight(CellStyle.BORDER_MEDIUM_DASH_DOT);
        cellStyle.setRightBorderColor(IndexedColors.GREEN.index);
        cellStyle.setBorderTop(CellStyle.BORDER_DASH_DOT_DOT);
        cellStyle.setTopBorderColor(IndexedColors.BLUE.index);
        cellStyle.setFillForegroundColor(IndexedColors.GREEN.index);
        cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        cell.setCellStyle(cellStyle);
        FileOutputStream fos = new FileOutputStream("c:\\poi\\样式.xls");
        wb.write(fos);
        fos.close();
    }
}
