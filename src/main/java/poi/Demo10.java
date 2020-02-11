package poi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;

public class Demo10 {
    public static void main(String[] args) throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet("第一个sheet页");
        HSSFRow row = sheet.createRow(1);
        row.createCell(1).setCellValue("合并单元格");

        sheet.addMergedRegion(new CellRangeAddress(1,2,1,2));
        FileOutputStream fos = new FileOutputStream("c:\\poi\\样式.xls");
        wb.write(fos);
        fos.close();
    }
}
