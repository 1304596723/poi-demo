package poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class Demo2 {

    public static void main(String[] args) throws IOException {
        Workbook workbook = new HSSFWorkbook();
        FileOutputStream fileOut = new FileOutputStream("c:\\poi\\poi创建sheet页.xls");
        Sheet sheet = workbook.createSheet("创建的一个sheet页");
        workbook.write(fileOut);
        fileOut.close();
    }
}
