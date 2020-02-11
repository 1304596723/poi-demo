package poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class Demo1 {

    public static void main(String[] args)throws IOException {
       Workbook workbook = new HSSFWorkbook();
        FileOutputStream output = new FileOutputStream("c:\\poi\\工作簿poi.xsl");
        workbook.write(output);
        output.close();
    }
}
