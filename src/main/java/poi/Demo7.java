package poi;

import org.apache.poi.POITextExtractor;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class Demo7 {
    public static void main(String[] args) throws IOException {
        InputStream is = new FileInputStream("c:\\poi\\poi日期类型.xls");
        POIFSFileSystem fileSystem = new POIFSFileSystem(is);
        HSSFWorkbook wb = new HSSFWorkbook(fileSystem);
        ExcelExtractor ex = new ExcelExtractor(wb);
        ex.setIncludeSheetNames(false);
        System.out.println(ex.getText());
    }
}
