package poi;

import javafx.scene.control.Cell;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class Demo6 {

    public static void main(String[] args) throws IOException {
        InputStream is = new FileInputStream("c:\\poi\\poi日期类型.xls");
        POIFSFileSystem poi = new POIFSFileSystem(is);
        HSSFWorkbook wb = new HSSFWorkbook(poi);

        HSSFSheet sheetAt = wb.getSheetAt(0);
       // sheetAt.getLastRowNum();//获取最后行数
        for (int rorNum = 0;rorNum<=sheetAt.getLastRowNum();rorNum++){
            HSSFRow row = sheetAt.getRow(rorNum);
            if (row == null){
                continue;
            }
            for (int cellNum = 0;cellNum<=row.getLastCellNum();cellNum++){
                HSSFCell cell = row.getCell(cellNum);
                if (cell==null){
                    continue;
                }
                System.out.print("  "+getValue(cell));
            }
            System.out.println();
        }
    }

    public static String getValue(HSSFCell cell){
        String s ;
        switch (cell.getCellType()){
            case HSSFCell.CELL_TYPE_BOOLEAN:
                s = String.valueOf(cell.getBooleanCellValue());
                break;
            case HSSFCell.CELL_TYPE_STRING:
                s = String.valueOf(cell.getStringCellValue());
                break;
                default:
                    s = String.valueOf(cell.getDateCellValue());
        }
        return s;
    }
}
