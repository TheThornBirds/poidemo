package output;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;

/**
 * Created by Administrator on 2017/9/29.
 */
public class PoiDemo10 {
    public static void main(String[] args) throws Exception {
        InputStream input = new FileInputStream("d:\\非费用付款申请报表导出.xls");
        POIFSFileSystem fs = new POIFSFileSystem(input);
        Workbook wb = new HSSFWorkbook(fs);
        Sheet sheet = wb.getSheetAt(0);
        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);

        if (cell!=null){
        cell = row.createCell(20);
    }
        cell.setCellType(Cell.CELL_TYPE_STRING);
        cell.setCellValue("单元格测试");

    FileOutputStream fos = new FileOutputStream("d:\\非费用付款申请报表导出2.xls");
        wb.write(fos);
        fos.close();
    }
}
