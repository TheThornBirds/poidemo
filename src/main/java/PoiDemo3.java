import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.Date;

/**
 * Created by Administrator on 2017/9/27.
 */
public class PoiDemo3 {
    public static void main(String[] args) throws Exception {
        Workbook book = new HSSFWorkbook(); //创建一个工作簿
        Sheet sheet = book.createSheet();
        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue(new Date());
        row.createCell(1).setCellValue(1);
        row.createCell(2).setCellValue("一个字符串");
        row.createCell(3).setCellValue(true);
        row.createCell(4).setCellValue(HSSFCell.CELL_TYPE_NUMERIC);
        row.createCell(5).setCellValue(false);

        FileOutputStream fos = new FileOutputStream("D:\\哈哈哈哈.xls");
        book.write(fos);
        fos.close();
    }
}
