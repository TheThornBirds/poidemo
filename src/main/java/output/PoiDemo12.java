package output;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import javax.swing.text.Style;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by Administrator on 2017/9/29.
 */
public class PoiDemo12 {
    public static void main(String[] args) throws IOException {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("one");
        Row row;
        Cell cell;
        CellStyle cs;
        DataFormat format = wb.createDataFormat();
        short rowNum = 0;
        short colNum = 0;

        row = sheet.createRow(rowNum++);
        cell = row.createCell(colNum);
        cell.setCellValue(111111.25);

        cs = wb.createCellStyle();
        cs.setDataFormat(format.getFormat("0.0")); //设置数据格式
        cell.setCellStyle(cs);

        row = sheet.createRow(rowNum++);
        cell = row.createCell(colNum);
        cell.setCellValue(111111.25);
        cs = wb.createCellStyle();
        cs.setDataFormat(format.getFormat("#,##0.000"));
        cell.setCellStyle(cs);

        FileOutputStream fos = new FileOutputStream("d://设置数据格式.xls");
        wb.write(fos);
        fos.close();
    }
}
