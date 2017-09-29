package output;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

/**
 * Created by Administrator on 2017/9/29.
 */
public class PoiDemo11 {
    public static void main(String[] args) throws Exception {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("one");
        Row row = sheet.createRow(2);
        Cell cell = row.createCell(2);
        cell.setCellValue("我要换行 \n 换行成功了吗?");
        CellStyle cellStyle = wb.createCellStyle();
        //设置可以换行
        cellStyle.setWrapText(true);
        cell.setCellStyle(cellStyle);

        //调整下行的高度
        row.setHeightInPoints(2*sheet.getDefaultRowHeightInPoints());
        //调整单元格宽度
        sheet.autoSizeColumn(2);

        FileOutputStream fos = new FileOutputStream("D:\\lalala.xls");
        wb.write(fos);
        fos.close();
    }
}
