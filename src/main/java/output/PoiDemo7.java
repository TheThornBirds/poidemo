package output;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

/**
 * Created by Administrator on 2017/9/28.
 */
public class PoiDemo7 {
    public static void main(String[] args) throws Exception {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("one");
        Row row = sheet.createRow(1);

        Cell cell = row.createCell(1);
        cell.setCellValue(4);

        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setBorderBottom(CellStyle.BORDER_THIN); //设置底部边框
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex()); //设置底部框颜色为黑色

        cellStyle.setBorderLeft(CellStyle.BORDER_THIN); //设置左边边框
        cellStyle.setLeftBorderColor(IndexedColors.RED.getIndex()); //设置左边边框颜色为绿色

        cellStyle.setBorderRight(CellStyle.BORDER_THIN); //设置右边边框
        cellStyle.setRightBorderColor(IndexedColors.RED.getIndex()); //设置右边框颜色为红色

        cellStyle.setBorderTop(CellStyle.BORDER_MEDIUM_DASHED); //设置上边框为虚线
        cellStyle.setTopBorderColor(IndexedColors.RED.getIndex()); //设置上边框颜色为红色

        cell.setCellStyle(cellStyle);

        FileOutputStream fos = new FileOutputStream("D:\\颜色.xls");
        wb.write(fos);
        fos.close();
    }
}
