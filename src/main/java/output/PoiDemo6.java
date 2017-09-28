package output;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

/**
 * Created by Administrator on 2017/9/28.
 */
public class PoiDemo6 {
    public static void main(String[] args) throws Exception {
        Workbook wb = new HSSFWorkbook(); //创建一个工作簿
        Sheet sheet = wb.createSheet("sheet1"); //定义一个sheet
        Row row = sheet.createRow(2);
        row.setHeightInPoints(30);

        createCell(wb, row, (short)0, HSSFCellStyle.ALIGN_CENTER, HSSFCellStyle.VERTICAL_BOTTOM);
        createCell(wb, row, (short)1, HSSFCellStyle.ALIGN_FILL, HSSFCellStyle.VERTICAL_CENTER);
        createCell(wb, row, (short)2, HSSFCellStyle.ALIGN_LEFT, HSSFCellStyle.VERTICAL_TOP);
        createCell(wb, row, (short)3, HSSFCellStyle.ALIGN_RIGHT, HSSFCellStyle.VERTICAL_TOP);

        FileOutputStream fos = new FileOutputStream("D:\\工作表.xls");
        wb.write(fos);
        fos.close();
    }


    private static void createCell(Workbook wb, Row row, short column, short halign, short valign){
        Cell cell = row.createCell(column);
        cell.setCellValue(new HSSFRichTextString("Align It")); //设置值
        CellStyle cellStyle = wb.createCellStyle(); //创建单元格样式
        cellStyle.setAlignment(halign); // 设置单元格水平方向对其方式
        cellStyle.setVerticalAlignment(valign); // 设置单元格垂直方向对其方式
        cell.setCellStyle(cellStyle); // 设置单元格样式
    }
}
