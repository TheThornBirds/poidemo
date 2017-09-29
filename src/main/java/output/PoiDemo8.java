package output;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

/**
 * Created by Administrator on 2017/9/29.
 */
public class PoiDemo8 {
    public static void main(String[] args) throws Exception {
        Workbook wb = new HSSFWorkbook();  //����һ��Excel
        Sheet sheet = wb.createSheet("one");
        Row row = sheet.createRow(3);

        Cell cell = row.createCell(2);
        cell.setCellValue("xx");
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setFillBackgroundColor(IndexedColors.RED.getIndex()); //���ñ���ɫ
        cellStyle.setFillPattern(CellStyle.BIG_SPOTS);
        cell.setCellStyle(cellStyle);

        Cell cell2 = row.createCell(5);
        cell2.setCellValue("lala");
        CellStyle cellStyle1 = wb.createCellStyle();
        cellStyle1.setFillForegroundColor(IndexedColors.RED.getIndex()); //����ǰ��ɫ
        cellStyle1.setFillPattern(CellStyle.SOLID_FOREGROUND);
        cell2.setCellStyle(cellStyle1);

        FileOutputStream fos = new FileOutputStream("D:����ɫ.xls");
        wb.write(fos);
        fos.close();
    }
}
