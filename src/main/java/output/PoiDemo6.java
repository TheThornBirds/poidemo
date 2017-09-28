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
        Workbook wb = new HSSFWorkbook(); //����һ��������
        Sheet sheet = wb.createSheet("sheet1"); //����һ��sheet
        Row row = sheet.createRow(2);
        row.setHeightInPoints(30);

        createCell(wb, row, (short)0, HSSFCellStyle.ALIGN_CENTER, HSSFCellStyle.VERTICAL_BOTTOM);
        createCell(wb, row, (short)1, HSSFCellStyle.ALIGN_FILL, HSSFCellStyle.VERTICAL_CENTER);
        createCell(wb, row, (short)2, HSSFCellStyle.ALIGN_LEFT, HSSFCellStyle.VERTICAL_TOP);
        createCell(wb, row, (short)3, HSSFCellStyle.ALIGN_RIGHT, HSSFCellStyle.VERTICAL_TOP);

        FileOutputStream fos = new FileOutputStream("D:\\������.xls");
        wb.write(fos);
        fos.close();
    }


    private static void createCell(Workbook wb, Row row, short column, short halign, short valign){
        Cell cell = row.createCell(column);
        cell.setCellValue(new HSSFRichTextString("Align It")); //����ֵ
        CellStyle cellStyle = wb.createCellStyle(); //������Ԫ����ʽ
        cellStyle.setAlignment(halign); // ���õ�Ԫ��ˮƽ������䷽ʽ
        cellStyle.setVerticalAlignment(valign); // ���õ�Ԫ��ֱ������䷽ʽ
        cell.setCellStyle(cellStyle); // ���õ�Ԫ����ʽ
    }
}
