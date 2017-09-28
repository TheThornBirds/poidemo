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
        cellStyle.setBorderBottom(CellStyle.BORDER_THIN); //���õײ��߿�
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex()); //���õײ�����ɫΪ��ɫ

        cellStyle.setBorderLeft(CellStyle.BORDER_THIN); //������߱߿�
        cellStyle.setLeftBorderColor(IndexedColors.RED.getIndex()); //������߱߿���ɫΪ��ɫ

        cellStyle.setBorderRight(CellStyle.BORDER_THIN); //�����ұ߱߿�
        cellStyle.setRightBorderColor(IndexedColors.RED.getIndex()); //�����ұ߿���ɫΪ��ɫ

        cellStyle.setBorderTop(CellStyle.BORDER_MEDIUM_DASHED); //�����ϱ߿�Ϊ����
        cellStyle.setTopBorderColor(IndexedColors.RED.getIndex()); //�����ϱ߿���ɫΪ��ɫ

        cell.setCellStyle(cellStyle);

        FileOutputStream fos = new FileOutputStream("D:\\��ɫ.xls");
        wb.write(fos);
        fos.close();
    }
}
