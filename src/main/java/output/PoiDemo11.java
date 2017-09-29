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
        cell.setCellValue("��Ҫ���� \n ���гɹ�����?");
        CellStyle cellStyle = wb.createCellStyle();
        //���ÿ��Ի���
        cellStyle.setWrapText(true);
        cell.setCellStyle(cellStyle);

        //�������еĸ߶�
        row.setHeightInPoints(2*sheet.getDefaultRowHeightInPoints());
        //������Ԫ����
        sheet.autoSizeColumn(2);

        FileOutputStream fos = new FileOutputStream("D:\\lalala.xls");
        wb.write(fos);
        fos.close();
    }
}
