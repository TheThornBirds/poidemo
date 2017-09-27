package output;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.Date;

/**
 * Created by Administrator on 2017/9/27.
 */
public class PoiDemo2 {

    public static void main(String[] args) throws Exception {
        Workbook book = new HSSFWorkbook(); //����һ���µı��
        Sheet sheet1 = book.createSheet("��һ��sheet");
        Row row = sheet1.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue(new Date());

        CreationHelper creationHelper = book.getCreationHelper();
        CellStyle cellStyle = book.createCellStyle(); //��Ԫ����ʽ��
        cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyy-mm-dd hh:mm:ss"));
        cell = row.createCell(1); //�ڶ���
        cell.setCellValue(new Date());

        cell = row.createCell(2); //������
        cell.setCellValue("lalal");
        cell.setCellStyle(cellStyle);
        System.out.println(cell);
        FileOutputStream fos = new FileOutputStream("D://haha.xls");
        book.write(fos);
        fos.close();

    }

}
