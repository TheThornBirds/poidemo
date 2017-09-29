package output;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

/**
 * Created by Administrator on 2017/9/29.
 */
public class PoiDemo9 {
    public static void main(String[] args) throws Exception {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("one");
        Row row = sheet.createRow(1);

        Cell cell = row.createCell(1);
        cell.setCellValue("��Ԫ��ϲ�����");

        sheet.addMergedRegion(new CellRangeAddress(
                1, //��ʼ��
                2, //������
                1, //��ʼ��
                2  //������
        ));

        FileOutputStream fileOut=new FileOutputStream("D:\\������.xls");
        wb.write(fileOut);
        fileOut.close();
    }
}
