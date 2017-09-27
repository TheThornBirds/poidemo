package output;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

/**
 * Created by wuchen on 2017/9/26.
 */
public class PoiDemo1 {

    public static void main(String[] args) throws Exception {

        Workbook book = new HSSFWorkbook(); //����һ���µı��
        Sheet sheet = book.createSheet("��һ��sheetҳ"); //������һ��sheetҳ
        Sheet sheet1 = book.createSheet("�ڶ���sheetҳ"); //�����ڶ���sheetҳ
        Row row = sheet.createRow(0); //����һ��
        Cell cell = row.createCell(0); //����һ�д���һ����Ԫ��
        cell.setCellValue(1); //����Ԫ��ֵ
        row.createCell(1).setCellValue("����"); //�����ڶ�����Ԫ�񲢸�ֵ
        row.createCell(2).setCellValue("english"); //������������Ԫ�񲢸�ֵ
        row.createCell(3).setCellValue(false); //�������ĸ���Ԫ�񲢸�ֵ
        FileOutputStream fos = new FileOutputStream("D:\\poiDemo1.xls");
        book.write(fos);
        fos.close();
}
}
