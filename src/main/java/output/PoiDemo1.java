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

        Workbook book = new HSSFWorkbook(); //创建一个新的表格
        Sheet sheet = book.createSheet("第一个sheet页"); //创建第一个sheet页
        Sheet sheet1 = book.createSheet("第二个sheet页"); //创建第二个sheet页
        Row row = sheet.createRow(0); //创建一行
        Cell cell = row.createCell(0); //给那一行创建一个单元格
        cell.setCellValue(1); //给单元格赋值
        row.createCell(1).setCellValue("中文"); //创建第二个单元格并赋值
        row.createCell(2).setCellValue("english"); //创建第三个单元格并赋值
        row.createCell(3).setCellValue(false); //创建第四个单元格并赋值
        FileOutputStream fos = new FileOutputStream("D:\\poiDemo1.xls");
        book.write(fos);
        fos.close();
}
}
