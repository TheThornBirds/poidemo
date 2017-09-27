package input;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;

/**
 * Created by Administrator on 2017/9/27.
 */
public class PoiDemo4 {
    public static void main(String[] args) throws Exception {
        InputStream is = new FileInputStream("D:\\�Ƿ��ø������뱨����.xls");
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook hwb = new HSSFWorkbook(fs);
        HSSFSheet sheet = hwb.getSheetAt(0); //��ȡ��һ��sheetҳ
        if (sheet == null){
            return;
        }
        //����row
        for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++){
            HSSFRow hssfRow = sheet.getRow(rowNum);
            if (hssfRow == null){
                continue;
            }
            //����cell
            for (int cellNum = 0; cellNum <= hssfRow.getLastCellNum(); cellNum++){
                HSSFCell hssfCell = hssfRow.getCell(cellNum);
                if (hssfCell == null){
                    continue;
                }
                System.out.println(getValue(hssfCell));
            }
        }
    }

    //ͨ����ȡ���ĵ�Ԫ������ַ����ͣ�������Ӧ�Ĳ�������
    private static String getValue(HSSFCell hssfCell){
        int cellType = hssfCell.getCellType();
        if (cellType == HSSFCell.CELL_TYPE_BOOLEAN){
            return String.valueOf(hssfCell.getBooleanCellValue());
        }else if (cellType == HSSFCell.CELL_TYPE_NUMERIC){
            return String.valueOf(hssfCell.getNumericCellValue());
        }else{
            return String .valueOf(hssfCell.getStringCellValue());
        }
    }

}
