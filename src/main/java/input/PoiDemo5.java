package input;

import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;

/**
 * Created by Administrator on 2017/9/27.
 */
public class PoiDemo5 {
    public static void main(String[] args) throws Exception {
        InputStream is = new FileInputStream("D://�Ƿ��ø������뱨����.xls");
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook hwb = new HSSFWorkbook(fs);
        ExcelExtractor excelExtractor = new ExcelExtractor(hwb);
        excelExtractor.setIncludeSheetNames(true); //���ǲ���Ҫsheetҳ������
        System.out.println(excelExtractor.getText());
    }
}
