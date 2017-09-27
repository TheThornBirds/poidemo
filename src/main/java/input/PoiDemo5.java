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
        InputStream is = new FileInputStream("D://非费用付款申请报表导出.xls");
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook hwb = new HSSFWorkbook(fs);
        ExcelExtractor excelExtractor = new ExcelExtractor(hwb);
        excelExtractor.setIncludeSheetNames(true); //我们不需要sheet页的名字
        System.out.println(excelExtractor.getText());
    }
}
