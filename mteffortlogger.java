package excelWriterexample;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class mteffortlogger {



     public static void main1(String[] args) {
            FileOutputStream looger = null;
            try {
                looger = new FileOutputStream(new File("Copy of EL.xlsx"));
                
                String SAMPLE_XLSX_FILE_PATH = "170046_B_AppDev_W10_EffortLogger.xlsm";
                Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));
                workbook.write(looger);
                workbook.close();
                looger.flush();
            }
            catch (Exception v) {
                v.printStackTrace();
            }
     }
}


