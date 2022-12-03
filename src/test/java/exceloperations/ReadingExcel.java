package exceloperations;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ReadingExcel {

    public static void main(String[] args) throws IOException {
        //type your file path
        String excelFilePath = "/Users/emre/IdeaProjects/ApachePOI/dataFiles/countries.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);

        XSSFWorkbook workbook = new XSSFWorkbook(inputStream); //represantive workbook
        XSSFSheet sheet = workbook.getSheetAt(0); //represantive sheet
        //XSSFSheet sheet = workbook.getSheet("Sheet1"); -- u can use or up

        ////USING FOR LOOP
       int rows =  sheet.getLastRowNum();
       int cols = sheet.getRow(1).getLastCellNum();

       for(int r = 0; r <=rows; r++){
           XSSFRow row = sheet.getRow(r); //represantive row

           for (int c=0; c<cols; c++){
                XSSFCell cell = row.getCell(c); //represantive cell
               switch (cell.getCellType())
               {
                   case  STRING  :System.out.println(cell.getStringCellValue());break;
                   case  NUMERIC :System.out.println(cell.getNumericCellValue());break;
                   case  BOOLEAN :System.out.println(cell.getBooleanCellValue());break;
               }
           }
           System.out.println();
       }

    }
}
