package exceloperations;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class ReadingExcel {

    public static void main(String[] args) throws IOException {
        //type your file path
        String excelFilePath = "/Users/emre/IdeaProjects/ApachePOI/dataFiles/countries.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);

        XSSFWorkbook workbook = new XSSFWorkbook(inputStream); //represantive workbook
        XSSFSheet sheet = workbook.getSheetAt(0); //represantive sheet
        //XSSFSheet sheet = workbook.getSheet("Sheet1"); -- u can use or up

        ////1. Way USING FOR LOOP - read data from excel sheet
//       int rows =  sheet.getLastRowNum();
//       int cols = sheet.getRow(1).getLastCellNum();
//
//       for(int r = 0; r <=rows; r++){
//           XSSFRow row = sheet.getRow(r); //represantive row
//
//           for (int c=0; c<cols; c++){
//                XSSFCell cell = row.getCell(c); //represantive cell
//               switch (cell.getCellType())
//               {
//                   case  STRING  :System.out.print(cell.getStringCellValue());break;
//                   case  NUMERIC :System.out.print(cell.getNumericCellValue());break;
//                   case  BOOLEAN :System.out.print(cell.getBooleanCellValue());break;
//               }
//               System.out.print(" | ");
//           }
//           System.out.println();
//       }


        ///////ITERATOR ///////

        Iterator iterator = sheet.iterator();

        while (iterator.hasNext()) {

            XSSFRow row = (XSSFRow) iterator.next();

           Iterator cellIterator =  row.cellIterator();

           while (cellIterator.hasNext()){

               XSSFCell cell = (XSSFCell) cellIterator.next();
               switch (cell.getCellType())
               {
                   case  STRING  :System.out.print(cell.getStringCellValue());break;
                   case  NUMERIC :System.out.print(cell.getNumericCellValue());break;
                   case  BOOLEAN :System.out.print(cell.getBooleanCellValue());break;
               }
               System.out.print(" | ");
           }
            System.out.println();
        }

    }
}
