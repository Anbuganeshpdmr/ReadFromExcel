package Package1;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ExcelReadClasss {

    public static void main(String args[]) throws FileNotFoundException {


    FileInputStream file = new FileInputStream("src/main/resources/ReadExcel1.xlsx");
    {
        Workbook workbook = null; // Load the workbook
        try
        {
            workbook = new XSSFWorkbook(file);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        int columnVal = 2;

        Sheet sheet = workbook.getSheet("data_AddPublisherFeature_Books"); // Assuming you want to read the first sheet
        //readSheet(sheet);


        for (Row row : sheet) {

            Cell cell = row.getCell(columnVal);
            if (cell != null) {
                // Access cell data based on cell type
                switch (cell.getCellType()) {
                    case STRING:
                        String stringValue = cell.getStringCellValue();
                        System.out.println(stringValue);
                        break;
                    case NUMERIC:
                        double numericValue = cell.getNumericCellValue();
                        System.out.println(numericValue);
                        break;
                    case BOOLEAN:
                        boolean booleanValue = cell.getBooleanCellValue();
                        System.out.println(booleanValue);
                        break;
                    // Handle other cell types as needed
                    default:
                        System.out.println();
                }
            } else {
                System.out.println(); // Empty cell
            }
        }

        try {
            workbook.close(); // Close the workbook
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }



    }
}
