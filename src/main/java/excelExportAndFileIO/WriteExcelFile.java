package excelExportAndFileIO;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteExcelFile {

    /**
     * This method is used to write data to an Excel file.
     *
     * @param filePath    The file path where the Excel file is located.
     * @param fileName    The name of the Excel file.
     * @param sheetName   The name of the sheet where data needs to be written.
     * @param dataToWrite The data to be written to the Excel file.
     * @throws IOException If there is an I/O error.
     */
    public static void writeExcel(String filePath, String fileName, String sheetName, String dataToWrite) throws IOException {

        // Create an object of File class to open xlsx/xls file
        File file = new File(filePath + "\\" + fileName);

        // Create an object of FileInputStream class to read excel file
        FileInputStream inputStream = new FileInputStream(file);
        Workbook workbook = null;

        // Find the file extension by splitting file name in substring and getting only extension name
        String fileExtensionName = fileName.substring(fileName.indexOf("."));

        // Check if the file is xlsx or xls and create appropriate workbook object
        if (fileExtensionName.equals(".xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        } else if (fileExtensionName.equals(".xls")) {
            workbook = new HSSFWorkbook(inputStream);
        }

        // Read excel sheet by sheet name
        Sheet sheet = workbook.getSheet(sheetName);

        // Get the current count of rows in excel file
        int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();

        // Get the first row from the sheet
        Row row = sheet.getRow(0);

        // Create a new row and append it at the end of the sheet
        Row newRow = sheet.createRow(rowCount + 1);

        // Fill data in the new row
        Cell cell = newRow.createCell(0);
        cell.setCellValue(dataToWrite);

        // Close input stream
        inputStream.close();

        // Create an object of FileOutputStream class to write data in excel file
        FileOutputStream outputStream = new FileOutputStream(file);

        // Write data in the excel file
        workbook.write(outputStream);

        // Close output stream
        outputStream.close();
    }

    /**
     * This method is used to print a string and write it to an Excel file.
     *
     * @param s The string to be printed and written to the Excel file.
     * @throws IOException If there is an I/O error.
     */
    public static void print(String s) throws IOException {
        // Print the string to the console
        System.out.println(s);

        // Write the string to the Excel file
        writeExcel("C:\\Users\\Prefme_Matrix\\IdeaProjects\\untitled\\src\\main\\java\\excelExportAndFileIO", "ExportExcel.xlsx", "Sheet1", s);
    }
}
