package newproject;

import excelExportAndFileIO.WriteExcelFile;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

public class CheckCity {

    private static WebDriver driver;

    public static void main(String[] args) {
        try {
            // Set up the WebDriver
            setupDriver();
            // Open Google Maps
            driver.get("https://maps.google.com/");
            // Start the process of reading from the Excel file and checking city/state
            getState();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            // Ensure the WebDriver is properly closed
            if (driver != null) {
                driver.quit();
            }
        }
    }

    // Method to set up the WebDriver with Chrome options
    private static void setupDriver() {
        System.setProperty("webdriver.chrome.driver", "C:\\chromedriver.exe");
        ChromeOptions options = new ChromeOptions();
        options.addArguments("headless");
        driver = new ChromeDriver(options);
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
    }

    // Method to read data from an Excel file
    private void readExcel(String filePath, String fileName, String sheetName) throws IOException {
        File file = new File(filePath + "\\" + fileName);
        FileInputStream inputStream = new FileInputStream(file);

        Workbook workbook = null;
        String fileExtensionName = fileName.substring(fileName.indexOf("."));
        if (fileExtensionName.equals(".xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        } else if (fileExtensionName.equals(".xls")) {
            workbook = new HSSFWorkbook(inputStream);
        }

        // If the workbook was successfully created, process the sheet
        if (workbook != null) {
            Sheet sheet = workbook.getSheet(sheetName);
            int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();

            // Iterate through the rows of the sheet
            for (int i = 0; i <= rowCount; i++) {
                Row row = sheet.getRow(i);
                String state = row.getCell(0).getStringCellValue();
                String city = row.getCell(1).getStringCellValue();
                searchCity(city, state);
            }

            workbook.close();
        }
        inputStream.close();
    }

    // Method to search for a city and state on Google Maps
    private void searchCity(String city, String state) throws InterruptedException {
        driver.findElement(By.xpath("//input[@id='searchboxinput']")).click();
        driver.findElement(By.xpath("//input[@id='searchboxinput']")).sendKeys(city + " " + state);
        driver.findElement(By.xpath("//input[@id='searchboxinput']")).sendKeys(Keys.ENTER);
        Thread.sleep(2000); // Wait for the search results to load

        try {
            String actualState = driver.findElement(By.xpath("/html[1]/body[1]/div[3]/div[9]/div[8]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/h2[2]/span[1]")).getText();
            String actualCity = driver.findElement(By.xpath("//*[@id=\"pane\"]/div/div[1]/div/div/div[2]/div[1]/div[1]/div[1]/h1/span[1]")).getText();
            validateCityAndState(city, state, actualCity, actualState, null);
        } catch (Exception e) {
            handleSearchException(city, state);
        }
    }

    // Method to validate the city and state obtained from Google Maps against the expected values
    private void validateCityAndState(String city, String state, String actualCity, String actualState, String pinCode) {
        if (actualState != null && !actualState.contains(state)) {
            WriteExcelFile.print(state + " is different");
        } else if (pinCode != null && pinCode.contains("79") && actualCity.contains(city)) {
            WriteExcelFile.print(city + " is correct");
        } else {
            WriteExcelFile.print(city + " - City name is different");
        }
    }

    // Method to handle exceptions that occur during the search
    private void handleSearchException(String city, String state) {
        try {
            String pinCode = driver.findElement(By.xpath("//*[@id=\"pane\"]/div/div[1]/div/div/div[2]/div[1]/div[1]/h2/span")).getText();
            String actualCity = driver.findElement(By.xpath("//*[@id=\"pane\"]/div/div[1]/div/div/div[2]/div[1]/div[1]/div[1]/h1/span[1]")).getText();
            validateCityAndState(city, state, actualCity, null, pinCode);
        } catch (Exception e) {
            WriteExcelFile.print(city + " has incorrect state");
            clickClearSearchBox();
        }
    }

    // Method to clear the search box in Google Maps
    private void clickClearSearchBox() {
        try {
            driver.findElement(By.cssSelector("#sb_cb50")).click();
        } catch (Exception e) {
            driver.findElement(By.xpath("//*[@id=\"omnibox-directions\"]/div/div[2]/div/button/div")).click();
        }
    }

    // Method to initiate the process of reading from the Excel file and validating cities/states
    public static void getState() throws Exception {
        CheckCity objExcelFile = new CheckCity();
        String filePath = "C:\\Users\\Prefme_Matrix\\OneDrive\\Documents";
        objExcelFile.readExcel(filePath, "ExportExcel.xlsx", "Sheet1");
    }
}
