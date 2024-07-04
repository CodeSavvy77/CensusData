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

public class CheckCity2 {

    private static WebDriver driver;
    public String ActualState;
    public String ActualCity;
    public String State;
    public String City;
    public String QuickFacts = "";

    public static void main(String[] args) throws Exception {
        // Set up WebDriver and initialize Chrome in headless mode
        setupDriver();

        // Set the base URL for Google Maps
        String baseUrl = "https://maps.google.com/";
        driver.get(baseUrl);

        // Start the process of reading from the Excel file and checking city/state
        getState();

        // Quit the WebDriver
        driver.quit();
    }

    // Method to set up WebDriver
    private static void setupDriver() {
        System.setProperty("webdriver.chrome.driver", "C:\\chromedriver.exe");
        ChromeOptions options = new ChromeOptions();
        options.addArguments("headless");
        driver = new ChromeDriver(options);
        driver.manage().window().maximize();
    }

    // Method to read data from an Excel file
    public void readExcel(String filePath, String fileName, String sheetName) throws Exception {
        File file = new File(filePath + "\\" + fileName);
        FileInputStream inputStream = new FileInputStream(file);

        Workbook workbook = null;
        String fileExtensionName = fileName.substring(fileName.indexOf("."));

        // Check if the file is xlsx or xls and create the respective Workbook object
        if (fileExtensionName.equals(".xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        } else if (fileExtensionName.equals(".xls")) {
            workbook = new HSSFWorkbook(inputStream);
        }

        Sheet sheet = workbook.getSheet(sheetName);
        int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();

        // Loop over all the rows of the Excel file to read it
        for (int i = 0; i <= rowCount; i++) {
            Row row = sheet.getRow(i);
            State = row.getCell(0).getStringCellValue();
            City = row.getCell(1).getStringCellValue();

            // Perform the city and state search on Google Maps
            searchCityAndState(City, State);
        }
    }

    // Method to search for a city and state on Google Maps
    public void searchCityAndState(String city, String state) throws InterruptedException {
        driver.findElement(By.xpath("//input[@id='searchboxinput']")).click();
        driver.findElement(By.xpath("//input[@id='searchboxinput']")).sendKeys(city + " " + state);
        driver.findElement(By.xpath("//input[@id='searchboxinput']")).sendKeys(Keys.ENTER);
        Thread.sleep(2000); // Wait for the search results to load

        try {
            ActualState = driver.findElement(By.xpath("/html[1]/body[1]/div[3]/div[9]/div[8]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/h2[2]/span[1]")).getText();
            ActualCity = driver.findElement(By.xpath("//*[@id=\"pane\"]/div/div[1]/div/div/div[2]/div[1]/div[1]/div[1]/h1/span[1]")).getText();
        } catch (Exception e) {
            handleSearchException(city);
            return;
        }

        // Validate the actual state and city
        validateStateAndCity(city, state);
    }

    // Method to handle exceptions that occur during the search
    private void handleSearchException(String city) throws InterruptedException {
        try {
            ActualState = driver.findElement(By.xpath("//*[@id=\"pane\"]/div/div[1]/div/div/div[2]/div[1]/div[1]/h2/span")).getText();
        } catch (Exception exception) {
            WriteExcelFile.print("check " + city);
            closeSearchBox();
            return;
        }
        ActualCity = driver.findElement(By.xpath("//*[@id=\"pane\"]/div/div[1]/div/div/div[2]/div[1]/div[1]/div[1]/h1/span[1]")).getText();

        if (!ActualState.contains(State) && !ActualState.contains("14")) {
            WriteExcelFile.print("check " + city);
        } else {
            WriteExcelFile.print(city + " is correct");
        }
        closeSearchBox();
    }

    // Method to validate the state and city
    private void validateStateAndCity(String city, String state) throws InterruptedException {
        if (!ActualState.contains(state)) {
            WriteExcelFile.print("State is different");
        } else if (ActualState.contains("14") || ActualState.contains(state)) {
            if (ActualCity.contains(city)) {
                WriteExcelFile.print(city + " is correct");
            } else {
                try {
                    QuickFacts = driver.findElement(By.xpath("//*[@id=\"pane\"]/div/div[1]/div/div/div[7]/div[1]/span/span[1]")).getText();
                } catch (Exception e) {
                    closeSearchBox();
                }
                if (QuickFacts.contains(city)) {
                    WriteExcelFile.print(city + " is correct");
                } else {
                    WriteExcelFile.print(city + " - City name is different");
                }
            }
        } else {
            WriteExcelFile.print(city + " has incorrect state");
        }
        closeSearchBox();
    }

    // Method to close the search box in Google Maps
    private void closeSearchBox() {
        try {
            driver.findElement(By.cssSelector("#sb_cb50")).click();
        } catch (Exception e) {
            driver.findElement(By.xpath("//*[@id=\"omnibox-directions\"]/div/div[2]/div/button/div")).click();
        }
    }

    // Method to initiate the process of reading from the Excel file and validating cities/states
    public static void getState() throws Exception {
        CheckCity2 objExcelFile = new CheckCity2();
        String filePath = "C:\\Users\\Prefme_Matrix\\IdeaProjects\\untitled\\src\\main\\java\\excelExportAndFileIO";
        objExcelFile.readExcel(filePath, "ImportExcel.xlsx", "Sheet1");
    }
}
