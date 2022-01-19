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
//    public String PinCode;
    public String State;
    public String City;
    public String QuickFacts = "";

    public static void main(String[] args) throws Exception {
        System.setProperty("webdriver.chrome.driver", "C:\\chromedriver.exe");

//        create object of chrome options
        ChromeOptions options = new ChromeOptions();

//        add the headless argument
        options.addArguments("headless");

//        pass the options parameter in the Chrome driver declaration
        driver = new ChromeDriver(options);

        //set the path of the Gecko driver as per the location on local machine
//        System.setProperty("webdriver.gecko.driver", "C:\\Users\\Prefme_Matrix\\OneDrive\\Documents\\geckodriver.exe");
//
//        //Set Firefox Headless mode as TRUE
//        FirefoxOptions options = new FirefoxOptions();
//        options.setHeadless(true);
//
//        //pass the options parameter in the Firefox driver declaration
//        driver = new FirefoxDriver(options);


//        driver = new ChromeDriver();

        driver.manage().window().maximize();
//        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

        String baseUrl = "https://maps.google.com/";
        driver.get(baseUrl);

        getState();

        driver.quit();
    }

    public void readExcel(String filePath, String fileName, String sheetName) throws Exception {

        //Create an object of File class to open xlsx file

        File file = new File(filePath + "\\" + fileName);

        //Create an object of FileInputStream class to read excel file

        FileInputStream inputStream = new FileInputStream(file);

        Workbook Workbook = null;

        //Find the file extension by splitting file name in substring  and getting only extension name

        String fileExtensionName = fileName.substring(fileName.indexOf("."));

        //Check condition if the file is xlsx file

        if (fileExtensionName.equals(".xlsx")) {

            //If it is xlsx file then create object of XSSFWorkbook class

            Workbook = new XSSFWorkbook(inputStream);

        }

        //Check condition if the file is xls file
        else if (fileExtensionName.equals(".xls")) {

            //If it is xls file then create object of HSSFWorkbook class

            Workbook = new HSSFWorkbook(inputStream);

        }

        //Read sheet inside the workbook by its name

        Sheet Sheet = Workbook.getSheet(sheetName);

        //Find number of rows in excel file

        int rowCount = Sheet.getLastRowNum() - Sheet.getFirstRowNum();

        //Create a loop over all the rows of excel file to read it

        for (int i = 0; i < rowCount + 1; i++) {

            Row row = Sheet.getRow(i);

            //Create a loop to print cell values in a row


            //Print Excel data in console
            int j = 0;

            State = row.getCell(j).getStringCellValue();
            City = row.getCell(j + 1).getStringCellValue();

            driver.findElement(By.xpath("//input[@id='searchboxinput']")).click();
            driver.findElement(By.xpath("//input[@id='searchboxinput']")).sendKeys(City + " " + State);
            driver.findElement(By.xpath("//input[@id='searchboxinput']")).sendKeys(Keys.ENTER);
            Thread.sleep(2000);


            try {
                ActualState = driver.findElement(By.xpath("/html[1]/body[1]/div[3]/div[9]/div[8]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/h2[2]/span[1]")).getText();
//                System.out.println(ActualState);
                ActualCity = driver.findElement(By.xpath("//*[@id=\"pane\"]/div/div[1]/div/div/div[2]/div[1]/div[1]/div[1]/h1/span[1]")).getText();
//                System.out.println(ActualCity);
            } catch (Exception e) {
                try {
                    ActualState = driver.findElement(By.xpath("//*[@id=\"pane\"]/div/div[1]/div/div/div[2]/div[1]/div[1]/h2/span")).getText();
                } catch (Exception exception) {
                    WriteExcelFile.print("check " + City);
//                    try {
//                        driver.findElement(By.cssSelector("#sb_cb50")).click();
//                    } catch (Exception exception1) {
//                        driver.findElement(By.xpath("//*[@id=\"omnibox-directions\"]/div/div[2]/div/button/div")).click();
//                    }
//
//                    //Search for city name and state
//                    driver.findElement(By.xpath("//input[@id='searchboxinput']")).click();
//                    driver.findElement(By.xpath("//input[@id='searchboxinput']")).sendKeys(City + " " + State);
//
//                    //Getting all results for city search
//                    List<WebElement> listOfElements = driver.findElements(By.className("ZHeE1b ZHeE1b-tPcied-dkl3Ye"));
//                    for (WebElement k: listOfElements) {
//                        String StateName = k.getText();
//                        if (StateName.contains(State)){
//                            k.click();
//                            ActualCity = driver.findElement(By.xpath("//*[@id=\"pane\"]/div/div[1]/div/div/div[2]/div[1]/div[1]/div[1]/h1/span[1]")).getText();
//                            if (ActualCity.contains(City)) {
////                    System.out.println(City + " is correct");
//                                WriteExcelFile.print(City + " is correct");
//                            } else {
//
//                            }
//                        }
//                    }

//                    new WebDriverWait(driver, 20).until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"omnibox-directions\"]/div/div[2]/div/button/div"))).click();
//                    driver.findElement(By.xpath("//*[@id=\"omnibox-directions\"]/div/div[2]/div/button/div")).click();
                    try {
                        driver.findElement(By.xpath("//*[@id=\"searchbox\"]/a")).click();
                    } catch (Exception exception1) {
                        driver.findElement(By.xpath("//div[@guidedhelpid='close_directions']")).click();
                    }
                    continue;
                }
                ActualCity = driver.findElement(By.xpath("//*[@id=\"pane\"]/div/div[1]/div/div/div[2]/div[1]/div[1]/div[1]/h1/span[1]")).getText();

                if (!ActualState.contains(State) && !ActualState.contains("14") ) {
                    WriteExcelFile.print("check " + City);
                } else {
                    WriteExcelFile.print(City + " is correct");
                }
                try {
                    driver.findElement(By.cssSelector("#sb_cb50")).click();
                } catch (Exception exception1) {
                    driver.findElement(By.xpath("//*[@id=\"omnibox-directions\"]/div/div[2]/div/button/div")).click();
                    continue;
                }
                continue;
            }
            if (!ActualState.contains(State)) {
//                System.out.println("Expected State: " + State);
//                System.out.println(ActualState);
                WriteExcelFile.print("State is different");
//                WriteExcelFile.print(State + " is different");
            } else if (ActualState.contains("14") || ActualState.contains(State)) {
                if (ActualCity.contains(City)) {
//                    System.out.println(City + " is correct");
                    WriteExcelFile.print(City + " is correct");
                } else {
                    try {
                        QuickFacts = driver.findElement(By.xpath("//*[@id=\"pane\"]/div/div[1]/div/div/div[7]/div[1]/span/span[1]")).getText();
                    } catch (Exception e) {
                        driver.findElement(By.xpath("//*[@id=\"sb_cb50\"]")).click();
                    }
                    if (QuickFacts.contains(City)) {
                            WriteExcelFile.print(City + " is correct");
                        } else {
//                    System.out.println(City + " - City name is different");
                        WriteExcelFile.print(City + " - City name is different");
                    }
                }
            } else {
//                System.out.println(City + " has incorrect state");
                WriteExcelFile.print(City + " has incorrect state");
                driver.findElement(By.cssSelector("#sb_cb50")).click();
            }
            try {
                driver.findElement(By.cssSelector("#sb_cb50")).click();
            } catch (Exception e) {
            }
        }

        driver.findElement(By.cssSelector("#sb_cb50")).click();
    }


    //Main function is calling readExcel function to read data from excel file

    public static void getState() throws Exception {

        //Create an object of ReadExcelFile class

        CheckCity2 objExcelFile = new CheckCity2();

        //Prepare the path of excel file

        String filePath = "C:\\Users\\Prefme_Matrix\\IdeaProjects\\untitled\\src\\main\\java\\excelExportAndFileIO";

        //Call read file method of the class to read data

        objExcelFile.readExcel(filePath, "ImportExcel.xlsx", "Sheet1");

    }

}