package newproject;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

public class PG1 {

    private static final String GECKO_DRIVER_PATH = "C:\\geckodriver.exe";
    private static final String BASE_URL = "http://demo.guru99.com/test/newtours/";
    private static final String EXPECTED_TITLE = "Welcome: Mercury Tours";

    public static void main(String[] args) {
        // Setup WebDriver and run the test
        setupAndRunTest();
    }

    private static void setupAndRunTest() {
        // Set system property for GeckoDriver
        System.setProperty("webdriver.gecko.driver", GECKO_DRIVER_PATH);

        // Initialize WebDriver
        WebDriver driver = new FirefoxDriver();

        // Uncomment below lines to use Chrome instead of Firefox
        // System.setProperty("webdriver.chrome.driver", "C:\\chromedriver.exe");
        // WebDriver driver = new ChromeDriver();

        try {
            // Launch Firefox and direct it to the Base URL
            driver.get(BASE_URL);

            // Get the actual title of the page
            String actualTitle = driver.getTitle();

            // Compare the actual title with the expected title
            if (actualTitle.equals(EXPECTED_TITLE)) {
                System.out.println("Test Passed!");
            } else {
                System.out.println("Test Failed");
            }
        } catch (Exception e) {
            System.out.println("An error occurred: " + e.getMessage());
        } finally {
            // Close Firefox
            driver.close();
        }
    }
}
