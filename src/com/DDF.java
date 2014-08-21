package com;


import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import java.util.concurrent.TimeUnit;
import com.ExcelMethods;

import static com.ExcelMethods.writeExcel;

/**
 * Created by Maruf on 8/19/2014.
 */
public class DDF {
    // write code to search amazon
    static WebDriver driver;
   static String tdID, dept, searchTerm, result, actual;
    static String[][] testData;

    @BeforeMethod
    public void setUp() throws InterruptedException {
        driver = new ChromeDriver();
        driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
        driver.get("http://www.amazon.com");
        Thread.sleep(1000);

    }
    @Test
    public void dataDF() throws Exception {
        // get data from excel
        String path="C:\\Users\\Maruf\\DataDriven\\src\\com\\test_data\\KDF_TestData.xlsx";
        String sheetName = "Test Data";
        testData = ExcelMethods.readXL(path, sheetName);
        for (int i = 1; i < testData.length; i++) {
            tdID=testData[i][0];
            dept =testData[i][1];
            searchTerm = testData[i][2];
            testData[i][3]=result;
            testData[i][4]=actual;
            // test using excel file
            searchAmazon();
                        Thread.sleep(2000);
            // get results
             actual=driver.findElement(By.cssSelector(".nav-subnav-item")).getText();
//            System.out.println("actual = " + actual);
            if (actual.equalsIgnoreCase(dept)) {
                result = "pass";

            } else {
                result="fail";
                System.out.println("Failed at "+tdID);
                System.out.println("expected = "+ dept);
                System.out.println("actual  =  " + actual);
            }
        }
    }

    @AfterMethod
    public void teardown() throws Exception {
        driver.quit();
        writeExcel("C:\\Users\\Maruf\\DataDriven\\src\\com\\test_data\\KDF_TestResults.xlsx", "Results", testData);
    }



    public static void searchAmazon() throws InterruptedException {
        // go to amazon.com

        // select department from dropdown menu
        driver.findElement(By.id("searchDropdownBox")).click();
        Select select = new Select(driver.findElement(By.cssSelector("#searchDropdownBox")));

        select.selectByVisibleText(dept);
        // enter search term
        driver.findElement(By.cssSelector("#twotabsearchtextbox")).clear();
        driver.findElement(By.cssSelector("#twotabsearchtextbox")).sendKeys(searchTerm);
        // click search
        driver.findElement(By.cssSelector("#twotabsearchtextbox")).submit();
        // verify results

    }

}
