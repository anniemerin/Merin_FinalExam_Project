package MerinFinalProject;

import org.testng.annotations.Test;

import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeSuite;
import com.aventstack.extentreports.Status;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterMethod;

public class FinalTest {
	
	WebDriver driver;
    XSSFWorkbook workbook = null;
    static Logger logger = Logger.getLogger(FinalTest.class);
    
    @Test
    public void readData() throws InterruptedException, IOException {

        WebElement tableElement = driver.findElement(By.xpath("//table[contains(@summary,'Bonds')]"));
        List<WebElement> tableBodyElement = tableElement.findElements(By.tagName("tbody"));

        // Code for Extent Report
        String methodName = new Exception().getStackTrace()[0].getMethodName();
        ExtentReportClass.test = ExtentReportClass.extent.createTest(methodName, "Validate Results for the last 90 days - Bonds");

        for (int tb = 1; tb < 6; tb++) {
            WebElement tbody = tableBodyElement.get(tb);
            List<WebElement> rowElement = tbody.findElements(By.tagName("tr"));

            for (int r = 0; r < rowElement.size() - 2; r++) {

                WebElement row = rowElement.get(r);
                String bondName = row.findElement(By.tagName("a")).getText();
                row.findElement(By.tagName("a")).click();
                
                logger.info(bondName+" Bond is selected");
             
                //ExtentReport code
                ExtentReportClass.test.log(Status.PASS, bondName + " is clicked successfully");
                ExtentReportClass.test.addScreenCaptureFromPath(ExtentReportClass.captureScreenshot(driver));

                Thread.sleep(2000); 
                
                logger.info(bondName+" Bond Report is Opened");

                System.out.println("Table "+tb+" Row "+(r+1)+" : " + bondName);
                driver.switchTo().frame(0); 
                WebElement scrollElement = driver.findElement(By.xpath("//strong[contains(text(),'Weighted average maturity based on nominal value of the bonds')]"));
                JavascriptExecutor js = (JavascriptExecutor) driver;
                js.executeScript("arguments[0].scrollIntoView();",scrollElement);
                
                //ExtentReport code
                ExtentReportClass.test.log(Status.INFO, "Opening report for bond: " + bondName);
                ExtentReportClass.test.log(Status.PASS, "Report is opened and validated successfully");
                ExtentReportClass.test.addScreenCaptureFromPath(ExtentReportClass.captureScreenshot(driver));

                WebElement reportTableElement = driver.findElement(By.xpath("//table[contains(@width,'500')]"));
                List<WebElement> reportRowElement = reportTableElement.findElements(By.tagName("tr"));

                for (int rr = 0; rr < reportRowElement.size(); rr++) {
                    WebElement reportRow = reportRowElement.get(rr);
                    List<WebElement> reportColumnElement = reportRow.findElements(By.tagName("td"));

                    for (int rc = 0; rc < reportColumnElement.size(); rc++) {
                        WebElement reportColumn = reportColumnElement.get(rc);
                        String value = reportColumn.getText();
                        writeIntoExcel(value, bondName, rr, rc);
                    }
                }
                
                logger.info(bondName+" Bond Report data is downloaded to an excel sheet");
                
                driver.switchTo().defaultContent();
                driver.findElement(By.xpath("//button[contains(@class,'close')]")).click();
                
                logger.info(bondName+" Bond Report is Closed");
          //      Thread.sleep(2000); 
            }
        }
    }

    public void writeIntoExcel(String value, String bondName, int rr, int rc) throws IOException {

        String path = "C:\\Users\\anupk\\OneDrive\\Desktop\\Automation Learning\\BondReports.xlsx";
        
        String sheetName=bondName ; 
		 if (bondName.length()>31)
		 {
			 sheetName=bondName.substring(0, 31);
		 }

        try {
            if (workbook == null) {
                workbook = new XSSFWorkbook();
            }

            XSSFSheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                sheet = workbook.createSheet(sheetName);
            }

            XSSFRow row = sheet.getRow(rr);
            if (row == null) {
                row = sheet.createRow(rr);
            }

            row.createCell(rc).setCellValue(value);

            FileOutputStream outputStream = new FileOutputStream(path);
            workbook.write(outputStream);
            outputStream.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @BeforeMethod
    public void beforeMethod() throws InterruptedException {
    	
    	ExtentReportClass.initializer();
    	PropertyConfigurator.configure("resources\\log4j.properties");

        driver = new ChromeDriver();
        driver.get("https://www.finmun.finances.gouv.qc.ca/finmun/f?p=100:3000::RESLT::::");
        driver.manage().window().maximize();
        logger.info("Quebec Auction and publication of results of debt securities issued for municipal financing purposes Website is opened successfully");

        driver.findElement(By.xpath("//a[contains(text(),'English')]")).click();
        logger.info("Website is translated to English language");
        Thread.sleep(2000); 
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
    }

    @AfterMethod
    public void afterMethod() throws IOException {
    	
    	workbook.close();
        ExtentReportClass.extent.flush();
        driver.quit();
    }

    @BeforeSuite
    public void beforeSuite() {
      
    }
}