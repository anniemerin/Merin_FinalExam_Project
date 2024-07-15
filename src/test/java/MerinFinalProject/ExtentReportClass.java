package MerinFinalProject;

import java.io.File;
import java.io.IOException;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;

public class ExtentReportClass {
	public static ExtentSparkReporter sparkReporter;
    public static ExtentReports extent;
    public static ExtentTest test;

    public static void initializer() {
        sparkReporter = new ExtentSparkReporter(System.getProperty("user.dir") + "/Reports/extentSparkReport.html");
        sparkReporter.config().setDocumentTitle("Automation Report");
        sparkReporter.config().setReportName("Test Execution Report");
        sparkReporter.config().setTheme(Theme.STANDARD);
        sparkReporter.config().setTimeStampFormat("yyyy-MM-dd HH:mm:ss");
        extent = new ExtentReports();
        extent.attachReporter(sparkReporter);
    }

    public static String captureScreenshot(WebDriver driver) throws IOException {
        String FileSeparator = System.getProperty("file.separator");
        String Extent_report_path = "." + FileSeparator + "Reports";
        File Src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
        String Screenshotname = "screenshot" + Math.random() + ".png";
        File Dst = new File(Extent_report_path + FileSeparator + "Screenshots" + FileSeparator + Screenshotname);
        FileUtils.copyFile(Src, Dst);
        String absPath = Dst.getAbsolutePath();
        return absPath;
    }
	
}



