package test;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hamcrest.Matchers;
import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

public class MoneyRediffStock {
   WebDriver driver;

    @BeforeMethod
    public void setUp() {
        WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver();
        driver.get("https://money.rediff.com/losers/bse/daily/groupall");
        driver.manage().window().maximize();
    }

    @AfterMethod
    public void tearDown() {
        if (driver != null) {
            driver.quit();
        }
    }

    @Test
    public void verifyDataFromTable() throws IOException {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(2));
        WebElement table = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//table[@class =\"dataTable\"]")));

        List<WebElement> rows = table.findElements(By.tagName("tr"));
        Map<String, String> tableData = new TreeMap<>();
        for (int i = 1; i < rows.size(); i++) {
            List<WebElement> cells = rows.get(i).findElements(By.tagName("td"));
            String key = cells.get(0).getText();
            String value = cells.get(3).getText();
            tableData.put(key, value);
        }
        String actualOutput = null;
        for (Map.Entry<String, String> entry : tableData.entrySet()) {
            actualOutput = entry.getKey() + " " + entry.getValue();
            System.out.println(actualOutput);
        }

        String excelFilePath = "src/test/java/Excel/MoneyRediffStockSheet.xlsx";
        Workbook workbook = null;
        if (excelFilePath.endsWith(".xlsx")) {
            workbook = new XSSFWorkbook(new FileInputStream(excelFilePath));
        } else if (excelFilePath.endsWith(".xls")) {
            workbook = new HSSFWorkbook(new FileInputStream(excelFilePath));
        } else {
            throw new IOException("Unsupported file format. Please provide an XLS or XLSX file.");
        }
        Sheet sheet = workbook.getSheetAt(0);
        String excelSheetValue = null;
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            for (int j = 0; j < row.getLastCellNum(); j++) {
                Cell cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                excelSheetValue = cell.toString();
                System.out.print(excelSheetValue + " ");
            }
            System.out.println();
        }
        workbook.close();

        Assert.assertEquals(actualOutput, Matchers.containsString(excelSheetValue));

    }
}


