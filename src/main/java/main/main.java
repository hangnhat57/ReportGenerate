package main;


import io.github.bonigarcia.wdm.ChromeDriverManager;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;



public class main {
    static WebDriver driver;
    public static final String URL = "https://tciasia.myjetbrains.com/youtrack/reports/time/112-1";
    static String Username = "//td[@class='yt-table__cell yt-table__group__title']";
    public static final String FILE = System.getProperty("user.dir") + "/report.xlsx";


    static ArrayList<Card> ArrayWorks = new ArrayList<>();

    public static void main(String[] args) {
        ArrayWorks = getArrayWorks();
        for (Card work:ArrayWorks
                ) {
            System.out.println(work.work);
        }
        XSSFWorkbook workbook = new XSSFWorkbook();
        for (Card work:ArrayWorks) {
            XSSFSheet sheet = workbook.createSheet(work.Username);
            ArrayList<Map<String,String>> working = work.work;
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("No.");
            header.createCell(1).setCellValue("Work description");
            header.createCell(2).setCellValue("Time Cost");
            header.createCell(3).setCellValue("Total Time cost: "+work.TotalTime);
            for (int i = 0; i < working.size(); i++) {
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(i+1);
                row.createCell(1).setCellValue(working.get(i).get("workname"));
                row.createCell(2).setCellValue(working.get(i).get("time"));
            }
        }
        try {
            FileOutputStream outputStream = new FileOutputStream(FILE);
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Done");


    }



    // Get all data from report to array
    public static ArrayList<Card> getArrayWorks() {
        ArrayList<Card> list = new ArrayList<>();
        driver = ChromeBrowser();
        WebDriverWait wait = new WebDriverWait(driver,60);
        driver.navigate().to(URL);
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//table")));
        List<WebElement> Cards = driver.findElements(By.xpath(Username));
        for (WebElement element:Cards){
            list.add(GetCard(element.getText()));
        }
        driver.quit();
        return list;
    }

    // Init Chrome browser
    private static WebDriver ChromeBrowser(){
        ChromeDriverManager.getInstance().setup();
        ChromeOptions option = new ChromeOptions();
        option.addArguments("--headless");
        return new ChromeDriver(option);
    }

    // Get card data from each user
    private static Card GetCard(String username){
        String Work = "//tbody[contains(.,'"+username+"')]//tr[@ng-repeat='line in lines[group.name]']//span[@ng-if='line.firstInSubGroup && line.description']";
        String TotalTime = "//tbody[contains(.,'"+username+"')]//td[@class='yt-table__cell time-report__value yt-bold' and not(@ng-if='shouldShowEstimation()')]";
        String Time = "//tbody[contains(.,'"+username+"')]//tr[@ng-repeat='line in lines[group.name]']//td[@class='yt-table__cell time-report__value' and not(@ng-if='shouldShowEstimation()')]";

        List<WebElement> works = driver.findElements(By.xpath(Work));
        List<WebElement> times = driver.findElements(By.xpath(Time));
        String TotalTiming = driver.findElement(By.xpath(TotalTime)).getText();

        ArrayList<Map<String,String>> ArWorks = new ArrayList<>();

        for(int i = 0;i<works.size();i++){
            Map<String,String> work = new HashMap<>();
            work.put("workname",works.get(i).getText());
            work.put("time",times.get(i).getText());
            ArWorks.add(work);
        }
        return new Card(username,TotalTiming,ArWorks);
    }






    // Define data type
    private static class Card {
        String Username;
        String TotalTime;
        ArrayList<Map<String,String>> work ;

        public Card(String username, String totalTime, ArrayList<Map<String, String>> work) {
            Username = username;
            TotalTime = totalTime;
            this.work = work;
        }
    }
}
