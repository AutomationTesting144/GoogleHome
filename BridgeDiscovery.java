package com.example.a310287808.huedisco;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.json.JSONException;
import org.openqa.selenium.By;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.concurrent.TimeUnit;

import io.appium.java_client.android.AndroidDriver;

/**
 * Created by 310287808 on 11/14/2017.
 */

public class BridgeDiscovery {
    public String ActualResult;
    public String Comments;
    public String ExpectedResult;
    public String Status;
    public boolean ob3;

    public boolean isElementPresent(By by, int timeOutInSec, AndroidDriver d)
    {
        for(int i = 0; i < timeOutInSec; i++)
        {
            try {
                Thread.sleep(1000);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
//            d.
            if((d.findElements(by)).size() >= 0)
                return true;

        }
        return false;
    }

    public void BridgeDiscovery(AndroidDriver driver, String fileName, String APIVersion, String SWVersion) throws IOException, JSONException, InterruptedException {
        //Checking the state of lights from API and turning them OFf if they are ON
        driver.navigate().back();
        driver.navigate().back();

        //Opening Harmony App
        driver.findElement(By.xpath("//android.widget.TextView[@text='Hue Disco']")).click();
        TimeUnit.SECONDS.sleep(5);

//        Clicking on top left side menu
        driver.findElement(By.xpath("//android.widget.ImageButton[@index='0']")).click();
        TimeUnit.SECONDS.sleep(2);

        //Clicking on Hue bridge option
        driver.findElement(By.xpath("//android.widget.CheckedTextView[@text='Hue Bridge']")).click();
        TimeUnit.SECONDS.sleep(2);



        //Clicking on FIND NEW BRIDGE
        long start = System.currentTimeMillis();
        driver.findElement(By.xpath("//android.widget.Button[@text='FIND NEW BRIDGE']")).click();
  //    TimeUnit.SECONDS.sleep(30);
        long startTime = System.currentTimeMillis();

        WebDriverWait wait = new WebDriverWait(driver, 30);

        wait.until(ExpectedConditions.invisibilityOfElementLocated(By.id("android:id/body")));
        long stopTime = System.currentTimeMillis();
        double elapsedTime=(stopTime - startTime)/1000.0;



        if (elapsedTime<=30.0)

        {
            Status = "1";
            ActualResult = "Bridge discovery is taking less then 30secs";
            Comments = "NA";
            ExpectedResult= "Bridge discovery by an application should take less then or equal to 30 secs";
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);
        } else {
            Status = "0";
            ActualResult = "Bridge discovery is takes more then 30secs";
            Comments = "NA";
            ExpectedResult= "Bridge discovery by an application should take less then or equal to 30 secs";
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);
        }

        //going back
        driver.navigate().back();
        driver.navigate().back();


        driver.navigate().back();
        //CALLING THE FUNCTION FOR WRITING THE CODE IN EXCEL FILE
        storeResultsExcel(Status, ActualResult, Comments, fileName, ExpectedResult,APIVersion,SWVersion);

    }
    //WRITING THE RESULT IN EXCEL FILE
    public String CurrentdateTime;
    public int nextRowNumber;
    public void storeResultsExcel (String excelStatus, String excelActualResult, String excelComments, String resultFileName, String ExcelExpectedResult, String resultAPIVersion, String resultSWVersion) throws IOException
    {
        Calendar cal = Calendar.getInstance();
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
        CurrentdateTime = sdf.format(cal.getTime());
        FileInputStream fsIP = new FileInputStream(new File("C:\\Users\\310287808\\AndroidStudioProjects\\AnkitasTrial\\" + resultFileName));
        HSSFWorkbook workbook = new HSSFWorkbook(fsIP);
        nextRowNumber = workbook.getSheetAt(0).getLastRowNum();
        nextRowNumber++;
        HSSFSheet sheet = workbook.getSheetAt(0);

        HSSFRow row2 = sheet.createRow(nextRowNumber);
        HSSFCell r2c1 = row2.createCell((short) 0);
        r2c1.setCellValue(CurrentdateTime);

        HSSFCell r2c2 = row2.createCell((short) 1);
        r2c2.setCellValue("3");

        HSSFCell r2c3 = row2.createCell((short) 2);
        r2c3.setCellValue(excelStatus);

        HSSFCell r2c4 = row2.createCell((short) 3);
        r2c4.setCellValue(excelActualResult);

        HSSFCell r2c5 = row2.createCell((short) 4);
        r2c5.setCellValue(excelComments);

        HSSFCell r2c6 = row2.createCell((short) 5);
        r2c6.setCellValue(resultAPIVersion);

        HSSFCell r2c7 = row2.createCell((short) 6);
        r2c7.setCellValue(resultSWVersion);

        fsIP.close();
        FileOutputStream out =new FileOutputStream(new File("C:\\Users\\310287808\\AndroidStudioProjects\\AnkitasTrial\\" + resultFileName));
        workbook.write(out);
        out.close();

    }
}
