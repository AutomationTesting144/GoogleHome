package com.example.a310287808.huedisco;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.json.JSONException;
import org.openqa.selenium.By;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.concurrent.TimeUnit;

import io.appium.java_client.android.AndroidDriver;

/**
 * Created by 310287808 on 11/15/2017.
 */

public class BridgeAuthenticationFirstTime {

    public String ActualResult;
    public String Comments;
    public String ExpectedResult;
    public String Status;
    public boolean ob3;
    public void BridgeAuthenticationFirstTime(AndroidDriver driver, String fileName, String APIVersion, String SWVersion) throws IOException, JSONException, InterruptedException {
        //Checking the state of lights from API and turning them OFf if they are ON
        driver.navigate().back();
        driver.navigate().back();

        //Opening Harmony App
        driver.findElement(By.xpath("//android.widget.TextView[@text='Hue Disco']")).click();
        TimeUnit.SECONDS.sleep(5);


        //looking for pushlink
        ob3= driver.findElement(By.id("nl.ijsdesign.huedisco:id/buttonGreen")).isDisplayed();

        Boolean result=(ob3==true);

        if(result==true) {

            Status = "1";
            ActualResult = "Application is asking user to press bridge pushlink while connecting for first time";
            Comments = "NA";
            ExpectedResult= "Application should ask user to press bridge pushlink while connecting for first time";
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);
        } else {
            Status = "0";
            ActualResult = "Application is not asking user to press bridge pushlink while connecting for first time";
            Comments = "FAIL: Bridge is already connected to the application";
            ExpectedResult= "Application should ask user to press bridge pushlink while connecting for first time";
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);
        }


        //going back
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
        r2c2.setCellValue("2");

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
