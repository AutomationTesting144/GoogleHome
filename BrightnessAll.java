package com.example.a310287808.huedisco;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.json.JSONException;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.concurrent.TimeUnit;

import io.appium.java_client.android.AndroidDriver;

/**
 * Created by 310287808 on 12/12/2017.
 */

public class BrightnessAll {
    public String IPAddress = "192.168.86.21/api";
    public String HueUserName = "BCcSt4PIHiUvfCO-JmYwuQfWvwQRH8HRXPMM0lzi";
    public String HueBridgeParameterType = "groups/0";
    public String finalURL;
    public String lightStatusReturned;
    public String Status;
    public String Comments;
    public String ActualResult;
    public String ExpectedResult;
    public String AllLightIDs;

    public void BrightnessAll(AndroidDriver driver, String fileName, String APIVersion, String SWVersion) throws IOException, JSONException, InterruptedException {
        driver.navigate().back();
        driver.navigate().back();

//       long pressing the home button to activate mic

        Runtime.getRuntime().exec("adb shell input keyevent --longpress 5");
        TimeUnit.SECONDS.sleep(3);

        //Get the size of screen.
        Dimension size = driver.manage().window().getSize();


        //Find swipe start and end point from screen's with and height.
        //Find starty point which is at bottom side of screen.
        int starty = (int) (size.height * 0.80);
        //Find endy point which is at top side of screen.
        int endy = (int) (size.height * 0.20);
        //Find horizontal point where you wants to swipe. It is in middle of screen width.
        int startx = size.width / 2;


        //Swipe from Bottom to Top.
        driver.swipe(startx, starty, startx, endy, 3000);
        Thread.sleep(2000);


        // click on keyboard
        driver.findElement(By.xpath("//android.widget.ImageView[@content-desc='Type mode']")).click();


        //Send turn on all the lights
        driver.findElement(By.id("com.google.android.googlequicksearchbox:id/input_text")).sendKeys("bright all lights to 100%");
        TimeUnit.SECONDS.sleep(2);

        //Click on send button
        driver.findElement(By.id("com.google.android.googlequicksearchbox:id/send_button")).click();
        TimeUnit.SECONDS.sleep(2);


        driver.navigate().back();
        driver.navigate().back();

        HttpURLConnection connection;

        finalURL = "http://" + IPAddress + "/" + HueUserName + "/" + HueBridgeParameterType;
        URL url = new URL(finalURL);
        connection = (HttpURLConnection) url.openConnection();
        connection.connect();

        InputStream stream = connection.getInputStream();

        BufferedReader reader = new BufferedReader(new InputStreamReader(stream));

        StringBuffer br = new StringBuffer();

        String line = " ";
        while ((line = reader.readLine()) != null) {
            br.append(line);
        }
        String output = br.toString();
        //System.out.println(output);

        DimmingLightStatus DimStatus=new DimmingLightStatus();
        lightStatusReturned = DimStatus.DimmingLightStatus(output);
        System.out.println(lightStatusReturned);


        StringBuffer sb = new StringBuffer();


        if (lightStatusReturned.equals("254")){
            Status = "1";
            ActualResult ="All the lights should get brighter to 100%";
            Comments = "NA";
            ExpectedResult="Lights should increase its brightness to 100% ";
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);

        }else
        {
            Status = "0";
            ActualResult = "colors are changing continously";
            Comments = "FAIL:Lights are not brighter enough";
            ExpectedResult="Lights should change their color ";
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);


        }
        storeResultsExcel(Status, ActualResult, Comments, fileName, ExpectedResult,APIVersion,SWVersion);

        driver.navigate().back();
    }
    public String CurrentdateTime;
    public int nextRowNumber;
    public void storeResultsExcel(String excelStatus, String excelActualResult, String excelComments, String resultFileName, String ExcelExpectedResult
            ,String resultAPIVersion, String resultSWVersion) throws IOException {

        Calendar cal = Calendar.getInstance();
        SimpleDateFormat sdf  = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
        CurrentdateTime = sdf.format(cal.getTime());
        FileInputStream fsIP = new FileInputStream(new File("C:\\Users\\310287808\\AndroidStudioProjects\\AnkitasTrial\\" + resultFileName));
        HSSFWorkbook workbook = new HSSFWorkbook(fsIP);
        nextRowNumber=workbook.getSheetAt(0).getLastRowNum();
        nextRowNumber++;
        HSSFSheet sheet = workbook.getSheetAt(0);

        HSSFRow row2 = sheet.createRow(nextRowNumber);
        HSSFCell r2c1 = row2.createCell((short) 0);
        r2c1.setCellValue(CurrentdateTime);

        HSSFCell r2c2 = row2.createCell((short) 1);
        r2c2.setCellValue("9");

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
        FileOutputStream out =
                new FileOutputStream(new File("C:\\Users\\310287808\\AndroidStudioProjects\\AnkitasTrial\\" + resultFileName));
        workbook.write(out);
        out.close();


    }

}
