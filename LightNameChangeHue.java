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
import java.util.List;
import java.util.Random;
import java.util.concurrent.TimeUnit;

import io.appium.java_client.MobileElement;
import io.appium.java_client.android.AndroidDriver;

/**
 * Created by 310287808 on 11/20/2017.
 */

public class LightNameChangeHue {

    MobileElement listItem;
    public int lightCounter=0;
    public String ActualResult;
    public String ExpectedResult;
    public String Status;
    public String Comments;

    public void LightRename (AndroidDriver driver, String fileName, String APIVersion, String SWVersion) throws IOException, JSONException, InterruptedException  {
        driver.navigate().back();
        driver.navigate().back();

        //Opening Hue application
        driver.findElement(By.xpath("//android.widget.TextView[@text='Hue']")).click();
        TimeUnit.SECONDS.sleep(10);

        //Clicking on settings button
        driver.findElement(By.xpath("//android.view.View[@bounds='[1122,280][1398,448]']")).click();
        TimeUnit.SECONDS.sleep(2);

        //Selecting light setup
        driver.findElement(By.xpath("//android.widget.TextView[@text='Light setup']")).click();

        //Choosing light to remane
        driver.findElement(By.xpath("//android.widget.ImageButton[@bounds='[1230,343][1440,553]']")).click();
        TimeUnit.SECONDS.sleep(3);

        //editing the name
        driver.findElement(By.id("com.philips.lighting.hue2:id/form_field_text")).click();
        String lightValue=driver.findElement(By.id("com.philips.lighting.hue2:id/form_field_text")).getText();
        //Clearing the old name
        driver.findElement(By.id("com.philips.lighting.hue2:id/form_field_text")).clear();
        TimeUnit.SECONDS.sleep(2);
        //Entering new name
        Random rand = new Random();
        int  n = rand.nextInt() + 1;
        String status = String.valueOf(n);
        driver.findElement(By.id("com.philips.lighting.hue2:id/form_field_text")).sendKeys(status);
        TimeUnit.SECONDS.sleep(5);
        driver.hideKeyboard();

        //Going back from the application
        driver.findElement(By.xpath("//android.widget.ImageButton[@bounds='[0,84][196,280]']")).click();
        driver.findElement(By.xpath("//android.widget.ImageButton[@bounds='[0,84][196,280]']")).click();
        driver.navigate().back();

        //Opening Harmony App
        driver.findElement(By.xpath("//android.widget.TextView[@text='Hue Disco']")).click();
        TimeUnit.SECONDS.sleep(10);

        // click on bulbs
        driver.findElement(By.xpath("//android.widget.TextView[@text='BULBS']")).click();
        TimeUnit.SECONDS.sleep(2);

        //Locate all drop down list elements
        List dropList = driver.findElements(By.id("nl.ijsdesign.huedisco:id/textViewBulbName"));
        StringBuffer sb = new StringBuffer();
        for(int i=0; i< dropList.size(); i++){
            listItem = (MobileElement) dropList.get(i);
            Boolean result=listItem.getText().equals(status);
            if (result==true){
                lightCounter++;
                sb.append(lightValue);
                sb.append("\n");
            }
            else{
                continue;
            }
        }
        if (lightCounter==0)
        {
            Status = "0";
            ActualResult ="Light: "+lightValue+" is not renamed to "+status;
            Comments = "FAIL: Light is not renamed";
            ExpectedResult="Light: "+lightValue+" should be renamed to "+status;
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);

        }
        else {
            Status = "1";
            ActualResult ="Light: " + lightValue + " is renamed to " + status;
            Comments = "NA";
            ExpectedResult="Light: "+lightValue+" should be renamed to "+status;
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);
        }

        //Going back on home from Harmony
        driver.navigate().back();
        storeResultsExcel(Status, ActualResult, Comments, fileName, ExpectedResult,APIVersion,SWVersion);
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

        HSSFCell r2c2 = row2.createCell((short)1);
        r2c2.setCellValue("11");

        HSSFCell r2c3 = row2.createCell((short)2);
        r2c3.setCellValue(excelStatus);

        HSSFCell r2c4 = row2.createCell((short)3);
        r2c4.setCellValue(excelActualResult);

        HSSFCell r2c5 = row2.createCell((short)4);
        r2c5.setCellValue(excelComments);

        HSSFCell r2c6 = row2.createCell((short)5);
        r2c6.setCellValue(resultAPIVersion);

        HSSFCell r2c7 = row2.createCell((short)6);
        r2c7.setCellValue(resultSWVersion);

        fsIP.close();
        FileOutputStream out =
                new FileOutputStream(new File("C:\\Users\\310287808\\AndroidStudioProjects\\AnkitasTrial\\" + resultFileName));
        workbook.write(out);
        out.close();


    }
}
