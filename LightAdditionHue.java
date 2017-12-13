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
import java.util.HashMap;
import java.util.List;
import java.util.concurrent.TimeUnit;

import io.appium.java_client.MobileElement;
import io.appium.java_client.android.AndroidDriver;

/**
 * Created by 310287808 on 11/16/2017.
 */

public class LightAdditionHue {
    public int lightCounter=0;
    public List dropListHueOld;
    public List dropListHueNew;
    public String lightNameFromHue;
    public MobileElement listItemHueNew;
    public MobileElement listItemHueOld;
    public int oldCounter = 0;
    public int newCounter = 0;
    public int counterForNewLights=1;
    public int hashmapcounter=0;
    public String ActualResult;
    public String ExpectedResult;
    public String Status;
    public String Comments;

    public void LightAddition(AndroidDriver driver, String fileName, String APIVersion, String SWVersion) throws IOException, JSONException, InterruptedException {

        driver.navigate().back();
        driver.navigate().back();

        //Opening Hue application
        driver.findElement(By.xpath("//android.widget.TextView[@text='Hue']")).click();
        TimeUnit.SECONDS.sleep(5);

        //Clicking on settings button
        driver.findElement(By.xpath("//android.view.View[@bounds='[1122,280][1398,448]']")).click();
        TimeUnit.SECONDS.sleep(2);

        //Selecting light setup
        driver.findElement(By.xpath("//android.widget.TextView[@text='Light setup']")).click();
        HashMap<String,Integer> oldlist = new HashMap<>();
        HashMap<String,Integer> newList = new HashMap<>();

        //Locate all drop down list elements
        dropListHueOld = driver.findElements(By.id("com.philips.lighting.hue2:id/list_item_title"));

        //Extract text from each element of drop down list one by one.
        for(int i=0; i< dropListHueOld.size(); i++){
            listItemHueOld = (MobileElement) dropListHueOld.get(i);
            lightNameFromHue = listItemHueOld.getText();

            if(lightNameFromHue.contains("Living room")==true){
                break;
            }else{
                oldlist.put(lightNameFromHue,oldCounter);
                oldCounter++;
            }
            //System.out.println("List of All Elements on Page:"+listItemHueOld.getText());
        }

        //clicking on Add button
        driver.findElement(By.xpath("//android.widget.ImageButton[@bounds='[1174,2126][1370,2322]']")).click();
        //Opening Add Lights Page and clicking on SEARCH button
        driver.findElement(By.id("com.philips.lighting.hue2:id/start_search_button")).click();
        TimeUnit.SECONDS.sleep(60);

        dropListHueNew = driver.findElements(By.id("com.philips.lighting.hue2:id/list_item_title"));

        //Extract text from each element of drop down list one by one.
        for(int i=0; i< dropListHueNew.size(); i++)
        {
            listItemHueNew = (MobileElement) dropListHueNew.get(i);
            lightNameFromHue = listItemHueNew .getText();
            if(lightNameFromHue.contains("Living room")==true){
                break;
            }
            else
            {
                newList.put(lightNameFromHue,newCounter);
                newCounter++;
            }
            //System.out.println("List of All Elements on Page:"+listItemHueNew.getText());
        }

        HashMap<String,Integer> setOfNewLights = new HashMap<>();
        if(newList.size()>oldlist.size()){
            for(String newKey : newList.keySet()){
                for(String oldKey : oldlist.keySet()){
//                    System.out.println("New Key:"+newKey);
//                    System.out.println("Old Key:"+oldKey);

                    if(newKey.contains(oldKey)==true)
                    {
                        counterForNewLights=0;
                        break;
                    }
                    else
                    {
                        counterForNewLights++;
                    }

                }
                //System.out.println(counterForNewLights);
                if(counterForNewLights!=0){
                    setOfNewLights.put(newKey,hashmapcounter);
                    hashmapcounter++;
                }
            }
        }
        //Going back from Hue application to home screen
        driver.findElement(By.xpath("//android.widget.ImageButton[@bounds='[0,84][196,280]']")).click();
        driver.navigate().back();

        //Opening Harmony App
        driver.findElement(By.xpath("//android.widget.TextView[@text='Hue Disco']")).click();
        TimeUnit.SECONDS.sleep(5);

        // click on bulbs
        driver.findElement(By.xpath("//android.widget.TextView[@text='BULBS']")).click();
        TimeUnit.SECONDS.sleep(2);

        //Locate all drop down list elements
        List dropListNew = driver.findElements(By.id("nl.ijsdesign.huedisco:id/textViewBulbName"));
        StringBuffer sb = new StringBuffer();

        //Extract text from each element of drop down list one by one.
        for (int l = 0; l < dropListNew.size(); l++) {
            MobileElement listItemNew = (MobileElement) dropListNew.get(l);
            for(String compareName : setOfNewLights.keySet()){
                if(compareName.contains(listItemNew.getText())==true){
                    lightCounter++;
                    sb.append(compareName);
                    sb.append("\n");
                }
                else {continue;}
            }
        }
        if (lightCounter==0)
        {Status = "0";
            ActualResult ="Light: "+setOfNewLights.toString()+" is added in list";
            Comments = "FAIL: Lights are not added in the application";
            ExpectedResult= "Light: "+setOfNewLights.toString()+" should be added in list";
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);
        }
        else {
            Status = "1";
            ActualResult ="Light: "+setOfNewLights.toString()+" is added in list";
            Comments = "NA";
            ExpectedResult= "Light: "+setOfNewLights.toString()+" should be added in list";
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

        HSSFCell r2c2 = row2.createCell((short) 1);
        r2c2.setCellValue("10");

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
