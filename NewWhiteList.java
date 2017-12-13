package com.example.a310287808.huedisco;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.json.JSONException;
import org.json.JSONObject;

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
import java.util.Date;

import io.appium.java_client.android.AndroidDriver;

/**
 * Created by 310287808 on 11/15/2017.
 */

public class NewWhiteList {
    public String whiteListString;
    public String [] parts ;
    public int counter=0;
    public String iftttSubString;
    public String WLname;
    public String Status;
    public String Comments;
    public String ActualResult;
    public String ExpectedResult;

    public void NewWhiteList(AndroidDriver driver, String fileName, String APIVersion, String SWVersion) throws JSONException, IOException {
        driver.navigate().back();
        HttpURLConnection connection;
        URL url = new URL("http://192.168.86.21/api/BCcSt4PIHiUvfCO-JmYwuQfWvwQRH8HRXPMM0lzi/config");
        connection = (HttpURLConnection)url.openConnection();
        connection.connect();

        InputStream stream = connection.getInputStream();

        BufferedReader reader = new BufferedReader(new InputStreamReader(stream));

        StringBuffer br = new StringBuffer();

        String line = " ";
        while((line = reader.readLine())!=null){
            br.append(line);
        }
        String output = br.toString();

        Date curDate = new Date();
        SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
        String DateToStr = format.format(curDate);

        JSONObject jsonObject = new JSONObject(output);
        Object whitelistObject = jsonObject.get("whitelist");
        whiteListString = whitelistObject.toString();

        parts = whiteListString.split("}");
        for(int i =0;i<parts.length;i++){
            if(parts[i].contains("hue-hca-actions-on-google")){
                WLname = parts[i].substring(1,43);
                iftttSubString = parts[i].substring(98,108);
                boolean result=iftttSubString.equals(DateToStr);
                if(result==true){
                    counter++;
                }

            }
        }
        if(counter==0){
            Status = "0";
            ActualResult = "New Whitelist for Google Home is not created Today";
            Comments = "Fail: Whitelist not created: "+whiteListString;
            ExpectedResult= "New Whitelist should be created whenever any new device is connected to the application for  the first time";
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);

        }
        else{
            Status = "1";
            ActualResult = "New Whitelist for Google Home is created Today";
            Comments = "NA";
            ExpectedResult= "New Whitelist should be created whenever any new device is connected to the application for the first time";
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);
        }
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
        HSSFCell r2c1 = row2.createCell((short)0);
        r2c1.setCellValue(CurrentdateTime);

        HSSFCell r2c2 = row2.createCell((short)1);
        r2c2.setCellValue("1");

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
