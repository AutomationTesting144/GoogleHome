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

import io.appium.java_client.android.AndroidDriver;

/**
 * Created by 310287808 on 11/15/2017.
 */

public class WhiteListFriendlyName {


    public String Status;
    public String Comments;
    public String ActualResult;
    public String ExpectedResult;

    public void WhiteListFriendlyName(AndroidDriver driver, String fileName, String APIVersion, String SWVersion) throws JSONException, IOException {
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

        JSONObject jsonObject = new JSONObject(output);
        Object whitelistObject = jsonObject.get("whitelist");
        String whitelistString = whitelistObject.toString();

        if(whitelistString.contains("hue-hca-actions-on-google")){
            Status = "1";
            ActualResult = "Bridge has User Friendly whitelist name for Google Home";
            Comments = "NA";
            ExpectedResult= "Bridge should have User Friendly whitelist name for Google Home";
            System.out.println("Result: " + Status + "\n" + "Comment: " + Comments+ "\n"+"Actual Result: "+ActualResult+ "\n"+"Expected Result: "+ExpectedResult);

        }
        else {
            Status = "0";
            ActualResult = "Bridge does not have User Friendly whitelist name for Google Home";
            Comments = "FAIL: User name is not friendly or WhiteList is not created today";
            ExpectedResult= "Bridge should have User Friendly whitelist name for Google Home";
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
        r2c2.setCellValue("2");

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
