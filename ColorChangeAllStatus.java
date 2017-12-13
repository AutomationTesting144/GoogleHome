package com.example.a310287808.huedisco;

import org.json.JSONException;
import org.json.JSONObject;

/**
 * Created by 310287808 on 11/21/2017.
 */

public class ColorChangeAllStatus {

    public String finalStatus;

    public String ColorChangeStatus(String input) throws JSONException {

        JSONObject jsonObject = new JSONObject(input);
        //System.out.println(jsonObject.toString());

        Object ob =  jsonObject.get("action");
        String newString = ob.toString();
        //System.out.println("OB:"+newString);

        JSONObject jsonObject1 = new JSONObject(newString);
        Object ob1 = jsonObject1.get("xy");
        finalStatus = ob1.toString();

        return finalStatus;

    }
}
