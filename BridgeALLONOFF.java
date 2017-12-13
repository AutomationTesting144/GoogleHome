package com.example.a310287808.huedisco;

import org.json.JSONException;
import org.json.JSONObject;

/**
 * Created by 310287808 on 11/16/2017.
 */

public class BridgeALLONOFF {
    public String finalStatus;

    public String AllONOFFStatus(String input) throws JSONException {

        JSONObject jsonObject = new JSONObject(input);
        //System.out.println(jsonObject.toString());

        Object ob =  jsonObject.get("state");
        String newString = ob.toString();
        //System.out.println("OB:"+newString);

        JSONObject jsonObject1 = new JSONObject(newString);
        Object ob1 = jsonObject1.get("all_on");
        finalStatus = ob1.toString();
        return finalStatus;

    }
}
