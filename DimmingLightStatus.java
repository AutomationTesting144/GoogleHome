package com.example.a310287808.huedisco;
import org.json.JSONException;
import org.json.JSONObject;

/**
 * Created by 310287808 on 6/9/2017.
 */

public class DimmingLightStatus {
    public String finalStatus;

    public String DimmingLightStatus(String input) throws JSONException {

        JSONObject jsonObject = new JSONObject(input);
        //System.out.println(jsonObject.toString());

        Object ob =  jsonObject.get("action");
        String newString = ob.toString();
        //System.out.println("OB:"+newString);

        JSONObject jsonObject1 = new JSONObject(newString);
        Object ob1 = jsonObject1.get("bri");
        finalStatus = ob1.toString();

        return finalStatus;

    }
}
