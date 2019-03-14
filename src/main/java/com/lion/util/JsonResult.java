package com.lion.util;



import com.alibaba.fastjson.JSONObject;

import java.util.HashMap;
import java.util.Map;

public class JsonResult {

    public final static int SUCCESS = 0;
    public final static int FAIL = 1;

    /**
     * 成功
     *
     * @return
     */
    public static JSONObject ok() {
        return ok("成功");

    }

    public static JSONObject ok(String key, Object value) {
        JSONObject map = new JSONObject();
        map.put(key, value);
        return ok(map);
    }

    public static JSONObject ok(String key1, Object value1,
                                String key2, Object value2) {
        JSONObject map = new JSONObject();
        map.put(key1, value1);
        map.put(key2, value2);
        return ok(map);
    }

    public static JSONObject ok(String key1, Object value1,
                                String key2, Object value2,
                                String key3, Object value3) {
        JSONObject map = new JSONObject();
        map.put(key1, value1);
        map.put(key2, value2);
        map.put(key3, value3);
        return ok(map);
    }

    /**
     * 成功
     *
     * @return
     */
    public static JSONObject ok(String toast) {
        return ok(toast, new HashMap<>());

    }
    /**
     * 成功
     * @param
     * @return
     */
    public static JSONObject ok(Object obj) {
        JSONObject response = new JSONObject();
        response.put("code", SUCCESS);
        response.put("msg", "成功");
        response.put("detail", obj);
        return response;

    }
    /**
     * 成功
     *
     * @return
     */
    public static JSONObject ok(Map<String, Object> detail) {
        return ok("成功", detail);
    }

    /**
     * 成功
     *
     * @return
     */
    public static JSONObject ok(String toast, Map<String, Object> detail) {
        JSONObject response = new JSONObject();
        response.put("code", SUCCESS);
        response.put("msg", toast);
        response.put("detail", detail);
        return response;
    }

    /**
     * 失败
     *
     * @return
     */
    public static JSONObject fail(String msg) {
        return fail(FAIL, msg);
    }

    /**
     * 失败
     *
     * @return
     */
    public static JSONObject fail(int errorCode, String msg) {
        JSONObject response = new JSONObject();
        response.put("code", errorCode);
        response.put("msg", msg);
        return response;
    }
}
