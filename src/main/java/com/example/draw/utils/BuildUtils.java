package com.example.draw.utils;

import com.alibaba.fastjson.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class BuildUtils {
    private static Logger logger = LoggerFactory.getLogger(BuildUtils.class);

    public static void buildProgram(JSONObject dataMap, String fileNameWord, String getSignUrl) throws Exception {
        //word
        ProgramsWordUtils programsWordUtils = new ProgramsWordUtils();
        programsWordUtils.buildWord(fileNameWord, dataMap, getSignUrl);
    }
}
