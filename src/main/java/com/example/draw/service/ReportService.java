package com.example.draw.service;

import com.alibaba.fastjson.JSONObject;
import com.example.draw.utils.BuildUtils;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

@Service
public class ReportService {
    @Value("${attachFolder}")
    private String attachFolder;
    @Value("${getSignUrl}")
    private String getSignUrl;

    public void getReport(JSONObject data) throws Exception {
        //创建word文档
        String wordFileName = data.get("_id") + ".docx";
        data = (JSONObject) data.get("ziJieReportContent");
        buildWord(wordFileName, data);
    }

    /**
     * 生成word以及pdf文件
     */
    public void buildWord(String wordFilePath, JSONObject data)throws Exception {
        BuildUtils.buildProgram(data, attachFolder + wordFilePath, getSignUrl);
    }
}
