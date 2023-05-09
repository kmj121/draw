package com.example.draw.service;

import com.example.draw.dto.ReportData;
import com.example.draw.utils.BuildUtils;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

@Service
public class WordService {

    @Value("${attachFolder}")
    private String attachFolder;

    public void test(ReportData data) throws Exception {
        //创建word文档
        String wordFileName = "word1.docx";
        //创建pdf文档
        String pdfFileName = "pdf1.pdf";

        buildWord(wordFileName, pdfFileName, data);
    }

    /**
     * 生成word以及pdf文件
     */
    public void buildWord(String wordFilePath, String pdfFilePath, ReportData data)throws Exception {
        BuildUtils.buildProgram(data, attachFolder + wordFilePath, attachFolder + pdfFilePath);
    }
}
