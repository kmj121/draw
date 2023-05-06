package com.example.draw.utils;

import com.example.draw.dto.ReportData;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class BuildUtils {

    private static Logger logger = LoggerFactory.getLogger(WordUtils.class);

    public static void buildProgram(ReportData data, String fileNameWord, String fileNamePdf) throws Exception {

//        logger.info("fileNameWord={}, fileNamePdf={}", fileNameWord, fileNamePdf);
        Map<String, Object> dataMap = new HashMap<>();
        Map<String, List<List<String>>> tableMap = new HashMap<>();

        //word
//        ProgramsWordUtils programsWordUtils = new ProgramsWordUtils();
//        programsWordUtils.buildWord(fileNameWord, dataMap, tableMap);

        //pdf
        ProgramsPdfUtils programsPdfUtils = new ProgramsPdfUtils();
        programsPdfUtils.buildPdf(fileNamePdf, dataMap, tableMap);

    }
}
