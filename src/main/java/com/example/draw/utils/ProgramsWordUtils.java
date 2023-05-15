package com.example.draw.utils;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.net.URISyntaxException;
import java.util.*;

public class ProgramsWordUtils extends WordUtils {
    private String token = "Bearer eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIxMjM0NTY3ODkwIiwibmFtZSI6IkpvaG4gRG9lIiwiYWRtaW4iOnRydWUsImlhdCI6MTUxNjIzOTAyMn0.HojZlgf6Cuh7lk66uSst2U6CDv9Uq0Ccj_K0_0mB6Uev72Om1J7QeFqwNdQ2hcdroOZiz22hHG07vZwXknLlBofN-Y-o-U7_xlB54fye8JFaaBPkiZVemvsJ5UrfMHKe8NHP4ezXsRgi88R5Yy7UxgVR1CMjaQIGU5UF2iKL5LiTQoZVMCCi7evQJ0RjqO8KwK04T6DdtajzYEd8SWQQuusVofSIkEaGSQq5oMpQmJ_sour_Zw-c5NDtyGSsVA3ono0ETy9TK8qG_El3EzlrpV-Xd7Yvjdyxdz2Iq2SaC1g4Ab5SfqNin4m28C3P8ioz6sXSzuErjlgAaGc5ew23IA";

    @Override
    public void generateWord(XWPFDocument document, JSONObject params, String getSignUrl) throws Exception {
        // 页眉页脚
        createHeaderAndFooterSpecial(document);

        // 报告封面
        if (params.getOrDefault("homePage", null) != null) {
            buildHomePage(document, params);
        }

        // 下一页
        pageBreakSpecial(document);

        // 委托信息/报告概览
        if (params.getOrDefault("clientInfoAndReportOverview", null) != null) {
            buildWTXXAndBGGL(document, params);
        }

        // 下一页
        pageBreakSpecial(document);

        // 第一部分：基本信息及详情
        if (params.getOrDefault("basicInfoAndDetail", null) != null) {
            buildBasicMessageAndDetail(document, params);
        }

        // 下一页
        pageBreakSpecial(document);

        // 第二部分：工作履历及表现
        if (params.getOrDefault("xpAndPerf", null) != null) {
            buildXpPerfDetail(document, params);
        }

        // 下一页
        pageBreakSpecial(document);

        // 附件
        if (params.getOrDefault("attachment", null) != null) {
            builderAttachment(document, params, getSignUrl);
        }
    }

    public void builderAttachment(XWPFDocument document, JSONObject params, String getSignUrl) throws IOException, URISyntaxException, InvalidFormatException {
        // 标题
        buildTitleSpecial(document, " 附件");
        // 空一行
        blankParagraph(document);
        // 学历证书
        if (params.getOrDefault("attachment", null) != null) {
            JSONObject attachment = (JSONObject) params.get("attachment");
            if (attachment.getOrDefault("diplomaFilePath", null) != null) {
                getFile(document, attachment.getString("diplomaFilePath"), "diploma", getSignUrl);
            }

            // 学位证书
            if (attachment.getOrDefault("degreeFilePath", null) != null) {
                getFile(document, attachment.getString("degreeFilePath"), "degree", getSignUrl);
            }
            // 授权书
            if (attachment.getOrDefault("authLetterFilePath", null) != null) {
                getFile(document, attachment.getString("authLetterFilePath"), "authLetter", getSignUrl);
            }
        }

    }

    public void getFile(XWPFDocument document, String filePath, String type, String getSignUrl) throws IOException {
        Map<String, String> payload = new HashMap<>();
        payload.put("filePath", filePath);
        String response = HttpUtils.doPost2(getSignUrl, payload, null, token);
        JSONObject responseJson = JSON.parseObject(response);
        if (responseJson.getOrDefault("result", null) != null) {
            String tempUrl = (String) ((JSONArray) responseJson.get("result")).get(0);
            if (StringUtils.isNotBlank(tempUrl)) {
                InputStream is = HttpUtils.doGet2(tempUrl, null, null);
                if (is != null) {
                    try {
                        List<Object> strings1_1 = new ArrayList<>();
                        strings1_1.add(is);
                        List<Object> strings1_2 = new ArrayList<>();
                        if (type == "diploma") {
                            strings1_2.add("学历证书照片");
                        } else if (type == "degree") {
                            strings1_2.add("学位证书照片");
                        } else {
                            strings1_2.add("授权书");
                        }
                        List<List<Object>> content = new ArrayList<>();
                        content.add(strings1_1);
                        content.add(strings1_2);
                        buildTableSpecial_attachment(document, null, content,1, new Long[]{8310L}, 8310L, 0, null, colorBlack, 8, false, ParagraphAlignment.CENTER, type);
                    } catch (Exception e) {
                        e.printStackTrace();
                    } finally {
                        if (null != is) {
                            try {
                                is.close();
                            } catch (IOException e) {
                                e.printStackTrace();
                            }
                        }
                    }
                    // 空一行
                    blankParagraph(document);
                }
            }
        }
    }

    /**
     * 创建表格
     * @param document
     * @param title 表格标题
     * @param content 表格内容
     * @param numColumn 表格列数，如果为null，则取数组tableWidths的长度为列数
     * @param tableWidths 数组类型，表格每列的宽度
     * @param tableWidth 表格宽度
     * @param displayBorder 是否隐藏边框 0:不展示，1：展示，2：自定义边框（上下左右有边框）
     * @param tableBackground 表格背景色
     * @param wordColor 字体颜色
     * @param fontSize 字体大小
     * @param bold 字体是否加粗
     * @param align 对齐方式
     * @param type 图片类型
     * @throws InvalidFormatException
     * @throws IOException
     * @throws URISyntaxException
     */
    public void buildTableSpecial_attachment(XWPFDocument document
            , List<String> title
            , List<List<Object>> content
            , Integer numColumn
            , Long[] tableWidths
            , Long tableWidth
            , Integer displayBorder
            , String tableBackground
            , String wordColor
            , Integer fontSize
            , Boolean bold
            , ParagraphAlignment align
            , String type) throws InvalidFormatException, IOException, URISyntaxException {
        int rowNum = content.size();
        int columnNum = numColumn == null ? tableWidths.length : numColumn;
        if (CollectionUtils.isNotEmpty(title)) {
            rowNum++;
        }
        XWPFTable xwpfTable = document.createTable(rowNum, columnNum);
        //表格居中显示
        CTTblPr ctTblPr = xwpfTable.getCTTbl().addNewTblPr();
        ctTblPr.addNewJc().setVal(STJc.CENTER);
        //设置表格宽度
        CTTblWidth ctTblWidth = ctTblPr.addNewTblW();
        ctTblWidth.setW(BigInteger.valueOf(tableWidth));
        //设置表格宽度为非自动
        ctTblWidth.setType(STTblWidth.DXA);
        Long cellWidth = null;
        //设置边框
        if (displayBorder == 0) {
            displayBorder(xwpfTable);
        } else if(displayBorder == 2) {
            // 自定义边框：上下左右有灰色边框
            customizeBorderSpecial(xwpfTable, colorGary);
        }

        //创建内容
        for (int j = 0; j < content.size(); j++) {
            //有标题时和无标题时取值位置不同
            XWPFTableRow xwpfTableRow = xwpfTable.getRow(j);
            if (j == 0) {
                //行高
                xwpfTableRow.setHeight(5000);
            } else {
                xwpfTableRow.setHeight(300);
            }
            List<Object> row = content.get(j);
            for (int i = 0; i < row.size() && i < columnNum; i++) {
                //单元格对象
                XWPFTableCell xwpfTableCell = xwpfTableRow.getCell(i);
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                //添加文本，14号/黑体/左对齐
                buildCellSpecialAttachment(xwpfTableCell, row.get(i), fontSize, bold, wordColor, align, cellWidth, null, tableBackground, type);
            }
        }
    }


    /**
     * 委托信息 / 报告概览
     * @param document
     */
    public void buildWTXXAndBGGL(XWPFDocument document, JSONObject params) throws IOException, URISyntaxException, InvalidFormatException {
        //获取"委托信息 / 报告概览"页数据
        JSONObject clientInfoAndReportOverview = (JSONObject) params.get("clientInfoAndReportOverview");

        // 标题
        buildTitleSpecial(document, " 委托信息 / 报告概览");

        // 空一行
        blankParagraph(document);

        // 候选人、委托日期等信息的表格
        JSONArray deliverInfo = (JSONArray) clientInfoAndReportOverview.get("deliverInfo");
        buildWTXXAndBGGLMainTable(document, deliverInfo);

        // 空一行
        blankParagraph(document);

        // 红黄蓝绿灯表格
        List<Object> strings = new ArrayList<>();
        strings.add("");
        // todo 111
        strings.add(new StringBuffer("/target/classes/static/image/red.jpeg"));
//        strings.add(new StringBuffer("/static/image/red.jpeg"));
        strings.add("高风险");
        // todo 111
        strings.add(new StringBuffer("/target/classes/static/image/yellow.jpeg"));
//        strings.add(new StringBuffer("/static/image/yellow.jpeg"));
        strings.add("一般风险");
        strings.add("");
        // todo 111
        strings.add(new StringBuffer("/target/classes/static/image/blue.jpeg"));
//        strings.add(new StringBuffer("/static/image/blue.jpeg"));
        strings.add("低风险/无法核实");
        // todo 111
        strings.add(new StringBuffer("/target/classes/static/image/green.jpeg"));
//        strings.add(new StringBuffer("/static/image/green.jpeg"));
        strings.add("无风险");
        List<List<Object>> content = new ArrayList<>();
        content.add(strings);
        // 总长度8310 一半4155 第一列空格 55 第二列 340 第三列 1000 第四列710 总共4155， 后续。。。
        buildTableSpecial(document, null, content,10, new Long[]{55L, 340L, 1710L, 340L, 1710L, 55L, 340L, 1710L, 340L, 1710L}, 400, 8310L, 2, null, null, colorBlack, 10, false, ParagraphAlignment.LEFT);

        // 空一行
        blankParagraph(document);

        // 报告概览表格
        JSONObject verifyCategory = (JSONObject) clientInfoAndReportOverview.get("verifyCategory");
        List<String> title = new ArrayList<>();
        title.add("核实类目");
        title.add("类目明细");
        title.add("核实状态");
        title.add("核实结果");

        // 表格行数（不包括标题）
        Integer tableRow = 0;
        // 整理出需要合并的单元格数据 格式如下：
        // [[1, 1, 1, 2]]
        // 第一个元素：1表示列合并，2表示行合并
        // 第二个元素：表示第几列或者第几行
        // 第三个元素：表示从第几行或者第几列开始
        // 第四个元素：表示从第几行或者第几列结束
        List<List<Integer>> mergeData = new ArrayList<>();
        List<List<Object>> contentVerifyCategory = new ArrayList<>();

        // 身份风险
        if (verifyCategory.getOrDefault("IDENTITY", null) != null) {
            solveVerifyCategoryData((JSONArray) verifyCategory.get("IDENTITY"), tableRow, contentVerifyCategory, mergeData);
        }
        // 社会风险
        if (verifyCategory.getOrDefault("VIOLATION", null) != null) {
            solveVerifyCategoryData((JSONArray) verifyCategory.get("VIOLATION"), tableRow, contentVerifyCategory, mergeData);
        }
        // 诉讼风险
        if (verifyCategory.getOrDefault("LEGAL", null) != null) {
            solveVerifyCategoryData((JSONArray) verifyCategory.get("LEGAL"), tableRow, contentVerifyCategory, mergeData);
        }
        // 商业风险
        if (verifyCategory.getOrDefault("BIZ", null) != null) {
            solveVerifyCategoryData((JSONArray) verifyCategory.get("BIZ"), tableRow, contentVerifyCategory, mergeData);
        }
        // 教育风险
        if (verifyCategory.getOrDefault("EDUCATION", null) != null) {
            solveVerifyCategoryData((JSONArray) verifyCategory.get("EDUCATION"), tableRow, contentVerifyCategory, mergeData);
        }
        // 工作履历风险
        if (verifyCategory.getOrDefault("XP", null) != null) {
            solveVerifyCategoryData((JSONArray) verifyCategory.get("XP"), tableRow, contentVerifyCategory, mergeData);
        }
        // 工作表现风险
        if (verifyCategory.getOrDefault("PERF", null) != null) {
            solveVerifyCategoryData((JSONArray) verifyCategory.get("PERF"), tableRow, contentVerifyCategory, mergeData);
        }
        buildTableSpecial5(document, title, contentVerifyCategory,4, new Long[]{2000L, 5310L, 1000L, 1000L}, 400, 8310L, 1, colorGary, null, colorBlack, 10, false, ParagraphAlignment.CENTER, mergeData);

        // 空一行
        blankParagraph(document);

        // 风险说明表格
        if (clientInfoAndReportOverview.getOrDefault("riskDetail", null) != null) {
            JSONObject riskDetail = (JSONObject) clientInfoAndReportOverview.get("riskDetail");
            // 红灯
            if (riskDetail.getOrDefault("RED", null) != null) {
                List<Object> redList = new ArrayList<>();
                redList.add(riskDetail.get("RED"));
                redList.add("");
                // todo 111
                redList.add(new StringBuffer("/target/classes/static/image/red.jpeg"));
//                redList.add(new StringBuffer("/static/image/red.jpeg"));
                List<List<Object>> contentRed = new ArrayList<>();
                contentRed.add(redList);
                // 总长度8310 一半4155 第一列空格 55 第二列 340 第三列 1000 第四列710 总共4155
                buildTableSpecial6(document, null, contentRed,3, new Long[]{7810L, 500L, 500L}, 1000, 8310L, 0, null, colorPink, colorBlack, 10, false, ParagraphAlignment.LEFT);
            }
            // 黄灯
            if (riskDetail.getOrDefault("YELLOW", null) != null) {
                List<Object> yellowList = new ArrayList<>();
                yellowList.add(riskDetail.get("YELLOW"));
                yellowList.add("");
                // todo 111
                yellowList.add(new StringBuffer("/target/classes/static/image/yellow.jpeg"));
//                yellowList.add(new StringBuffer("/static/image/yellow.jpeg"));
                List<List<Object>> contentYellow = new ArrayList<>();
                contentYellow.add(yellowList);
                // 总长度8310 一半4155 第一列空格 55 第二列 340 第三列 1000 第四列710 总共4155
                buildTableSpecial6(document, null, contentYellow,3, new Long[]{7810L, 500L, 500L}, 1000, 8310L, 0, null, colorOrange, colorBlack, 10, false, ParagraphAlignment.LEFT);
            }
            // 蓝灯
            if (riskDetail.getOrDefault("BLUE", null) != null) {
                List<Object> blueList = new ArrayList<>();
                blueList.add(riskDetail.get("BLUE"));
                blueList.add("");
                // todo 111
                blueList.add(new StringBuffer("/target/classes/static/image/blue.jpeg"));
//                blueList.add(new StringBuffer("/static/image/blue.jpeg"));
                List<List<Object>> contentBlue = new ArrayList<>();
                contentBlue.add(blueList);
                // 总长度8310 一半4155 第一列空格 55 第二列 340 第三列 1000 第四列710 总共4155
                buildTableSpecial6(document, null, contentBlue,3, new Long[]{7810L, 500L, 500L}, 1000, 8310L, 0, null, colorBlue2, colorBlack, 10, false, ParagraphAlignment.LEFT);
            }
        }

        // 空一行
        blankParagraph(document);

        // 声明对应的表格
        if (clientInfoAndReportOverview.getOrDefault("statement", null) != null) {
            String statement = clientInfoAndReportOverview.getString("statement");
            List<Object> strings5_1 = new ArrayList<>();
            strings5_1.add("");
            strings5_1.add("");
            strings5_1.add("");
            List<Object> strings5_2 = new ArrayList<>();
            strings5_2.add("");
            strings5_2.add(statement);
            strings5_2.add("");
            List<Object> strings5_3 = new ArrayList<>();
            strings5_3.add("");
            strings5_3.add("");
            strings5_3.add("");
            List<List<Object>> content5 = new ArrayList<>();
            content5.add(strings5_1);
            content5.add(strings5_2);
            content5.add(strings5_3);
            buildTableSpecial8(document, null, content5, 3, new Long[]{100L, 8110L, 100L}, 8310L, 2, null, null, 10, false, ParagraphAlignment.LEFT);
        }
    }

    public static void solveVerifyCategoryData(JSONArray jsonArray, Integer tableRow, List<List<Object>> contentVerifyCategory, List<List<Integer>> mergeData) {
        tableRow += jsonArray.size();
        if (jsonArray.size() > 1) {
            List<Integer> merge = Arrays.asList(1, 0, contentVerifyCategory.size(), contentVerifyCategory.size() + jsonArray.size() -1);
            mergeData.add(merge);
        }
        for (Object jsonArrayItem : jsonArray) {
            List<Object> stringList = new ArrayList<>();
            JSONArray jsonArrayItem2 = (JSONArray) jsonArrayItem;
            for (Object jsonArrayItem3 : jsonArrayItem2) {
                if ("RED".equals((String) jsonArrayItem3)
                        || "YELLOW".equals((String) jsonArrayItem3)
                        || "BLUE".equals((String) jsonArrayItem3)
                        || "GREEN".equals((String) jsonArrayItem3)) {
                    // todo 111
                    jsonArrayItem3 = new StringBuffer("/target/classes/static/image/" + ((String) jsonArrayItem3).toLowerCase() + ".jpeg");
//                    jsonArrayItem3 = new StringBuffer("/static/image/" + ((String) jsonArrayItem3).toLowerCase() + ".jpeg");
                }
                stringList.add(jsonArrayItem3);
            }
            contentVerifyCategory.add(stringList);
        }
    }

    public static void solveVarifyCategoryDetailData(JSONArray jsonArray, Integer tableRow, List<List<Object>> contentVerifyCategory, List<List<Integer>> mergeData) {
        tableRow += jsonArray.size();
        if (jsonArray.size() > 1) {
            // 身份核实第一列和第三列都需要纵向合并
            if ("身份核实".equals(((JSONArray) jsonArray.get(0)).get(0))) {
                mergeData.add(Arrays.asList(1, 0, contentVerifyCategory.size(), contentVerifyCategory.size() + jsonArray.size() -1));
                mergeData.add(Arrays.asList(1, 2, contentVerifyCategory.size(), contentVerifyCategory.size() + jsonArray.size() -1));
            } else {
                mergeData.add(Arrays.asList(1, 0, contentVerifyCategory.size(), contentVerifyCategory.size() + jsonArray.size() -1));
            }
        }
        for (Object jsonArrayItem : jsonArray) {
            if (jsonArray.size() > 1) {
                if ("个人工商信息核实".equals(((JSONArray) jsonArrayItem).get(0))) {
                    if (((String) ((JSONArray) jsonArrayItem).get(1)).contains("：担任股东记录")
                            || ((String) ((JSONArray) jsonArrayItem).get(1)).contains("：担任高管记录")
                            || ((String) ((JSONArray) jsonArrayItem).get(1)).contains("：担任法人记录")) {
                        mergeData.add(Arrays.asList(2, contentVerifyCategory.size(), 1, 3));
                    }
                    if ((((JSONArray) jsonArrayItem).get(1)).equals("企业（机构）名称")
                            || (((JSONArray) jsonArrayItem).get(1)).equals("注册号")
                            || (((JSONArray) jsonArrayItem).get(1)).equals("统一社会信用编码")
                            || (((JSONArray) jsonArrayItem).get(1)).equals("注册资本")
                            || (((JSONArray) jsonArrayItem).get(1)).equals("企业状态")
                            || (((JSONArray) jsonArrayItem).get(1)).equals("企业类型")
                            || (((JSONArray) jsonArrayItem).get(1)).equals("出资比例")) {
                        mergeData.add(Arrays.asList(2, contentVerifyCategory.size(), 2, 3));
                    }
                }
            }
            List<Object> stringList = new ArrayList<>();
            int i = 0;
            for (Object jsonArrayItem2 : (JSONArray) jsonArrayItem) {
                if ("RED".equals((String) jsonArrayItem2)
                        || "YELLOW".equals((String) jsonArrayItem2)
                        || "BLUE".equals((String) jsonArrayItem2)
                        || "GREEN".equals((String) jsonArrayItem2)) {
                    // todo 111
                    jsonArrayItem2 = new StringBuffer("/target/classes/static/image/" + ((String) jsonArrayItem2).toLowerCase() + ".jpeg");
//                    jsonArrayItem2 = new StringBuffer("/static/image/" + ((String) jsonArrayItem2).toLowerCase() + ".jpeg");
                }
                stringList.add(jsonArrayItem2);

                // 合并的单元格填充空字符串
                if (i == 1 && (((String) jsonArrayItem2).contains("：担任股东记录")
                        || ((String) jsonArrayItem2).contains("：担任高管记录")
                        || ((String) jsonArrayItem2).contains("：担任法人记录"))) {
                    stringList.add("");
                    stringList.add("");
                }
                if (i == 2 && ("企业（机构）名称".equals(((JSONArray) jsonArrayItem).get(1))
                        || "注册号".equals(((JSONArray) jsonArrayItem).get(1))
                        || "统一社会信用编码".equals(((JSONArray) jsonArrayItem).get(1))
                        || "注册资本".equals(((JSONArray) jsonArrayItem).get(1))
                        || "企业状态".equals(((JSONArray) jsonArrayItem).get(1))
                        || "企业类型".equals(((JSONArray) jsonArrayItem).get(1))
                        || "出资比例".equals(((JSONArray) jsonArrayItem).get(1)))) {
                    stringList.add("");
                }
                i++;
            }
            contentVerifyCategory.add(stringList);
        }
    }

    public static void solveDiplomaDegreeData(JSONArray jsonArray, Integer tableRow, List<List<Object>> content) {
        tableRow += jsonArray.size();
        for (Object jsonArrayItem : jsonArray) {
            List<Object> stringList = new ArrayList<>();
            for (Object jsonArrayItem2 : (JSONArray) jsonArrayItem) {
                if ("RED".equals((String) jsonArrayItem2)
                        || "YELLOW".equals((String) jsonArrayItem2)
                        || "BLUE".equals((String) jsonArrayItem2)
                        || "GREEN".equals((String) jsonArrayItem2)) {
                    // todo 111
                    jsonArrayItem2 = new StringBuffer("/target/classes/static/image/" + ((String) jsonArrayItem2).toLowerCase() + ".jpeg");
//                    jsonArrayItem2 = new StringBuffer("/static/image/" + ((String) jsonArrayItem2).toLowerCase() + ".jpeg");
                }
                stringList.add(jsonArrayItem2);
            }
            content.add(stringList);
        }
    }

    public static void solvePerfData(JSONArray jsonArray, List<List<Object>> content) {
        int i = 0;
        for (Object jsonArrayItem : jsonArray) {
            List<Object> stringList = new ArrayList<>();
            int j = 0;
            for (Object jsonArrayItem2 : (JSONArray) jsonArrayItem) {
                if ("RED".equals((String) jsonArrayItem2)
                        || "YELLOW".equals((String) jsonArrayItem2)
                        || "BLUE".equals((String) jsonArrayItem2)
                        || "GREEN".equals((String) jsonArrayItem2)) {
                    // todo 111
                    jsonArrayItem2 = new StringBuffer("/target/classes/static/image/" + ((String) jsonArrayItem2).toLowerCase() + ".jpeg");
//                    jsonArrayItem2 = new StringBuffer("/static/image/" + ((String) jsonArrayItem2).toLowerCase() + ".jpeg");
                }
                stringList.add(jsonArrayItem2);
                if (i == 1 && j == 2) {
                    stringList.add("");
                    stringList.add("");
                    stringList.add("");
                }
                if (i >= 2 && j == 0) {
                    stringList.add("");
                }
                if (i >= 2 && j == 1) {
                    stringList.add("");
                    stringList.add("");
                    stringList.add("");
                }
                j++;
            }
            content.add(stringList);
            i++;
        }
    }

    /**
     * 第一部分：基本信息及详情
     * @param document
     */
    public void buildBasicMessageAndDetail(XWPFDocument document, JSONObject params) throws IOException, URISyntaxException, InvalidFormatException {
        JSONObject basicInfoAndDetail = (JSONObject) params.get("basicInfoAndDetail");
        // 标题
        buildTitleSpecial(document, " 第一部分：基本信息及详情");
        // 空一行
        blankParagraph(document);
        // 核实类目明细表
        if (basicInfoAndDetail.getOrDefault("varifyCategoryDetail", null) != null) {
            JSONObject varifyCategoryDetail = (JSONObject) basicInfoAndDetail.get("varifyCategoryDetail");
            List<String> title = new ArrayList<>();
            title.add("核实类目明细");
            title.add("核实内容");
            title.add("核实结果");
            title.add("说明");
            // 表格行数（不包括标题）
            Integer tableRow = 0;
            // 整理出需要合并的单元格数据 格式如下：
            // [[1, 1, 1, 2]]
            // 第一个元素：1表示列合并，2表示行合并
            // 第二个元素：表示第几列或者第几行
            // 第三个元素：表示从第几行或者第几列开始
            // 第四个元素：表示从第几行或者第几列结束
            List<List<Integer>> mergeData = new ArrayList<>();
            List<List<Object>> contentVarifyCategoryDetail = new ArrayList<>();

            // 身份风险
            if (varifyCategoryDetail.getOrDefault("IDENTITY", null) != null) {
                solveVarifyCategoryDetailData((JSONArray) varifyCategoryDetail.get("IDENTITY"), tableRow, contentVarifyCategoryDetail, mergeData);
            }
            // 社会风险
            if (varifyCategoryDetail.getOrDefault("VIOLATION", null) != null) {
                solveVarifyCategoryDetailData((JSONArray) varifyCategoryDetail.get("VIOLATION"), tableRow, contentVarifyCategoryDetail, mergeData);
            }
            // 诉讼风险
            if (varifyCategoryDetail.getOrDefault("LEGAL", null) != null) {
                solveVarifyCategoryDetailData((JSONArray) varifyCategoryDetail.get("LEGAL"), tableRow, contentVarifyCategoryDetail, mergeData);
            }
            // 商业风险
            if (varifyCategoryDetail.getOrDefault("BIZ", null) != null) {
                solveVarifyCategoryDetailData((JSONArray) varifyCategoryDetail.get("BIZ"), tableRow, contentVarifyCategoryDetail, mergeData);
            }
            buildTableSpecial7(document, title, contentVarifyCategoryDetail,4, new Long[]{2000L, 3000L, 1000L, 3310L}, 400, 8310L, 1, colorGary, null, colorBlack, 10, false, ParagraphAlignment.CENTER, mergeData);
        }
        // 空一行
        blankParagraph(document);
        // 学历（可能多个）
        if (basicInfoAndDetail.getOrDefault("diploma", null) != null) {
            JSONObject diploma = (JSONObject) basicInfoAndDetail.get("diploma");
            if (diploma.getOrDefault("detail", null) != null) {
                for (Object item : (JSONArray) diploma.get("detail")) {
                    // 整理出需要合并的单元格数据 格式如下：
                    // [[1, 1, 1, 2]]
                    // 第一个元素：1表示列合并，2表示行合并
                    // 第二个元素：表示第几列或者第几行
                    // 第三个元素：表示从第几行或者第几列开始
                    // 第四个元素：表示从第几行或者第几列结束
                    List<List<Integer>> mergeData = new ArrayList<>();
                    List<String> title = new ArrayList<>();
                    title.add("中国大陆高等教育学历核实");
                    title.add("");
                    title.add("");
                    title.add("");
                    mergeData.add(Arrays.asList(2, 0, 0, 3));
                    List<List<Object>> contentDiplomaDetail = new ArrayList<>();
                    contentDiplomaDetail.add(Arrays.asList("核实内容", "", "核实结果", "说明"));
                    mergeData.add(Arrays.asList(2, 1, 0, 1));
                    // 表格行数
                    Integer tableRow = 1;
                    solveDiplomaDegreeData((JSONArray) item, tableRow, contentDiplomaDetail);
                    buildTableSpecial9(document, title, contentDiplomaDetail,4, new Long[]{2000L, 3000L, 1000L, 3310L}, 400, 8310L, 1, colorGary, null, colorBlack, 10, false, ParagraphAlignment.CENTER, mergeData);
                    // 空一行
                    blankParagraph(document);
                }
            }
        }
        // 学位（可能多个）
        if (basicInfoAndDetail.getOrDefault("degree", null) != null) {
            JSONObject degree = (JSONObject) basicInfoAndDetail.get("degree");
            if (degree.getOrDefault("detail", null) != null) {
                for (Object item : (JSONArray) degree.get("detail")) {
                    // 整理出需要合并的单元格数据 格式如下：
                    // [[1, 1, 1, 2]]
                    // 第一个元素：1表示列合并，2表示行合并
                    // 第二个元素：表示第几列或者第几行
                    // 第三个元素：表示从第几行或者第几列开始
                    // 第四个元素：表示从第几行或者第几列结束
                    List<List<Integer>> mergeData = new ArrayList<>();
                    List<String> title = new ArrayList<>();
                    title.add("中国大陆高等教育学位核实");
                    title.add("");
                    title.add("");
                    title.add("");
                    mergeData.add(Arrays.asList(2, 0, 0, 3));
                    List<List<Object>> contentDegreeDetail = new ArrayList<>();
                    contentDegreeDetail.add(Arrays.asList("核实内容", "", "核实结果", "说明"));
                    mergeData.add(Arrays.asList(2, 1, 0, 1));
                    // 表格行数
                    Integer tableRow = 1;
                    solveDiplomaDegreeData((JSONArray) item, tableRow, contentDegreeDetail);
                    buildTableSpecial9(document, title, contentDegreeDetail,4, new Long[]{2000L, 3000L, 1000L, 3310L}, 400, 8310L, 1, colorGary, null, colorBlack, 10, false, ParagraphAlignment.CENTER, mergeData);
                    // 空一行
                    blankParagraph(document);
                }
            }
        }
    }

    /**
     * 第二部分：工作履历及表现
     * @param document
     */
    public void buildXpPerfDetail(XWPFDocument document, JSONObject params) throws IOException, URISyntaxException, InvalidFormatException {
        JSONObject xpAndPerf = (JSONObject) params.get("xpAndPerf");
        // 标题
        buildTitleSpecial(document, " 第二部分：工作履历及表现");
        // 空一行
        blankParagraph(document);
        // 核实类目明细表
        if (xpAndPerf.getOrDefault("xpAndPerfVerifyCategoryDetail", null) != null) {
            JSONArray xpAndPerfVerifyCategoryDetail = (JSONArray) xpAndPerf.get("xpAndPerfVerifyCategoryDetail");
            if (xpAndPerfVerifyCategoryDetail.size() > 0) {
                // 整理出需要合并的单元格数据 格式如下：
                // [[1, 1, 1, 2]]
                // 第一个元素：1表示列合并，2表示行合并
                // 第二个元素：表示第几列或者第几行
                // 第三个元素：表示从第几行或者第几列开始
                // 第四个元素：表示从第几行或者第几列结束
                List<List<Integer>> mergeData = new ArrayList<>();
                List<List<Object>> contentXpAndPerfDetail = new ArrayList<>();
                // 表格行数（不包括标题）
                Integer tableRow = 0;
                if (xpAndPerfVerifyCategoryDetail.size() % 10 > 0) {
                    throw new RuntimeException("报告数据错误 xpAndPerfVerifyCategoryDetail=" + xpAndPerfVerifyCategoryDetail.toJSONString());
                }
                for (int i = 0; i < xpAndPerfVerifyCategoryDetail.size() / 10; i++) {
                    mergeData.add(Arrays.asList(1, 0, 1 + 10 * i, 10 * (i + 1)));
                }
                List<String> title = new ArrayList<>();
                title.add("核实类目明细");
                title.add("核实内容");
                title.add("核实结果");
                title.add("说明");
                solveDiplomaDegreeData(xpAndPerfVerifyCategoryDetail, tableRow, contentXpAndPerfDetail);
                buildTableSpecial10(document, title, contentXpAndPerfDetail,4, new Long[]{2000L, 3000L, 1000L, 3310L}, 400, 8310L, 1, colorGary, null, colorBlack, 10, false, ParagraphAlignment.CENTER, mergeData);
            }
            // 空一行
            blankParagraph(document);
        }

        // 履历、证明人、表现（可能多个）
        if (xpAndPerf.getOrDefault("xpPerf", null) != null) {
            JSONObject xpPerf = (JSONObject) xpAndPerf.get("xpPerf");
            if (xpPerf.getOrDefault("detail", null) != null) {
                for (Object item : (JSONArray) xpPerf.get("detail")) {
                    // 履历表格（可能多个）
                    JSONObject xpVerify = (JSONObject) ((JSONObject) item).get("xpVerify");
                    for (Object xpDetailItem : (JSONArray) xpVerify.get("detail")) {
                        List<List<Integer>> mergeData = new ArrayList<>();
                        List<String> title = new ArrayList<>();
                        title.add((String) ((JSONObject) xpDetailItem).get("title"));
                        title.add("");
                        title.add("");
                        title.add("");
                        mergeData.add(Arrays.asList(2, 0, 0, 3));
                        List<List<Object>> contentXpDetail = new ArrayList<>();
                        contentXpDetail.add(Arrays.asList("核实内容", "", "核实结果", "说明"));
                        mergeData.add(Arrays.asList(2, 1, 0, 1));

                        // 合并入职时间、离职时间的判灯
                        mergeData.add(Arrays.asList(1, 2, 3, 4));
                        // 表格行数
                        Integer tableRow = 1;
                        solveDiplomaDegreeData((JSONArray) ((JSONObject) xpDetailItem).get("content"), tableRow, contentXpDetail);
                        buildTableSpecial9(document, title, contentXpDetail,4, new Long[]{2000L, 3000L, 1000L, 3310L}, 400, 8310L, 1, colorGary, null, colorBlack, 10, false, ParagraphAlignment.CENTER, mergeData);
                        // 空一行
                        blankParagraph(document);
                    }


                    // 证明人表格（单个）
                    JSONObject certifier = (JSONObject) ((JSONObject) item).get("certifier");
                    List<List<Integer>> mergeDataCertifier = new ArrayList<>();
                    List<String> titleCertifier = new ArrayList<>();
                    titleCertifier.add((String) certifier.get("title"));
                    titleCertifier.add("");
                    titleCertifier.add("");
                    titleCertifier.add("");
                    titleCertifier.add("");
                    titleCertifier.add("");
                    titleCertifier.add("");
                    mergeDataCertifier.add(Arrays.asList(2, 0, 0, 6));
                    List<List<Object>> contentCertifierDetail = new ArrayList<>();
                    contentCertifierDetail.add(Arrays.asList("证明人", "联系方式", "职位", "与候选人关系", "共事时长", "来源", "真实性"));
                    // 表格行数
                    Integer tableRowCertifier = 1;
                    solveDiplomaDegreeData((JSONArray) certifier.get("content"), tableRowCertifier, contentCertifierDetail);
                    // 填充说明内容
                    contentCertifierDetail.add(Arrays.asList(certifier.get("explanation"), "", "", "", "", "", ""));
                    // 说明内容的合并单元格
                    mergeDataCertifier.add(Arrays.asList(2, 2 + ((JSONArray) certifier.get("content")).size(), 0, 6));
                    buildTableSpecial11(document, titleCertifier, contentCertifierDetail,7, new Long[]{900L, 1400L, 900L, 1400L, 1400L, 1400L, 910L}, 400, 8310L, 1, colorGary, null, colorBlack, 10, false, ParagraphAlignment.CENTER, mergeDataCertifier);
                    // 空一行
                    blankParagraph(document);

                    // 表现表格（可能多个）
                    JSONObject perfVerify = (JSONObject) ((JSONObject) item).get("perfVerify");
                    for (Object perfDetailItem : (JSONArray) perfVerify.get("detail")) {
                        if (((JSONArray) ((JSONObject) perfDetailItem).get("content")).size() != 13) {
                            throw new RuntimeException("报告数据异常 perfVerify=" + perfVerify.toJSONString());
                        }
                        List<List<Integer>> mergeData = new ArrayList<>();
                        List<String> title = new ArrayList<>();
                        title.add((String) ((JSONObject) perfDetailItem).get("title"));
                        title.add("");
                        title.add("");
                        title.add("");
                        title.add("");
                        title.add("");
                        mergeData.add(Arrays.asList(2, 0, 0, 5));
                        List<List<Object>> contentPerfDetail = new ArrayList<>();
                        // 合并单元格
                        mergeData.add(Arrays.asList(2, 2, 2, 5));
                        mergeData.add(Arrays.asList(2, 3, 0, 1));
                        mergeData.add(Arrays.asList(2, 4, 0, 1));
                        mergeData.add(Arrays.asList(2, 5, 0, 1));
                        mergeData.add(Arrays.asList(2, 6, 0, 1));
                        mergeData.add(Arrays.asList(2, 7, 0, 1));
                        mergeData.add(Arrays.asList(2, 8, 0, 1));
                        mergeData.add(Arrays.asList(2, 9, 0, 1));
                        mergeData.add(Arrays.asList(2, 10, 0, 1));
                        mergeData.add(Arrays.asList(2, 11, 0, 1));
                        mergeData.add(Arrays.asList(2, 12, 0, 1));
                        mergeData.add(Arrays.asList(2, 13, 0, 1));

                        mergeData.add(Arrays.asList(2, 2, 2, 5));
                        mergeData.add(Arrays.asList(2, 3, 2, 5));
                        mergeData.add(Arrays.asList(2, 4, 2, 5));
                        mergeData.add(Arrays.asList(2, 5, 2, 5));
                        mergeData.add(Arrays.asList(2, 6, 2, 5));
                        mergeData.add(Arrays.asList(2, 7, 2, 5));
                        mergeData.add(Arrays.asList(2, 8, 2, 5));
                        mergeData.add(Arrays.asList(2, 9, 2, 5));
                        mergeData.add(Arrays.asList(2, 10, 2, 5));
                        mergeData.add(Arrays.asList(2, 11, 2, 5));
                        mergeData.add(Arrays.asList(2, 12, 2, 5));
                        mergeData.add(Arrays.asList(2, 13, 2, 5));
                        solvePerfData((JSONArray) ((JSONObject) perfDetailItem).get("content"), contentPerfDetail);
                        buildTableSpecial10(document, title, contentPerfDetail,6, new Long[]{2000L, 1000L, 1410L, 1000L, 1400L, 1500L}, 400, 8310L, 1, colorGary, null, colorBlack, 10, false, ParagraphAlignment.CENTER, mergeData);
                        // 空一行
                        blankParagraph(document);
                    }
                }
            }
        }
    }

    /**
     * 候选人、委托日期等信息的表格
     */
    public void buildWTXXAndBGGLMainTable(XWPFDocument document, JSONArray deliverInfo) throws InvalidFormatException, IOException, URISyntaxException {
        List<List<Object>> content = new ArrayList<>();
        for (Object item: deliverInfo) {
            List<Object> strings = new ArrayList<>();
            int i = 0;
            for (Object object : (List) item) {
                if ("RED".equals((String) object)
                        || "YELLOW".equals((String) object)
                        || "BLUE".equals((String) object)
                        || "GREEN".equals((String) object)) {
                    // todo 111
                    object = new StringBuffer("/target/classes/static/image/" + ((String) object).toLowerCase() + ".jpeg");
//                    object = new StringBuffer("/static/image/" + ((String) object).toLowerCase() + ".jpeg");
                }
                strings.add(object);
                if (i == 1 || i == 3) {
                    strings.add("");
                }
                i++;
            }
            content.add(strings);
        }
        buildTableSpecial3(document, null, content,6, new Long[]{1200L, 2500L, 310L, 1200L, 2500L, 600L}, 400, 8310L, true, null, null, 10, false);
    }

//    public void builderCell(XWPFTableCell cell, String content, String backgroundColor, ParagraphAlignment paragraphAlign, Integer width, Integer fontSize, String filePath) throws IOException, InvalidFormatException, URISyntaxException {
//        XWPFParagraph paragraph = cell.getParagraphs().get(0);
//        if(StringUtils.isNotBlank(filePath)) {
//            //图片
//            XWPFRun pictureRun = paragraph.createRun();
//            FileInputStream is = null;
//            try {
//                // filePath = /static/image/image1.png
//                is = new FileInputStream(new File(this.getClass().getResource(filePath).toURI()));
//                pictureRun.addPicture(is, Document.PICTURE_TYPE_JPEG, "c1.png", Units.toEMU(120), Units.toEMU(30));
//            } catch (IOException | InvalidFormatException | URISyntaxException e) {
//                throw e;
//            } finally {
//                if (is != null) {
//                    is.close();
//                }
//            }
//        } else {
//            CTTc cttc = cell.getCTTc();
//            CTTcPr ctPr = cttc.addNewTcPr();
//            /** 背景色 */
//            cell.setColor(backgroundColor);
//            /** 水平居中 */
//            cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//            /** 竖直居中 */
//            ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
//            cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
//            /**单元格宽度*/
//            CTTblWidth ctTblWidthCell = ctPr.addNewTcW();
//            ctTblWidthCell.setType(STTblWidth.DXA);
//            ctTblWidthCell.setW(BigInteger.valueOf(width));
//
//            XWPFParagraph paragraph1 = cell.getParagraphs().get(0);
//            buildParagraph(paragraph1, paragraphAlign, content, fontSize, null, colorBlack);
//        }
//
//        cell.setParagraph(paragraph);
//    }

    private void addBlank(XWPFDocument document) {
        try{
            XWPFParagraph blankp = document.createParagraph();
            XWPFRun xwpfRun = blankp.createRun();
            xwpfRun.addBreak();
            document.setParagraph(blankp, 50);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 报告首页
     */

    public void buildHomePage(XWPFDocument document, JSONObject params) throws IOException, InvalidFormatException, URISyntaxException {
        // 获取首页数据
        JSONObject homePage = (JSONObject) params.get("homePage");
        blankParagraph(document);
        blankParagraph(document);
        /**
         * 雇前背景调查报告
         */
        List<List<Object>> content = new ArrayList<>();
        content.add(Arrays.asList("雇 前 背 景 调 查 报 告"));
        buildTableSpecial(document, null, content,1, new Long[]{8310L}, 600, 8130L, 0, null, null, colorBlue, 28, true, ParagraphAlignment.CENTER);

        blankParagraph(document);

        /**
         * 委托日期
         */
        List<List<Object>> content2 = new ArrayList<>();
        content2.add(Arrays.asList("委托日期：" + homePage.getString("entrustAt")));
        // 背景色 浅蓝色
        buildTableSpecial(document, null, content2,1, new Long[]{4000L}, 400, 4000L, 0, null, colorBlue2, colorWrite, 14, true, ParagraphAlignment.CENTER);

        // 空一行
        blankParagraph(document);
        blankParagraph(document);
        /**
         * 公司名称、委托方名称、报告编号
         */
        //标题
        //标题字体风格
        //设置中文字体
        try {

            XWPFParagraph titlep = document.createParagraph();
//            XWPFRun titlerun = titlep.createRun();
            Integer fontSize = new Integer(14);
            Boolean bold = false;
            String color = "000000";
            buildParagraphSpecial(titlep, ParagraphAlignment.CENTER, homePage.getString("fromName"), fontSize, bold, color);
            document.setParagraph(titlep, 100);
        } catch (Exception e) {
            e.printStackTrace();
        }
        blankParagraph(document);

        try {

            XWPFParagraph titlep = document.createParagraph();
//            XWPFRun titlerun = titlep.createRun();
            Integer fontSize = new Integer(14);
            Boolean bold = false;
            String color = "000000";
            buildParagraphSpecial(titlep, ParagraphAlignment.CENTER, homePage.getString("candidateName"), fontSize, bold, color);
            document.setParagraph(titlep, 100);
        } catch (Exception e) {
            e.printStackTrace();
        }
        blankParagraph(document);
        try {

            XWPFParagraph titlep = document.createParagraph();
//            XWPFRun titlerun = titlep.createRun();
            Integer fontSize = new Integer(14);
            Boolean bold = false;
            String color = "000000";
            buildParagraphSpecial(titlep, ParagraphAlignment.CENTER, "报告编号：" + homePage.getString("zjReportSerialNumber"), fontSize, bold, color);
            document.setParagraph(titlep, 100);
        } catch (Exception e) {
            e.printStackTrace();
        }
        blankParagraph(document);
        blankParagraph(document);
//        addBlank(document);

        try {

            XWPFParagraph titlep = document.createParagraph();
//            XWPFRun titlerun = titlep.createRun();
            Integer fontSize = new Integer(14);
            Boolean bold = false;
            String color = "EB382A";
            System.out.println("1");
            buildParagraphSpecial(titlep, ParagraphAlignment.CENTER, "<内部保密文件>", fontSize, bold, color);
            document.setParagraph(titlep, 100);
        } catch (Exception e) {
            e.printStackTrace();
        }
        blankParagraph(document);

        /**
         * L4级机密、禁止分享、限期删除
         */
        List<Object> list = new ArrayList<>();
        list.add("");
        // todo 111
        list.add(new StringBuffer("/target/classes/static/image/1-L4级机密.jpeg"));
//        list.add(new StringBuffer("/static/image/1-L4级机密.jpeg"));
        list.add("L4级机密");
        list.add("");
        // todo 111
        list.add(new StringBuffer("/target/classes/static/image/2-禁止分享.jpeg"));
//        list.add(new StringBuffer("/static/image/2-禁止分享.jpeg"));
        list.add("禁止分享");
        list.add("");
        // todo 111
        list.add(new StringBuffer("/target/classes/static/image/2-禁止分享.jpeg"));
//        list.add(new StringBuffer("/static/image/2-禁止分享.jpeg"));
        list.add("限期删除");
        list.add("");
        List<List<Object>> content5 = new ArrayList();
        content5.add(list);
        buildTableSpecial2(document, null, content5,10, new Long[]{1655L, 300L, 1100L, 400L, 300L, 1100L, 400L, 300L, 1100L, 1655L}, 200, 8310L, true, null, colorPink, colorRed, 10, true);

    }

    /**
     * 分页
     *
     * @param document
     */
    public void pageBreakSpecial(XWPFDocument document) {
        XWPFParagraph p = document.createParagraph();
        p.setPageBreak(true);
    }

    /**
     * word跨列并单元格
     *
     * @param table
     * @param row
     * @param fromCell
     * @param toCell
     */
    public void mergeCellsHorizontalSpecial(XWPFTable table, int row, int fromCell, int toCell) {
        for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
            XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
            if (cellIndex == fromCell) {
                // The first merged cell is set with RESTART merge value
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    /**
     * word跨行并单元格
     *
     * @param table
     * @param col
     * @param fromRow
     * @param toRow
     */
    public void mergeCellsVerticallySpecial(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            if (rowIndex == fromRow) {
                // The first merged cell is set with RESTART merge value
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    public void buildTableGaoDengXueLi(XWPFDocument document) throws Exception {
        //空一行
        blankParagraph(document);

        int rowNum = 12;
        int columnNum = 4;
        XWPFTable xwpfTable = document.createTable(rowNum, columnNum);
        //表格居中显示
        CTTblPr ctTblPr = xwpfTable.getCTTbl().addNewTblPr();
        ctTblPr.addNewJc().setVal(STJc.CENTER);
        //设置表格宽度
        CTTblWidth ctTblWidth = ctTblPr.addNewTblW();
        ctTblWidth.setW(BigInteger.valueOf(9000));
        //设置表格宽度为非自动
        ctTblWidth.setType(STTblWidth.DXA);
        // 自定义边框：上下左右有灰色边框
        customizeBorderSpecial(xwpfTable, colorGary);

        //创建标题
        //标题行对象
        XWPFTableRow xwpfTableRowTitle = xwpfTable.getRow(0);
        //标题行高
        xwpfTableRowTitle.setHeight(200);
        //单元格对象
        XWPFTableCell xwpfTableCell = xwpfTableRowTitle.getCell(0);
        //标题行
        String titletest = "中国大陆高等教育\r\n" +
                "学历核实";
        buildCellSpecialWithTextWrap(xwpfTableCell, titletest, 12, true, colorBlack, ParagraphAlignment.CENTER, 200L, null, "E5E5E5");
        mergeCellsHorizontalSpecial(xwpfTable, 0, 0, 3);
        //子标题
        XWPFTableRow xwpfTableRowTitle1 = xwpfTable.getRow(1);
        XWPFTableCell xwpfTableCell1_0 = xwpfTableRowTitle1.getCell(0);
        buildCellSpecial(xwpfTableCell1_0, "核实内容", 12, true, colorBlack, ParagraphAlignment.CENTER, 2250L, null, null);
        XWPFTableCell xwpfTableCell1_2 = xwpfTableRowTitle1.getCell(2);
        buildCellSpecial(xwpfTableCell1_2, "核实结果", 12, true, colorBlack, ParagraphAlignment.CENTER, 2250L, null, null);
        XWPFTableCell xwpfTableCell1_3 = xwpfTableRowTitle1.getCell(3);
        buildCellSpecial(xwpfTableCell1_3, "说明", 12, true, colorBlack, ParagraphAlignment.CENTER, 2250L, null, null);
        mergeCellsHorizontalSpecial(xwpfTable, 1, 0, 1);

    }

    /**
     * 创建表格
     * @param document
     * @param title 表格标题
     * @param content 表格内容
     * @param numColumn 表格列数，如果为null，则取数组tableWidths的长度为列数
     * @param tableWidths 数组类型，表格每列的宽度
     * @param rowsHeight 数组类型，表格每行的宽度
     * @param tableWidth 表格宽度
     * @param displayBorder 是否隐藏边框 0:不展示，1：展示，2：自定义边框（上下左右有边框）
     * @param titleBackground 表格标题背景色
     * @param tableBackground 表格背景色
     * @param wordColor 字体颜色
     * @param fontSize 字体大小
     * @param bold 字体是否加粗
     * @param align 对齐方式
     * @throws InvalidFormatException
     * @throws IOException
     * @throws URISyntaxException
     */
    public void buildTableSpecial(XWPFDocument document
            , List<String> title
            , List<List<Object>> content
            , Integer numColumn
            , Long[] tableWidths
            , Integer rowsHeight
            , Long tableWidth
            , Integer displayBorder
            , String titleBackground
            , String tableBackground
            , String wordColor
            , Integer fontSize
            , Boolean bold
            , ParagraphAlignment align) throws InvalidFormatException, IOException, URISyntaxException {
        int rowNum = content.size();
        int columnNum = numColumn == null ? tableWidths.length : numColumn;
        if (CollectionUtils.isNotEmpty(title)) {
            rowNum++;
        }
        XWPFTable xwpfTable = document.createTable(rowNum, columnNum);
        //表格居中显示
        CTTblPr ctTblPr = xwpfTable.getCTTbl().addNewTblPr();
        ctTblPr.addNewJc().setVal(STJc.CENTER);
        //设置表格宽度
        CTTblWidth ctTblWidth = ctTblPr.addNewTblW();
        ctTblWidth.setW(BigInteger.valueOf(tableWidth));
        //设置表格宽度为非自动
        ctTblWidth.setType(STTblWidth.DXA);
        Long cellWidth = null;
        //设置边框
        if (displayBorder == 0) {
            displayBorder(xwpfTable);
        } else if(displayBorder == 2) {
            // 自定义边框：上下左右有灰色边框
            customizeBorderSpecial(xwpfTable, colorGary);
        }

        //创建标题
        if (CollectionUtils.isNotEmpty(title)) {
            //标题行对象
            XWPFTableRow xwpfTableRowTitle = xwpfTable.getRow(0);
            //标题行高
            xwpfTableRowTitle.setHeight(rowsHeight);
            for (int i = 0; i < title.size() && i < columnNum; i++) {
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                String msg = title.get(i);
                //单元格对象
                XWPFTableCell xwpfTableCell = xwpfTableRowTitle.getCell(i);
                //添加文本，9号/黑体/黑色/居中
                buildCellSpecial(xwpfTableCell, msg, 9, true, colorBlack, align, cellWidth, null, titleBackground);
                // 添加边框
//                addBottomBorder(xwpfTableCell, 1, colorBlue, true);
            }
        }

        //创建内容
        for (int j = 0; j < content.size(); j++) {
            //有标题时和无标题时取值位置不同
            XWPFTableRow xwpfTableRow;
            if (CollectionUtils.isNotEmpty(title)) {
                // 行对象
                xwpfTableRow = xwpfTable.getRow(j + 1);
                //行高
                xwpfTableRow.setHeight(rowsHeight);
            } else {
                // 行对象
                xwpfTableRow = xwpfTable.getRow(j);
                //行高
                xwpfTableRow.setHeight(rowsHeight);
            }
            List<Object> row = content.get(j);
            for (int i = 0; i < row.size() && i < columnNum; i++) {
                //单元格对象
                XWPFTableCell xwpfTableCell = xwpfTableRow.getCell(i);
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                //添加文本，14号/黑体/左对齐
                buildCellSpecial(xwpfTableCell, row.get(i), fontSize, bold, wordColor, align, cellWidth, null, tableBackground);
            }
        }
    }

    /**
     * 创建表格（单元格有规律添加背景色）
     * @param document
     * @param title 表格标题
     * @param content 表格内容
     * @param numColumn 表格列数，如果为null，则取数组tableWidths的长度为列数
     * @param tableWidths 数组类型，表格每列的宽度
     * @param rowsHeight 数组类型，表格每行的宽度
     * @param tableWidth 表格宽度
     * @param displayBorder 是否隐藏边框 true为隐藏
     * @param titleBackground 表格标题背景色
     * @param tableBackground 表格背景色
     * @param wordColor 字体颜色
     * @param fontSize 字体大小
     * @param bold 字体是否加粗
     * @throws InvalidFormatException
     * @throws IOException
     * @throws URISyntaxException
     */
    public void buildTableSpecial2(XWPFDocument document
            , List<String> title
            , List<List<Object>> content
            , Integer numColumn
            , Long[] tableWidths
            , Integer rowsHeight
            , Long tableWidth
            , Boolean displayBorder
            , String titleBackground
            , String tableBackground
            , String wordColor
            , Integer fontSize
            , Boolean bold) throws InvalidFormatException, IOException, URISyntaxException {
        int rowNum = content.size();
        int columnNum = numColumn == null ? tableWidths.length : numColumn;
        if (CollectionUtils.isNotEmpty(title)) {
            rowNum++;
        }
        XWPFTable xwpfTable = document.createTable(rowNum, columnNum);
        //表格居中显示
        CTTblPr ctTblPr = xwpfTable.getCTTbl().addNewTblPr();
        ctTblPr.addNewJc().setVal(STJc.CENTER);
        //设置表格宽度
        CTTblWidth ctTblWidth = ctTblPr.addNewTblW();
        ctTblWidth.setW(BigInteger.valueOf(tableWidth));
        //设置表格宽度为非自动
        ctTblWidth.setType(STTblWidth.DXA);
        Long cellWidth = null;
        if (displayBorder) {
            // 设置无边框
            displayBorder(xwpfTable);
        }

        //创建标题
        if (CollectionUtils.isNotEmpty(title)) {
            //标题行对象
            XWPFTableRow xwpfTableRowTitle = xwpfTable.getRow(0);
            //标题行高
            xwpfTableRowTitle.setHeight(rowsHeight);
            for (int i = 0; i < title.size() && i < columnNum; i++) {
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                String msg = title.get(i);
                //单元格对象
                XWPFTableCell xwpfTableCell = xwpfTableRowTitle.getCell(i);
                //添加文本，9号/黑体/黑色/居中
                buildCellSpecial(xwpfTableCell, msg, 9, true, colorBlack, ParagraphAlignment.CENTER, cellWidth, null, titleBackground);
                // 添加边框
//                addBottomBorder(xwpfTableCell, 1, colorBlue, true);
            }
        }

        //创建内容
        for (int j = 0; j < content.size(); j++) {
            //有标题时和无标题时取值位置不同
            XWPFTableRow xwpfTableRow;
            if (CollectionUtils.isNotEmpty(title)) {
                // 行对象
                xwpfTableRow = xwpfTable.getRow(j + 1);
                //行高
                xwpfTableRow.setHeight(rowsHeight);
            } else {
                // 行对象
                xwpfTableRow = xwpfTable.getRow(j);
                //行高
                xwpfTableRow.setHeight(rowsHeight);
            }
            List<Object> row = content.get(j);
            for (int i = 0; i < row.size() && i < columnNum; i++) {
                //单元格对象
                XWPFTableCell xwpfTableCell = xwpfTableRow.getCell(i);
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                //黑字 居中
//                String wordColor = colorBlack;
                ParagraphAlignment align = ParagraphAlignment.CENTER;
                if (i == 2 || i == 5 || i == 8) {
                    buildCellSpecial(xwpfTableCell, row.get(i), fontSize, bold, wordColor, align, cellWidth, null, tableBackground);
                } else {
                    buildCellSpecial(xwpfTableCell, row.get(i), fontSize, bold, wordColor, align, cellWidth, null, null);
                }
            }
        }
    }

    /**
     * 委托信息 / 报告概览
     * @param document
     * @param title
     * @param content
     * @param numColumn
     * @param tableWidths
     * @param rowsHeight
     * @param tableWidth
     * @param displayBorder
     * @param titleBackground
     * @param wordColor
     * @param fontSize
     * @param bold
     * @throws InvalidFormatException
     * @throws IOException
     * @throws URISyntaxException
     */
    public void buildTableSpecial3(XWPFDocument document
            , List<String> title
            , List<List<Object>> content
            , Integer numColumn
            , Long[] tableWidths
            , Integer rowsHeight
            , Long tableWidth
            , Boolean displayBorder
            , String titleBackground
            , String wordColor
            , Integer fontSize
            , Boolean bold) throws InvalidFormatException, IOException, URISyntaxException {
        int rowNum = content.size();
        int columnNum = numColumn == null ? tableWidths.length : numColumn;
        if (CollectionUtils.isNotEmpty(title)) {
            rowNum++;
        }
        XWPFTable xwpfTable = document.createTable(rowNum, columnNum);
        //表格居中显示
        CTTblPr ctTblPr = xwpfTable.getCTTbl().addNewTblPr();
        ctTblPr.addNewJc().setVal(STJc.CENTER);
        //设置表格宽度
        CTTblWidth ctTblWidth = ctTblPr.addNewTblW();
        ctTblWidth.setW(BigInteger.valueOf(tableWidth));
        //设置表格宽度为非自动
        ctTblWidth.setType(STTblWidth.DXA);
        Long cellWidth = null;
        if (displayBorder) {
            // 设置无边框
            displayBorder(xwpfTable);
        }

        //创建标题
        if (CollectionUtils.isNotEmpty(title)) {
            //标题行对象
            XWPFTableRow xwpfTableRowTitle = xwpfTable.getRow(0);
            //标题行高
            xwpfTableRowTitle.setHeight(rowsHeight);
            for (int i = 0; i < title.size() && i < columnNum; i++) {
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                String msg = title.get(i);
                //单元格对象
                XWPFTableCell xwpfTableCell = xwpfTableRowTitle.getCell(i);
                //添加文本，9号/黑体/黑色/居中
                buildCellSpecial(xwpfTableCell, msg, 9, true, colorBlack, ParagraphAlignment.CENTER, cellWidth, null, titleBackground);
                // 添加边框
//                addBottomBorder(xwpfTableCell, 1, colorBlue, true);
            }
        }

        //创建内容
        for (int j = 0; j < content.size(); j++) {
            //有标题时和无标题时取值位置不同
            XWPFTableRow xwpfTableRow;
            if (CollectionUtils.isNotEmpty(title)) {
                // 行对象
                xwpfTableRow = xwpfTable.getRow(j + 1);
                //行高
                xwpfTableRow.setHeight(rowsHeight);
            } else {
                // 行对象
                xwpfTableRow = xwpfTable.getRow(j);
                //行高
                xwpfTableRow.setHeight(rowsHeight);
            }
            List<Object> row = content.get(j);
            for (int i = 0; i < row.size() && i < columnNum; i++) {
                //单元格对象
                XWPFTableCell xwpfTableCell = xwpfTableRow.getCell(i);
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                //居左
                ParagraphAlignment align = ParagraphAlignment.LEFT;

                // 添加背景色
                String background = "";
                if (i == 0 || i == 1) {
                    background = colorGary;
                } else if (i == 3 || i == 4) {
                    background = colorPink;
                } else {
                    background = null;
                }
                buildCellSpecial2(xwpfTableCell, row.get(i), fontSize, bold, wordColor, align, cellWidth, null, background);
            }
        }
    }

    /**
     * 创建表格（合并单元格）
     * @param document
     * @param title 表格标题
     * @param content 表格内容
     * @param numColumn 表格列数，如果为null，则取数组tableWidths的长度为列数
     * @param tableWidths 数组类型，表格每列的宽度
     * @param rowsHeight 数组类型，表格每行的宽度
     * @param tableWidth 表格宽度
     * @param displayBorder 是否隐藏边框 0:不展示，1：展示，2：自定义边框（上下左右有边框）
     * @param titleBackground 表格标题背景色
     * @param tableBackground 表格背景色
     * @param wordColor 字体颜色
     * @param fontSize 字体大小
     * @param bold 字体是否加粗
     * @param align 对齐方式
     * @throws InvalidFormatException
     * @throws IOException
     * @throws URISyntaxException
     */
    public void buildTableSpecial4_shenfenheshi(XWPFDocument document
            , List<String> title
            , List<List<Object>> content
            , Integer numColumn
            , Long[] tableWidths
            , Integer rowsHeight
            , Long tableWidth
            , Integer displayBorder
            , String titleBackground
            , String tableBackground
            , String wordColor
            , Integer fontSize
            , Boolean bold
            , ParagraphAlignment align) throws InvalidFormatException, IOException, URISyntaxException {
        int rowNum = content.size();
        int columnNum = numColumn == null ? tableWidths.length : numColumn;
        if (CollectionUtils.isNotEmpty(title)) {
            rowNum++;
        }
        XWPFTable xwpfTable = document.createTable(rowNum, columnNum);
        //表格居中显示
        CTTblPr ctTblPr = xwpfTable.getCTTbl().addNewTblPr();
        ctTblPr.addNewJc().setVal(STJc.CENTER);
        //设置表格宽度
        CTTblWidth ctTblWidth = ctTblPr.addNewTblW();
        ctTblWidth.setW(BigInteger.valueOf(tableWidth));
        //设置表格宽度为非自动
        ctTblWidth.setType(STTblWidth.DXA);
        Long cellWidth = null;
        //设置边框
        if (displayBorder == 0) {
            displayBorder(xwpfTable);
        } else if(displayBorder == 2) {
            // 自定义边框：上下左右有灰色边框
            customizeBorderSpecial(xwpfTable, colorGary);
        }

        // 跨行合并，合并第一列中第一行至第二行
        mergeCellsVerticallySpecial(xwpfTable, 0, 0, 1);
        // 跨行合并，合并第三列中第一行至第二行
        mergeCellsVerticallySpecial(xwpfTable, 2, 0, 1);


        //创建标题
        if (CollectionUtils.isNotEmpty(title)) {
            //标题行对象
            XWPFTableRow xwpfTableRowTitle = xwpfTable.getRow(0);
            //标题行高
            xwpfTableRowTitle.setHeight(rowsHeight);
            for (int i = 0; i < title.size() && i < columnNum; i++) {
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                String msg = title.get(i);
                //单元格对象
                XWPFTableCell xwpfTableCell = xwpfTableRowTitle.getCell(i);
                //添加文本，9号/黑体/黑色/居中
                buildCellSpecial(xwpfTableCell, msg, 9, true, colorBlack, align, cellWidth, null, titleBackground);
                // 添加边框
//                addBottomBorder(xwpfTableCell, 1, colorBlue, true);
            }
        }

        //创建内容
        for (int j = 0; j < content.size(); j++) {
            //有标题时和无标题时取值位置不同
            XWPFTableRow xwpfTableRow;
            if (CollectionUtils.isNotEmpty(title)) {
                // 行对象
                xwpfTableRow = xwpfTable.getRow(j + 1);
                //行高
                xwpfTableRow.setHeight(rowsHeight);
            } else {
                // 行对象
                xwpfTableRow = xwpfTable.getRow(j);
                //行高
                xwpfTableRow.setHeight(rowsHeight);
            }
            List<Object> row = content.get(j);
            for (int i = 0; i < row.size() && i < columnNum; i++) {
                //单元格对象
                XWPFTableCell xwpfTableCell = xwpfTableRow.getCell(i);
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                //添加文本，14号/黑体/左对齐
                buildCellSpecial(xwpfTableCell, row.get(i), fontSize, bold, wordColor, align, cellWidth, null, tableBackground);
            }
        }
    }

    /**
     * 创建表格
     * @param document
     * @param title 表格标题
     * @param content 表格内容
     * @param numColumn 表格列数，如果为null，则取数组tableWidths的长度为列数
     * @param tableWidths 数组类型，表格每列的宽度
     * @param rowsHeight 数组类型，表格每行的宽度
     * @param tableWidth 表格宽度
     * @param displayBorder 是否隐藏边框 0:不展示，1：展示，2：自定义边框（上下左右有边框）
     * @param titleBackground 表格标题背景色
     * @param tableBackground 表格背景色
     * @param wordColor 字体颜色
     * @param fontSize 字体大小
     * @param bold 字体是否加粗
     * @param align 对齐方式
     * @param mergeData 合并单元格数据
     * @throws InvalidFormatException
     * @throws IOException
     * @throws URISyntaxException
     */
    public void buildTableSpecial5(XWPFDocument document
            , List<String> title
            , List<List<Object>> content
            , Integer numColumn
            , Long[] tableWidths
            , Integer rowsHeight
            , Long tableWidth
            , Integer displayBorder
            , String titleBackground
            , String tableBackground
            , String wordColor
            , Integer fontSize
            , Boolean bold
            , ParagraphAlignment align
            , List<List<Integer>> mergeData) throws InvalidFormatException, IOException, URISyntaxException {
        int rowNum = content.size();
        int columnNum = numColumn == null ? tableWidths.length : numColumn;
        if (CollectionUtils.isNotEmpty(title)) {
            rowNum++;
        }
        XWPFTable xwpfTable = document.createTable(rowNum, columnNum);
        //表格居中显示
        CTTblPr ctTblPr = xwpfTable.getCTTbl().addNewTblPr();
        ctTblPr.addNewJc().setVal(STJc.CENTER);
        //设置表格宽度
        CTTblWidth ctTblWidth = ctTblPr.addNewTblW();
        ctTblWidth.setW(BigInteger.valueOf(tableWidth));
        //设置表格宽度为非自动
        ctTblWidth.setType(STTblWidth.DXA);
        Long cellWidth = null;
        //设置边框
        if (displayBorder == 0) {
            displayBorder(xwpfTable);
        } else if(displayBorder == 2) {
            // 自定义边框：上下左右有灰色边框
            customizeBorderSpecial(xwpfTable, colorGary);
        }

        // 合并单元格
        if (CollectionUtils.isNotEmpty(mergeData)) {
            for (List<Integer> item : mergeData) {
                if (item.get(0) == 1) {
                    if (CollectionUtils.isNotEmpty(title)) {
                        mergeCellsVerticallySpecial(xwpfTable, item.get(1), item.get(2) + 1, item.get(3) + 1);
                    } else {
                        mergeCellsVerticallySpecial(xwpfTable, item.get(1), item.get(2), item.get(3));
                    }
                }
                if (item.get(0) == 2) {
                    if (CollectionUtils.isNotEmpty(title)) {
                        mergeCellsHorizontalSpecial(xwpfTable, item.get(1) + 1, item.get(2), item.get(3));
                    } else {
                        mergeCellsHorizontalSpecial(xwpfTable, item.get(1), item.get(2), item.get(3));
                    }
                }
            }
        }

        //创建标题
        //标题行对象
        XWPFTableRow xwpfTableRowTitle = xwpfTable.getRow(0);
        //标题行高
        xwpfTableRowTitle.setHeight(rowsHeight);
        for (int i = 0; i < title.size() && i < columnNum; i++) {
            if (tableWidths != null) {
                cellWidth = tableWidths[i];
            }
            String msg = title.get(i);
            //单元格对象
            XWPFTableCell xwpfTableCell = xwpfTableRowTitle.getCell(i);
            //添加文本，9号/黑体/黑色/居中
            buildCellSpecial(xwpfTableCell, msg, 9, true, colorBlack, align, cellWidth, null, titleBackground);
        }

//        Boolean flag1 = false;
//        Boolean flag2 = false;
//        Boolean flag3 = false;
        //创建内容
        for (int j = 0; j < content.size(); j++) {
            //有标题时和无标题时取值位置不同
            XWPFTableRow xwpfTableRow;
            if (CollectionUtils.isNotEmpty(title)) {
                // 行对象
                xwpfTableRow = xwpfTable.getRow(j + 1);
                //行高
                xwpfTableRow.setHeight(rowsHeight);
            } else {
                // 行对象
                xwpfTableRow = xwpfTable.getRow(j);
                //行高
                xwpfTableRow.setHeight(rowsHeight);
            }
            List<Object> row = content.get(j);
            for (int i = 0; i < row.size() && i < columnNum; i++) {
                // 合并单元格
//                if (i == 0 && row.get(i) instanceof String && !flag1) {
//                    if ("教育风险".equals((String) row.get(i))) {
//                        flag1 = true;
//                        // todo count值取决于json中"教育风险"对应的数组长度-1
//                        Integer count = 1;
//                        mergeCellsVerticallySpecial(xwpfTable, i, j + 1, j + 1 + count);
//                    }
//                }
//                if (i == 0 && row.get(i) instanceof String && !flag2) {
//                    if ("工作履历风险".equals((String) row.get(i))) {
//                        flag2 = true;
//                        // todo count值取决于json中"工作履历风险"对应的数组长度-1
//                        Integer count = 1;
//                        mergeCellsVerticallySpecial(xwpfTable, i, j + 1, j + 1 + count);
//                    }
//                }
//                if (i == 0 && row.get(i) instanceof String && !flag3) {
//                    if ("工作表现风险".equals((String) row.get(i))) {
//                        flag3 = true;
//                        // todo count值取决于json中"工作表现风险"对应的数组长度-1
//                        Integer count = 1;
//                        mergeCellsVerticallySpecial(xwpfTable, i, j + 1, j + 1 + count);
//                    }
//                }

                //单元格对象
                XWPFTableCell xwpfTableCell = xwpfTableRow.getCell(i);
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                //添加文本，14号/黑体/左对齐
                buildCellSpecial(xwpfTableCell, row.get(i), fontSize, bold, wordColor, align, cellWidth, null, tableBackground);
            }
        }
    }

    public void buildTableSpecial6(XWPFDocument document
            , List<String> title
            , List<List<Object>> content
            , Integer numColumn
            , Long[] tableWidths
            , Integer rowsHeight
            , Long tableWidth
            , Integer displayBorder
            , String titleBackground
            , String tableBackground
            , String wordColor
            , Integer fontSize
            , Boolean bold
            , ParagraphAlignment align) throws InvalidFormatException, IOException, URISyntaxException {
        int rowNum = content.size();
        int columnNum = numColumn == null ? tableWidths.length : numColumn;
        if (CollectionUtils.isNotEmpty(title)) {
            rowNum++;
        }
        XWPFTable xwpfTable = document.createTable(rowNum, columnNum);
        //表格居中显示
        CTTblPr ctTblPr = xwpfTable.getCTTbl().addNewTblPr();
        ctTblPr.addNewJc().setVal(STJc.CENTER);
        //设置表格宽度
        CTTblWidth ctTblWidth = ctTblPr.addNewTblW();
        ctTblWidth.setW(BigInteger.valueOf(tableWidth));
        //设置表格宽度为非自动
        ctTblWidth.setType(STTblWidth.DXA);
        Long cellWidth = null;
        //设置边框
        if (displayBorder == 0) {
            displayBorder(xwpfTable);
        } else if(displayBorder == 2) {
            // 自定义边框：上下左右有灰色边框
            customizeBorderSpecial(xwpfTable, colorGary);
        }

        //创建标题
        if (CollectionUtils.isNotEmpty(title)) {
            //标题行对象
            XWPFTableRow xwpfTableRowTitle = xwpfTable.getRow(0);
            //标题行高
            xwpfTableRowTitle.setHeight(rowsHeight);
            for (int i = 0; i < title.size() && i < columnNum; i++) {
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                String msg = title.get(i);
                //单元格对象
                XWPFTableCell xwpfTableCell = xwpfTableRowTitle.getCell(i);
                buildCellSpecialWithTextWrap(xwpfTableCell, msg, 9, true, colorBlack, align, cellWidth, null, titleBackground);
                // 添加边框
//                addBottomBorder(xwpfTableCell, 1, colorBlue, true);
            }
        }

        //创建内容
        for (int j = 0; j < content.size(); j++) {
            //有标题时和无标题时取值位置不同
            XWPFTableRow xwpfTableRow;
            if (CollectionUtils.isNotEmpty(title)) {
                // 行对象
                xwpfTableRow = xwpfTable.getRow(j + 1);
                //行高
                xwpfTableRow.setHeight(rowsHeight);
            } else {
                // 行对象
                xwpfTableRow = xwpfTable.getRow(j);
                //行高
                xwpfTableRow.setHeight(rowsHeight);
            }
            List<Object> row = content.get(j);
            for (int i = 0; i < row.size() && i < columnNum; i++) {
                //单元格对象
                XWPFTableCell xwpfTableCell = xwpfTableRow.getCell(i);
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                buildCellSpecialWithTextWrap(xwpfTableCell, row.get(i), fontSize, bold, wordColor, align, cellWidth, null, tableBackground);
            }
        }
    }

    /**
     * 创建表格-核实类目明细表
     */
    public void buildTableSpecial7(XWPFDocument document
            , List<String> title
            , List<List<Object>> content
            , Integer numColumn
            , Long[] tableWidths
            , Integer rowsHeight
            , Long tableWidth
            , Integer displayBorder
            , String titleBackground
            , String tableBackground
            , String wordColor
            , Integer fontSize
            , Boolean bold
            , ParagraphAlignment align
            , List<List<Integer>> mergeData) throws InvalidFormatException, IOException, URISyntaxException {
        int rowNum = content.size();
        int columnNum = numColumn == null ? tableWidths.length : numColumn;
        if (CollectionUtils.isNotEmpty(title)) {
            rowNum++;
        }
        XWPFTable xwpfTable = document.createTable(rowNum, columnNum);
        //表格居中显示
        CTTblPr ctTblPr = xwpfTable.getCTTbl().addNewTblPr();
        ctTblPr.addNewJc().setVal(STJc.CENTER);
        //设置表格宽度
        CTTblWidth ctTblWidth = ctTblPr.addNewTblW();
        ctTblWidth.setW(BigInteger.valueOf(tableWidth));
        //设置表格宽度为非自动
        ctTblWidth.setType(STTblWidth.DXA);
        Long cellWidth = null;
        //设置边框
        if (displayBorder == 0) {
            displayBorder(xwpfTable);
        } else if(displayBorder == 2) {
            // 自定义边框：上下左右有灰色边框
            customizeBorderSpecial(xwpfTable, colorGary);
        }

        // 合并单元格
        if (CollectionUtils.isNotEmpty(mergeData)) {
            for (List<Integer> item : mergeData) {
                if (item.get(0) == 1) {
                    if (CollectionUtils.isNotEmpty(title)) {
                        mergeCellsVerticallySpecial(xwpfTable, item.get(1), item.get(2) + 1, item.get(3) + 1);
                    } else {
                        mergeCellsVerticallySpecial(xwpfTable, item.get(1), item.get(2), item.get(3));
                    }
                }
                if (item.get(0) == 2) {
                    if (CollectionUtils.isNotEmpty(title)) {
                        mergeCellsHorizontalSpecial(xwpfTable, item.get(1) + 1, item.get(2), item.get(3));
                    } else {
                        mergeCellsHorizontalSpecial(xwpfTable, item.get(1), item.get(2), item.get(3));
                    }
                }
            }
        }

        //创建标题
        //标题行对象
        XWPFTableRow xwpfTableRowTitle = xwpfTable.getRow(0);
        //标题行高
        xwpfTableRowTitle.setHeight(rowsHeight);
        for (int i = 0; i < title.size() && i < columnNum; i++) {
            if (tableWidths != null) {
                cellWidth = tableWidths[i];
            }
            String msg = title.get(i);
            //单元格对象
            XWPFTableCell xwpfTableCell = xwpfTableRowTitle.getCell(i);
            //添加文本，9号/黑体/黑色/居中
            buildCellSpecial(xwpfTableCell, msg, 9, true, colorBlack, align, cellWidth, null, titleBackground);
        }
        //创建内容
        for (int j = 0; j < content.size(); j++) {
            //有标题时和无标题时取值位置不同
            XWPFTableRow xwpfTableRow;
            if (CollectionUtils.isNotEmpty(title)) {
                // 行对象
                xwpfTableRow = xwpfTable.getRow(j + 1);
                //行高
                xwpfTableRow.setHeight(rowsHeight);
            } else {
                // 行对象
                xwpfTableRow = xwpfTable.getRow(j);
                //行高
                xwpfTableRow.setHeight(rowsHeight);
            }
            List<Object> row = content.get(j);

//            // 纵向合并单元格 身份核实合并（合并第一列第二行和第一列第三行）
//            mergeCellsVerticallySpecial(xwpfTable, 0, 1, 2);
//            // 纵向合并单元格 身份核实-核实结果的合并（合并第三列第二行和第三列第三行）
//            mergeCellsVerticallySpecial(xwpfTable, 2, 1, 2);
            for (int i = 0; i < row.size() && i < columnNum; i++) {
                XWPFTableCell xwpfTableCell = xwpfTableRow.getCell(i);
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                //添加文本
                buildCellSpecial(xwpfTableCell, row.get(i), fontSize, bold, wordColor, align, cellWidth, null, tableBackground);
            }
        }
    }


    public void buildTableSpecial8(XWPFDocument document
            , List<String> title
            , List<List<Object>> content
            , Integer numColumn
            , Long[] tableWidths
            , Long tableWidth
            , Integer displayBorder
            , String tableBackground
            , String wordColor
            , Integer fontSize
            , Boolean bold
            , ParagraphAlignment align) throws InvalidFormatException, IOException, URISyntaxException {
        int rowNum = content.size();
        int columnNum = numColumn == null ? tableWidths.length : numColumn;
        if (CollectionUtils.isNotEmpty(title)) {
            rowNum++;
        }
        XWPFTable xwpfTable = document.createTable(rowNum, columnNum);
        //表格居中显示
        CTTblPr ctTblPr = xwpfTable.getCTTbl().addNewTblPr();
        ctTblPr.addNewJc().setVal(STJc.CENTER);
        //设置表格宽度
        CTTblWidth ctTblWidth = ctTblPr.addNewTblW();
        ctTblWidth.setW(BigInteger.valueOf(tableWidth));
        //设置表格宽度为非自动
        ctTblWidth.setType(STTblWidth.DXA);
        Long cellWidth = null;
        //设置边框
        if (displayBorder == 0) {
            displayBorder(xwpfTable);
        } else if(displayBorder == 2) {
            // 自定义边框：上下左右有灰色边框
            customizeBorderSpecial(xwpfTable, colorGary);
        }

        //创建内容
        for (int j = 0; j < content.size(); j++) {
            //有标题时和无标题时取值位置不同
            XWPFTableRow xwpfTableRow = xwpfTable.getRow(j);
            //行高
            // todo 行高展示出来有问题
            if (j == 0) {
                xwpfTableRow.setHeight(1);
            } else if (j == 1) {
                xwpfTableRow.setHeight(300);
            } else if (j == 2) {
                xwpfTableRow.setHeight(200);
            }
            List<Object> row = content.get(j);
            for (int i = 0; i < row.size() && i < columnNum; i++) {
                //单元格对象
                XWPFTableCell xwpfTableCell = xwpfTableRow.getCell(i);
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                if (j == 1 && i == 1) {
                    buildCellSpecialWithTextWrap(xwpfTableCell, row.get(i), fontSize, bold, wordColor, align, cellWidth, null, colorGary);
                } else {
                    buildCellSpecialWithTextWrap(xwpfTableCell, row.get(i), fontSize, bold, wordColor, align, cellWidth, null, tableBackground);
                }
            }
        }
    }

    public void buildTableSpecial9(XWPFDocument document
            , List<String> title
            , List<List<Object>> content
            , Integer numColumn
            , Long[] tableWidths
            , Integer rowsHeight
            , Long tableWidth
            , Integer displayBorder
            , String titleBackground
            , String tableBackground
            , String wordColor
            , Integer fontSize
            , Boolean bold
            , ParagraphAlignment align
            , List<List<Integer>> mergeData) throws InvalidFormatException, IOException, URISyntaxException {
        int rowNum = content.size();
        int columnNum = numColumn == null ? tableWidths.length : numColumn;
        if (CollectionUtils.isNotEmpty(title)) {
            rowNum++;
        }
        XWPFTable xwpfTable = document.createTable(rowNum, columnNum);
        //表格居中显示
        CTTblPr ctTblPr = xwpfTable.getCTTbl().addNewTblPr();
        ctTblPr.addNewJc().setVal(STJc.CENTER);
        //设置表格宽度
        CTTblWidth ctTblWidth = ctTblPr.addNewTblW();
        ctTblWidth.setW(BigInteger.valueOf(tableWidth));
        //设置表格宽度为非自动
        ctTblWidth.setType(STTblWidth.DXA);
        Long cellWidth = null;
        //设置边框
        if (displayBorder == 0) {
            displayBorder(xwpfTable);
        } else if(displayBorder == 2) {
            // 自定义边框：上下左右有灰色边框
            customizeBorderSpecial(xwpfTable, colorGary);
        }

        // 合并单元格
        if (CollectionUtils.isNotEmpty(mergeData)) {
            for (List<Integer> item : mergeData) {
                if (item.get(0) == 1) {
                    mergeCellsVerticallySpecial(xwpfTable, item.get(1), item.get(2), item.get(3));
                }
                if (item.get(0) == 2) {
                    mergeCellsHorizontalSpecial(xwpfTable, item.get(1), item.get(2), item.get(3));
                }
            }
        }

        //创建标题
        //标题行对象
        XWPFTableRow xwpfTableRowTitle = xwpfTable.getRow(0);
        //标题行高
        xwpfTableRowTitle.setHeight(rowsHeight);
        for (int i = 0; i < title.size() && i < columnNum; i++) {
            if (tableWidths != null) {
                cellWidth = tableWidths[i];
            }
            String msg = title.get(i);
            //单元格对象
            XWPFTableCell xwpfTableCell = xwpfTableRowTitle.getCell(i);
            //添加文本，9号/黑体/黑色/居中
            buildCellSpecial(xwpfTableCell, msg, 9, true, colorBlack, align, cellWidth, null, titleBackground);
        }
        //创建内容
        for (int j = 0; j < content.size(); j++) {
            //有标题时和无标题时取值位置不同
            XWPFTableRow xwpfTableRow;
            if (CollectionUtils.isNotEmpty(title)) {
                // 行对象
                xwpfTableRow = xwpfTable.getRow(j + 1);
                //行高
                xwpfTableRow.setHeight(rowsHeight);
            } else {
                // 行对象
                xwpfTableRow = xwpfTable.getRow(j);
                //行高
                xwpfTableRow.setHeight(rowsHeight);
            }
            List<Object> row = content.get(j);
            for (int i = 0; i < row.size() && i < columnNum; i++) {
                //单元格对象
                XWPFTableCell xwpfTableCell = xwpfTableRow.getCell(i);
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                if (j == 0) {
                    buildCellSpecial(xwpfTableCell, row.get(i), fontSize, true, wordColor, align, cellWidth, null, tableBackground);
                } else {
                    buildCellSpecial(xwpfTableCell, row.get(i), fontSize, bold, wordColor, align, cellWidth, null, tableBackground);
                }
            }
        }
    }

    public void buildTableSpecial10(XWPFDocument document
            , List<String> title
            , List<List<Object>> content
            , Integer numColumn
            , Long[] tableWidths
            , Integer rowsHeight
            , Long tableWidth
            , Integer displayBorder
            , String titleBackground
            , String tableBackground
            , String wordColor
            , Integer fontSize
            , Boolean bold
            , ParagraphAlignment align
            , List<List<Integer>> mergeData) throws InvalidFormatException, IOException, URISyntaxException {
        int rowNum = content.size();
        int columnNum = numColumn == null ? tableWidths.length : numColumn;
        if (CollectionUtils.isNotEmpty(title)) {
            rowNum++;
        }
        XWPFTable xwpfTable = document.createTable(rowNum, columnNum);
        //表格居中显示
        CTTblPr ctTblPr = xwpfTable.getCTTbl().addNewTblPr();
        ctTblPr.addNewJc().setVal(STJc.CENTER);
        //设置表格宽度
        CTTblWidth ctTblWidth = ctTblPr.addNewTblW();
        ctTblWidth.setW(BigInteger.valueOf(tableWidth));
        //设置表格宽度为非自动
        ctTblWidth.setType(STTblWidth.DXA);
        Long cellWidth = null;
        //设置边框
        if (displayBorder == 0) {
            displayBorder(xwpfTable);
        } else if(displayBorder == 2) {
            // 自定义边框：上下左右有灰色边框
            customizeBorderSpecial(xwpfTable, colorGary);
        }

        // 合并单元格
        if (CollectionUtils.isNotEmpty(mergeData)) {
            for (List<Integer> item : mergeData) {
                if (item.get(0) == 1) {
                    mergeCellsVerticallySpecial(xwpfTable, item.get(1), item.get(2), item.get(3));
                }
                if (item.get(0) == 2) {
                    mergeCellsHorizontalSpecial(xwpfTable, item.get(1), item.get(2), item.get(3));
                }
            }
        }

        //创建标题
        //标题行对象
        XWPFTableRow xwpfTableRowTitle = xwpfTable.getRow(0);
        //标题行高
        xwpfTableRowTitle.setHeight(rowsHeight);
        for (int i = 0; i < title.size() && i < columnNum; i++) {
            if (tableWidths != null) {
                cellWidth = tableWidths[i];
            }
            String msg = title.get(i);
            //单元格对象
            XWPFTableCell xwpfTableCell = xwpfTableRowTitle.getCell(i);
            //添加文本，9号/黑体/黑色/居中
            buildCellSpecial(xwpfTableCell, msg, 9, true, colorBlack, align, cellWidth, null, titleBackground);
        }
        //创建内容
        for (int j = 0; j < content.size(); j++) {
            //有标题时和无标题时取值位置不同
            XWPFTableRow xwpfTableRow;
            if (CollectionUtils.isNotEmpty(title)) {
                // 行对象
                xwpfTableRow = xwpfTable.getRow(j + 1);
                //行高
                xwpfTableRow.setHeight(rowsHeight);
            } else {
                // 行对象
                xwpfTableRow = xwpfTable.getRow(j);
                //行高
                xwpfTableRow.setHeight(rowsHeight);
            }
            List<Object> row = content.get(j);
            for (int i = 0; i < row.size() && i < columnNum; i++) {
                XWPFTableCell xwpfTableCell = xwpfTableRow.getCell(i);
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                //添加文本
                buildCellSpecial(xwpfTableCell, row.get(i), fontSize, bold, wordColor, align, cellWidth, null, tableBackground);
            }
        }
    }

    public void buildTableSpecial11(XWPFDocument document
            , List<String> title
            , List<List<Object>> content
            , Integer numColumn
            , Long[] tableWidths
            , Integer rowsHeight
            , Long tableWidth
            , Integer displayBorder
            , String titleBackground
            , String tableBackground
            , String wordColor
            , Integer fontSize
            , Boolean bold
            , ParagraphAlignment align
            , List<List<Integer>> mergeData) throws InvalidFormatException, IOException, URISyntaxException {
        int rowNum = content.size();
        int columnNum = numColumn == null ? tableWidths.length : numColumn;
        if (CollectionUtils.isNotEmpty(title)) {
            rowNum++;
        }
        XWPFTable xwpfTable = document.createTable(rowNum, columnNum);
        //表格居中显示
        CTTblPr ctTblPr = xwpfTable.getCTTbl().addNewTblPr();
        ctTblPr.addNewJc().setVal(STJc.CENTER);
        //设置表格宽度
        CTTblWidth ctTblWidth = ctTblPr.addNewTblW();
        ctTblWidth.setW(BigInteger.valueOf(tableWidth));
        //设置表格宽度为非自动
        ctTblWidth.setType(STTblWidth.DXA);
        Long cellWidth = null;
        //设置边框
        if (displayBorder == 0) {
            displayBorder(xwpfTable);
        } else if(displayBorder == 2) {
            // 自定义边框：上下左右有灰色边框
            customizeBorderSpecial(xwpfTable, colorGary);
        }

        // 合并单元格
        if (CollectionUtils.isNotEmpty(mergeData)) {
            for (List<Integer> item : mergeData) {
                if (item.get(0) == 1) {
                    mergeCellsVerticallySpecial(xwpfTable, item.get(1), item.get(2), item.get(3));
                }
                if (item.get(0) == 2) {
                    mergeCellsHorizontalSpecial(xwpfTable, item.get(1), item.get(2), item.get(3));
                }
            }
        }

        //创建标题
        //标题行对象
        XWPFTableRow xwpfTableRowTitle = xwpfTable.getRow(0);
        //标题行高
        xwpfTableRowTitle.setHeight(rowsHeight);
        for (int i = 0; i < title.size() && i < columnNum; i++) {
            if (tableWidths != null) {
                cellWidth = tableWidths[i];
            }
            String msg = title.get(i);
            //单元格对象
            XWPFTableCell xwpfTableCell = xwpfTableRowTitle.getCell(i);
            //添加文本，9号/黑体/黑色/居中
            buildCellSpecial(xwpfTableCell, msg, 9, true, colorBlack, align, cellWidth, null, titleBackground);
        }
        //创建内容
        for (int j = 0; j < content.size(); j++) {
            //有标题时和无标题时取值位置不同
            XWPFTableRow xwpfTableRow;
            if (CollectionUtils.isNotEmpty(title)) {
                // 行对象
                xwpfTableRow = xwpfTable.getRow(j + 1);
                //行高
                xwpfTableRow.setHeight(rowsHeight);
            } else {
                // 行对象
                xwpfTableRow = xwpfTable.getRow(j);
                //行高
                xwpfTableRow.setHeight(rowsHeight);
            }
            List<Object> row = content.get(j);
            for (int i = 0; i < row.size() && i < columnNum; i++) {
                XWPFTableCell xwpfTableCell = xwpfTableRow.getCell(i);
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                if (j == content.size() - 1) {
                    buildCellSpecial(xwpfTableCell, row.get(i), fontSize, bold, wordColor, ParagraphAlignment.LEFT, cellWidth, null, tableBackground);
                } else {
                    buildCellSpecial(xwpfTableCell, row.get(i), fontSize, bold, wordColor, align, cellWidth, null, tableBackground);
                }
            }
        }
    }
    /**
     * 自定义边框
     * 隐藏内部边框
     *
     * @param table
     */
    public void customizeBorderSpecial(XWPFTable table, String color) {
        CTTblBorders ctTblBorders = table.getCTTbl().getTblPr().addNewTblBorders();

        CTBorder leftBorder = ctTblBorders.addNewLeft();
        leftBorder.setVal(STBorder.THICK);
        leftBorder.setSz(BigInteger.valueOf(10L));
        leftBorder.setColor(color);
        ctTblBorders.setLeft(leftBorder);

        CTBorder rBorder = ctTblBorders.addNewRight();
        rBorder.setVal(STBorder.THICK);
        rBorder.setSz(BigInteger.valueOf(10L));
        rBorder.setColor(color);
        ctTblBorders.setRight(rBorder);

        CTBorder tBorder = ctTblBorders.addNewTop();
        tBorder.setVal(STBorder.THICK);
        tBorder.setSz(BigInteger.valueOf(10L));
        tBorder.setColor(color);
        ctTblBorders.setTop(tBorder);

        CTBorder bBorder = ctTblBorders.addNewBottom();
        bBorder.setVal(STBorder.THICK);
        bBorder.setSz(BigInteger.valueOf(10L));
        bBorder.setColor(color);
        ctTblBorders.setBottom(bBorder);

        CTBorder vBorder = ctTblBorders.addNewInsideV();
        vBorder.setVal(STBorder.NIL);
        ctTblBorders.setInsideV(vBorder);

        CTBorder hBorder = ctTblBorders.addNewInsideH();
        hBorder.setVal(STBorder.NIL);
        ctTblBorders.setInsideH(hBorder);
    }

    /**
     * 创建单元格 (对象，可以是String也可以是Image,指定字体，水平居...)
     *
     * @param cell XWPFTableCell对象
     * @param value 填入单元格的内容，可以是String也可以是Image
     * @param fontSize 字体大小
     * @param bold 字体是否加粗
     * @param color 字体颜色
     * @param align 字体对齐方式
     * @param width 单元格宽度
     * @param mediate 是否竖直居中
     * @param backgroundColor 单元格背景色
     * @throws IOException
     * @throws InvalidFormatException
     * @throws URISyntaxException
     */
    public void buildCellSpecialWithTextWrap(XWPFTableCell cell, Object value, Integer fontSize, Boolean bold, String color
            , ParagraphAlignment align, Long width, Boolean mediate, String backgroundColor) throws IOException, InvalidFormatException, URISyntaxException {
        // 单元格背景色
        if (StringUtils.isNotBlank(backgroundColor)) {
            cell.setColor(backgroundColor);
        }
        CTTc cttc = cell.getCTTc();
        CTTcPr ctPr = cttc.addNewTcPr();
        // 水平居中
        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        if (mediate != null && mediate == true) {
            // 竖直居中
            ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
            cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
        }
        // 单元格宽度
        if (width != null) {
            CTTblWidth ctTblWidthCell = ctPr.addNewTcW();
            ctTblWidthCell.setType(STTblWidth.DXA);
            ctTblWidthCell.setW(BigInteger.valueOf(width));
        }
        if (value instanceof String) {
            String str = value.toString();
            String[] strs = str.split("\n");
            for (String item : strs) {
                XWPFParagraph paragraph = cell.getParagraphs().get(0);
                buildParagraphSpecialWithTextWrap(paragraph, align, (String) item, fontSize, bold, color);
                cell.setParagraph(paragraph);
                cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
            }
        } else if (value instanceof StringBuffer) {
            String imageName = value.toString().substring(value.toString().lastIndexOf("/") + 1);
            XWPFRun pictureRun = cell.getParagraphs().get(0).createRun();
            FileInputStream is = null;
            try {
                // todo 111
//                is = new FileInputStream(new File(this.getClass().getResource(((StringBuffer) value).toString()).toURI()));
                is = new FileInputStream(new File(((StringBuffer) value).toString()));
                if (imageName.contains("attachment")) {
                    pictureRun.addPicture(is, Document.PICTURE_TYPE_JPEG, null, Units.toEMU(420), Units.toEMU(220));
                } else {
                    pictureRun.addPicture(is, Document.PICTURE_TYPE_JPEG, null, Units.toEMU(14), Units.toEMU(14));
                }
            } catch (IOException e) {
                throw e;
            } finally {
                if (is != null) {
                    is.close();
                }
            }
        }
    }


    /**
     * 创建单元格 (对象，可以是String也可以是Image,指定字体，水平居...)
     *
     * @param cell XWPFTableCell对象
     * @param value 填入单元格的内容，可以是String也可以是Image
     * @param fontSize 字体大小
     * @param bold 字体是否加粗
     * @param color 字体颜色
     * @param align 字体对齐方式
     * @param width 单元格宽度
     * @param mediate 是否竖直居中
     * @param backgroundColor 单元格背景色
     * @throws IOException
     * @throws InvalidFormatException
     * @throws URISyntaxException
     */
    public void buildCellSpecial(XWPFTableCell cell, Object value, Integer fontSize, Boolean bold, String color
            , ParagraphAlignment align, Long width, Boolean mediate, String backgroundColor) throws IOException, InvalidFormatException, URISyntaxException {
        // 单元格背景色
        if (StringUtils.isNotBlank(backgroundColor)) {
            cell.setColor(backgroundColor);
        }
        CTTc cttc = cell.getCTTc();
        CTTcPr ctPr = cttc.addNewTcPr();
        // 水平居中
        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        if (mediate != null && mediate == true) {
            // 竖直居中
            ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
            cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
        }
        // 单元格宽度
        if (width != null) {
            CTTblWidth ctTblWidthCell = ctPr.addNewTcW();
            ctTblWidthCell.setType(STTblWidth.DXA);
            ctTblWidthCell.setW(BigInteger.valueOf(width));
        }
        if (value instanceof String) {
            XWPFParagraph paragraph = cell.getParagraphs().get(0);
            buildParagraphSpecial(paragraph, align, (String) value, fontSize, bold, color);
            cell.setParagraph(paragraph);
        } else if (value instanceof StringBuffer) {
            XWPFParagraph paragraph = cell.getParagraphs().get(0);
            paragraph.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun pictureRun = paragraph.createRun();
            FileInputStream is = null;
            try {
                // todo 111
//                is = new FileInputStream(new File(this.getClass().getResource(((StringBuffer) value).toString()).toURI()));
                is = new FileInputStream(new File(((StringBuffer) value).toString()));
                pictureRun.addPicture(is, Document.PICTURE_TYPE_JPEG, null, Units.toEMU(14), Units.toEMU(14));
            } catch (IOException e) {
                throw e;
            } finally {
                if (is != null) {
                    is.close();
                }
            }
        } else if (value instanceof InputStream) {
            XWPFParagraph paragraph = cell.getParagraphs().get(0);
            paragraph.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun pictureRun = paragraph.createRun();
            pictureRun.addPicture((InputStream) value, Document.PICTURE_TYPE_PNG, null, Units.toEMU(420), Units.toEMU(220));
        }
    }

    public void buildCellSpecial2(XWPFTableCell cell, Object value, Integer fontSize, Boolean bold, String color
            , ParagraphAlignment align, Long width, Boolean mediate, String backgroundColor) throws IOException, InvalidFormatException, URISyntaxException {
        // 单元格背景色
        if (StringUtils.isNotBlank(backgroundColor)) {
            cell.setColor(backgroundColor);
        }
        CTTc cttc = cell.getCTTc();
        CTTcPr ctPr = cttc.addNewTcPr();
        // 水平居中
        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        if (mediate != null && mediate == true) {
            // 竖直居中
            ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
            cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
        }
        // 单元格宽度
        if (width != null) {
            CTTblWidth ctTblWidthCell = ctPr.addNewTcW();
            ctTblWidthCell.setType(STTblWidth.DXA);
            ctTblWidthCell.setW(BigInteger.valueOf(width));
        }
        if (value instanceof String) {
            XWPFParagraph paragraph = cell.getParagraphs().get(0);
            buildParagraphSpecial(paragraph, align, (String) value, fontSize, bold, color);
            cell.setParagraph(paragraph);
        } else if (value instanceof StringBuffer) {
            XWPFParagraph paragraph = cell.getParagraphs().get(0);
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            XWPFRun pictureRun = paragraph.createRun();
            FileInputStream is = null;
            try {
                // todo 111
//                is = new FileInputStream(new File(this.getClass().getResource(((StringBuffer) value).toString()).toURI()));
                is = new FileInputStream(new File(((StringBuffer) value).toString()));
                pictureRun.addPicture(is, Document.PICTURE_TYPE_JPEG, null, Units.toEMU(14), Units.toEMU(14));
            } catch (IOException e) {
                throw e;
            } finally {
                if (is != null) {
                    is.close();
                }
            }
        } else if (value instanceof InputStream) {
            XWPFParagraph paragraph = cell.getParagraphs().get(0);
            paragraph.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun pictureRun = paragraph.createRun();
            pictureRun.addPicture((InputStream) value, Document.PICTURE_TYPE_PNG, null, Units.toEMU(420), Units.toEMU(220));
        }
    }

    public void buildCellSpecialAttachment(XWPFTableCell cell, Object value, Integer fontSize, Boolean bold, String color
            , ParagraphAlignment align, Long width, Boolean mediate, String backgroundColor, String type) throws IOException, InvalidFormatException, URISyntaxException {
        // 单元格背景色
        if (StringUtils.isNotBlank(backgroundColor)) {
            cell.setColor(backgroundColor);
        }
        CTTc cttc = cell.getCTTc();
        CTTcPr ctPr = cttc.addNewTcPr();
        // 水平居中
        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        if (mediate != null && mediate == true) {
            // 竖直居中
            ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
            cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
        }
        // 单元格宽度
        if (width != null) {
            CTTblWidth ctTblWidthCell = ctPr.addNewTcW();
            ctTblWidthCell.setType(STTblWidth.DXA);
            ctTblWidthCell.setW(BigInteger.valueOf(width));
        }
        if (value instanceof String) {
            XWPFParagraph paragraph = cell.getParagraphs().get(0);
            buildParagraphSpecial(paragraph, align, (String) value, fontSize, bold, color);
            cell.setParagraph(paragraph);
        } else if (value instanceof StringBuffer) {
            XWPFParagraph paragraph = cell.getParagraphs().get(0);
            paragraph.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun pictureRun = paragraph.createRun();
            FileInputStream is = null;
            try {
                // todo 111
//                is = new FileInputStream(new File(this.getClass().getResource(((StringBuffer) value).toString()).toURI()));
                is = new FileInputStream(new File(((StringBuffer) value).toString()));
                pictureRun.addPicture(is, Document.PICTURE_TYPE_JPEG, null, Units.toEMU(14), Units.toEMU(14));
            } catch (IOException e) {
                throw e;
            } finally {
                if (is != null) {
                    is.close();
                }
            }
        } else if (value instanceof InputStream) {
            XWPFParagraph paragraph = cell.getParagraphs().get(0);
            paragraph.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun pictureRun = paragraph.createRun();
            if (type == "diploma") {
                pictureRun.addPicture((InputStream) value, Document.PICTURE_TYPE_PNG, null, Units.toEMU(420), Units.toEMU(220));
            } else if (type == "degree") {
                pictureRun.addPicture((InputStream) value, Document.PICTURE_TYPE_PNG, null, Units.toEMU(420), Units.toEMU(220));
            } else if (type == "authLetter") {
                pictureRun.addPicture((InputStream) value, Document.PICTURE_TYPE_PNG, null, Units.toEMU(300), Units.toEMU(500));
            }
        }
    }

    /**
     * 构建文本、文本位置、字体大小、是否加粗、颜色
     *
     * @param paragraph
     * @param align
     * @param content
     * @param fontSize
     * @param bold
     * @param color
     */
    public void buildParagraphSpecial(XWPFParagraph paragraph, ParagraphAlignment align, String content, int fontSize
            , Boolean bold, String color) {
        paragraph.setAlignment(align);
        buildParagraphSpecial(paragraph, content, fontSize, bold, color);
    }

    /**
     * 构建文本、文本位置、字体大小、是否加粗、颜色
     *
     * @param paragraph
     * @param align
     * @param content
     * @param fontSize
     * @param bold
     * @param color
     */
    public void buildParagraphSpecialWithTextWrap(XWPFParagraph paragraph, ParagraphAlignment align, String content, int fontSize
            , Boolean bold, String color) {
        paragraph.setAlignment(align);
        buildParagraphSpecialWithTextWrap(paragraph, content, fontSize, bold, color);
    }

    /**
     * 构建文本、字体大小、是否加粗、颜色
     *
     * @param paragraph XWPFParagraph对象
     * @param content 填充内容
     * @param fontSize 字体大小
     * @param bold 是否加粗
     * @param color 字体颜色
     */
    public void buildParagraphSpecialWithTextWrap(XWPFParagraph paragraph, String content, int fontSize
            , Boolean bold, String color) {
        XWPFRun xwpfRun = paragraph.createRun();
        xwpfRun.setText(content);
        xwpfRun.setFontSize(fontSize);
//        xwpfRun.setFontFamily("黑体");
        xwpfRun.setFontFamily("微软雅黑");
        if (bold != null) {
            xwpfRun.setBold(bold);
        }
        if (color != null) {
            xwpfRun.setColor(color);
        }
        xwpfRun.addBreak(BreakType.TEXT_WRAPPING);
    }

    /**
     * 构建文本、字体大小、是否加粗、颜色
     *
     * @param paragraph XWPFParagraph对象
     * @param content 填充内容
     * @param fontSize 字体大小
     * @param bold 是否加粗
     * @param color 字体颜色
     */
    public void buildParagraphSpecial(XWPFParagraph paragraph, String content, int fontSize
            , Boolean bold, String color) {
        XWPFRun xwpfRun = paragraph.createRun();
        xwpfRun.setText(content);
        xwpfRun.setFontSize(fontSize);
//        xwpfRun.setFontFamily("黑体");
        xwpfRun.setFontFamily("微软雅黑");
        if (bold != null) {
            xwpfRun.setBold(bold);
        }
        if (color != null) {
            xwpfRun.setColor(color);
        }
    }

    /**
     * 标题
     */
    public void buildTitleSpecial(XWPFDocument document, String titleName) throws IOException, InvalidFormatException, URISyntaxException {
        XWPFTable table = document.createTable(1, 2);
        //表格居中显示
        CTTblPr ctTblPr = table.getCTTbl().addNewTblPr();
        ctTblPr.addNewJc().setVal(STJc.CENTER);
        //设置表格宽度
        CTTblWidth ctTblWidth2 = ctTblPr.addNewTblW();
        ctTblWidth2.setW(BigInteger.valueOf(8310));
        //设置表格宽度为非自动
        ctTblWidth2.setType(STTblWidth.DXA);
        //设置边框
        displayBorder(table);

        // 第一行
        XWPFTableRow row = table.getRow(0);

        // 第一行第一列
        XWPFTableCell xwpfTableCell1 = row.getCell(0);
        xwpfTableCell1.setColor(colorBlue);
        CTTc cttc1 = xwpfTableCell1.getCTTc();
        CTTcPr ctPr1 = cttc1.addNewTcPr();
        // 设置宽度
        CTTblWidth ctTblWidthCell1 = ctPr1.addNewTcW();
        ctTblWidthCell1.setType(STTblWidth.DXA);
        ctTblWidthCell1.setW(BigInteger.valueOf(40));

        // 第一行第二列
        XWPFTableCell xwpfTableCell2 = row.getCell(1);
        buildCell(xwpfTableCell2, titleName, 14, true, colorBlack, ParagraphAlignment.LEFT, 8270L, true);
    }

    /**
     * 页眉页脚
     */
    public void createHeaderAndFooterSpecial( XWPFDocument document) throws Exception {
        // 页眉
        CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();

        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(document, sectPr);

        XWPFHeader header = headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);
        XWPFParagraph paragraph = header.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.RIGHT);
        XWPFRun run = paragraph.createRun();

        FileInputStream is = null;
        try {
            // todo 111
//            is = new FileInputStream(new File(this.getClass().getResource("/static/image/logo-01.jpeg").toURI()));
            is = new FileInputStream(new File("/target/classes/static/image/logo-01.jpeg"));
            XWPFPicture picture = run.addPicture( is, XWPFDocument.PICTURE_TYPE_JPEG, null, Units.toEMU( 160 ), Units.toEMU( 90 ) );
            String blipID = "";
            for( XWPFPictureData pictureData : header.getAllPackagePictures() ) { // 这段必须有，不然打开的logo图片不显示
                blipID = header.getRelationId( pictureData );
                picture.getCTPicture().getBlipFill().getBlip().setEmbed( blipID );
            }
        } catch (IOException e) {
            throw e;
        } finally {
            if (is != null) {
                is.close();
            }
        }

        // 页脚
        XWPFFooter footer = headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);
        XWPFParagraph paragraph1 = footer.createParagraph();
        paragraph1.setAlignment(ParagraphAlignment.RIGHT);

        paragraph1.getCTP().addNewFldSimple().setInstr("PAGE \\* MERGEFORMAT");
        XWPFRun runFooter1 = paragraph1.createRun();
        runFooter1.setText(" / ");
        paragraph1.getCTP().addNewFldSimple().setInstr("NUMPAGES \\* MERGEFORMAT");
    }
}
