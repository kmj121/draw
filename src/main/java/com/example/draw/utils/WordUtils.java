package com.example.draw.utils;

import com.alibaba.fastjson.JSONObject;
import com.itextpdf.text.DocumentException;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.net.URISyntaxException;
import java.util.List;

public abstract class WordUtils {

    //浅蓝色
    protected static String colorLightBlue = "dff0fa";
    //白色
    protected static String colorWrite = "ffffff";
    //蓝色
    protected static String colorBlue = "007ac9";
    //灰色
    protected static String colorGary = "efefef";
    //灰色
    protected static String colorGary2 = "adafaf";
    //黑色
    protected static String colorBlack = "000000";
    //绿色
    protected static String colorGreen = "70af47";
    //红色
    protected static String colorRed = "ff0000";
    // 粉色
    protected static String colorPink = "ffe5e5";
    // 浅蓝色
    protected static String colorBlue2 = "cce0ff";
    // 浅橘色
    protected static String colorOrange = "ffd699";


    private static Logger logger = LoggerFactory.getLogger(WordUtils.class);


    public void buildWord(String filePath, JSONObject params, String getSignUrl) throws Exception {
        FileOutputStream out = null;
        XWPFDocument document = new XWPFDocument();
        try {
            out = new FileOutputStream(filePath);
            generateWord(document, params, getSignUrl);
            document.write(out);
        } catch (Exception e) {
            throw e;
        } finally {
            document.close();
            if (out != null) {
                out.close();
            }
        }
    }

    /**
     * 生成word文档
     *
     * @param document
     * @param params
     * @throws IOException
     * @throws InvalidFormatException
     * @throws URISyntaxException
     * @throws DocumentException
     */
    public abstract void generateWord(XWPFDocument document, JSONObject params, String getSignUrl) throws Exception;

    /**
     * 构建主标题部分
     *
     * @param document
     * @param name
     * @throws IOException
     */
//    public void buildTitleTable(XWPFDocument document, String name) throws IOException, InvalidFormatException, URISyntaxException {
//        //创建表格
//        XWPFTable table = document.createTable(1, 1);
//        List<XWPFTableRow> rows = table.getRows();
//        XWPFTableRow row = table.getRow(0);
//        XWPFTableCell imgCell = row.getCell(0);
//
//        //设置表格宽度
//        CTTblPr tablePr = table.getCTTbl().addNewTblPr();
//        //表格宽度
//        CTTblWidth width = tablePr.addNewTblW();
//        width.setW(BigInteger.valueOf(8310));
//        //设置表格宽度为非自动
//        width.setType(STTblWidth.DXA);
//        //设置边框
//        displayBorder(table);
//        //图片
//        XWPFParagraph imgParagraph = imgCell.getParagraphs().get(0);
//        XWPFRun pictureRun = imgParagraph.createRun();
//        XWPFRun titleRun = imgParagraph.createRun();
//        FileInputStream is = null;
//        try {
//            is = new FileInputStream(new File(this.getClass().getResource("/target/classes/static/image/image1.png").toURI()));
//            pictureRun.addPicture(is, Document.PICTURE_TYPE_JPEG, "c1.png", Units.toEMU(120), Units.toEMU(30));
//        } catch (IOException e) {
//            throw e;
//        } finally {
//            if (is != null) {
//                is.close();
//            }
//        }
//        titleRun.setText("            ");
//        titleRun.setText("  ");
//        titleRun.setText(name);
//        titleRun.setBold(true);
//        titleRun.setFontFamily("黑体");
//        //字体颜色
//        titleRun.setColor(colorWrite);
//        titleRun.setFontSize(18);
//
//        //背景颜色
//        imgCell.setColor(colorBlue);
//
//    }

    /**
     * 构建主表格
     *
     * @param document
     * @param content
     * @param numColumn
     */
    public void buildMainTable(XWPFDocument document, List<List<String>> content, Integer numColumn) throws InvalidFormatException, IOException, URISyntaxException {
        XWPFTable table = document.createTable(2, 2);

        //设置表格宽度
        CTTblPr tablePr = table.getCTTbl().addNewTblPr();
        //表格宽度
        CTTblWidth tableWidth = tablePr.addNewTblW();
        tableWidth.setW(BigInteger.valueOf(8310));
        //设置表格宽度为非自动
        tableWidth.setType(STTblWidth.DXA);
        //获取行列
        for (int j = 0; j < content.size(); j++) {
            XWPFTableRow xwpfTableRow = table.getRow(j);
            List<String> row = content.get(j);
            for (int i = 0; i < row.size() && i < numColumn; i++) {
                XWPFTableCell cell = xwpfTableRow.getCell(i);
                cell.setColor(colorGary);
                buildCell(cell, row.get(i), 9, false, colorBlack, ParagraphAlignment.LEFT, 4155L, null);
            }
        }
        //设置边框
        displayBorder(table);
    }

    /**
     * 创建副标题数据
     *
     * @param document
     * @param titleName
     */
//    public void buildTitle(XWPFDocument document, String titleName, String content) throws IOException, InvalidFormatException, URISyntaxException {
//
//        document.createParagraph();
//
//        XWPFParagraph titleParagraph = document.createParagraph();
//        XWPFRun titleRun = titleParagraph.createRun();
//        titleRun.setText("");
//        titleRun.setText(titleName);
//        titleRun.setBold(true);
//        titleRun.setFontFamily("黑体");
//        titleRun.setColor(colorBlue);
//
//        XWPFRun pictureRun = document.createParagraph().createRun();
//        FileInputStream is = null;
//        try {
//            is = new FileInputStream(new File(this.getClass().getResource("/target/classes/static/image/line.png").toURI()));
//            pictureRun.addPicture(is, Document.PICTURE_TYPE_PNG, null, Units.toEMU(420), Units.toEMU(6));
//        } catch (IOException e) {
//            throw e;
//        } finally {
//            if (is != null) {
//                is.close();
//            }
//        }
//        XWPFParagraph paragraph = document.createParagraph();
//        buildParagraph(paragraph, "    " + content, 9, null, null);
//    }

    /**
     * 创建表格
     *
     * @param document
     * @param title
     * @param content
     * @param numColumn
     * @param width
     * @throws DocumentException
     */
    public void buildTable(XWPFDocument document, List<String> title, List<List<String>> content, Integer numColumn, Long width) throws InvalidFormatException, IOException, URISyntaxException {
        buildTable(document, title, null, content, numColumn, width);
    }

    public void buildTable(XWPFDocument document, List<String> title, List<String> foot, List<List<String>> content, Integer numColumn
            , Long width) throws InvalidFormatException, IOException, URISyntaxException {
        buildTable(document, title, foot, content, numColumn, null, width);
    }

    public void buildTable(XWPFDocument document, List<String> title, List<String> foot, List<List<String>> content, Long[] tableWidths
            , Long width) throws InvalidFormatException, IOException, URISyntaxException {
        buildTable(document, title, foot, content, null, tableWidths, width);
    }

    public void buildTable(XWPFDocument document, List<String> title, List<String> foot, List<List<String>> content, Integer numColumn
            , Long[] tableWidths, Long width) throws InvalidFormatException, IOException, URISyntaxException {
        int rowNum = content.size();
        int columnNum = numColumn == null ? tableWidths.length : numColumn;
        if (CollectionUtils.isNotEmpty(title)) {
            rowNum++;
        }
        if (CollectionUtils.isNotEmpty(foot)) {
            rowNum++;
        }
        XWPFTable xwpfTable = document.createTable(rowNum, columnNum);
        //表格居中显示
        CTTblPr ctTblPr = xwpfTable.getCTTbl().addNewTblPr();
        ctTblPr.addNewJc().setVal(STJc.CENTER);
        //设置表格宽度
        CTTblWidth ctTblWidth = ctTblPr.addNewTblW();
        ctTblWidth.setW(BigInteger.valueOf(width));
        //设置表格宽度为非自动
        ctTblWidth.setType(STTblWidth.DXA);
        Long cellWidth = null;


        //行对象
        XWPFTableRow xwpfTableRowTitle = xwpfTable.getRow(0);
        //行高
        xwpfTableRowTitle.setHeight(350);
        //创建标题
        for (int i = 0; i < title.size() && i < columnNum; i++) {
            if (tableWidths != null) {
                cellWidth = tableWidths[i];
            }
            String msg = title.get(i);
            //单元格对象
            XWPFTableCell xwpfTableCell = xwpfTableRowTitle.getCell(i);
            //添加文本，9号/黑体/白色/居中
            buildCell(xwpfTableCell, msg, 9, true, colorWrite, ParagraphAlignment.CENTER, cellWidth, null);
            //添加背景色，蓝色
            if (StringUtils.isNotEmpty(msg)) {
                xwpfTableCell.setColor(colorBlue);
            }
            //添加边框
            addBottomBorder(xwpfTableCell, 1, colorBlue, true);
        }

        String baseColor = colorLightBlue;
        //创建内容
        for (int j = 0; j < content.size(); j++) {
            //行对象
            XWPFTableRow xwpfTableRow = xwpfTable.getRow(j + 1);
            //行高
            xwpfTableRow.setHeight(350);

            List<String> row = content.get(j);
            //每行轮换颜色
            if (j % 2 == 0) {
                baseColor = colorLightBlue;
            } else {
                baseColor = colorWrite;
            }
            //如果没有尾行则最后一行添加边框
            boolean lastLineFlag = false;
            if (j == content.size() - 1 && (foot == null || foot.size() <= 0)) {
                lastLineFlag = true;
            }
            for (int i = 0; i < row.size() && i < columnNum; i++) {
                //单元格对象
                XWPFTableCell xwpfTableCell = xwpfTableRow.getCell(i);
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                //第一列为蓝字，后面为黑字
                //第一列居左, 后面为居中
                String wordColor = colorBlack;
                ParagraphAlignment align = ParagraphAlignment.CENTER;
                if (i == 0) {
                    wordColor = colorBlue;
                    align = ParagraphAlignment.LEFT;
                }
                //添加文本，9号/黑体/左对齐
                buildCell(xwpfTableCell, row.get(i), 8, false, wordColor, align, cellWidth, null);
                //添加边框
                addBottomBorder(xwpfTableCell, 1, colorBlue, lastLineFlag);
                //添加背景色，蓝色
                xwpfTableCell.setColor(baseColor);
            }
        }

        //创建尾行
        if (foot != null && foot.size() > 0) {
            //行对象
            XWPFTableRow xwpfTableRowFoot = xwpfTable.getRow(content.size() + 1);
            //行高
            xwpfTableRowFoot.setHeight(350);

            for (int i = 0; i < foot.size() && i < columnNum; i++) {
                //单元格对象
                XWPFTableCell xwpfTableCell = xwpfTableRowFoot.getCell(i);
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                //第一列为蓝色加粗字，后面为黑色加粗字
                //第一列居左，后面居中
                String wordColor = colorBlack;
                ParagraphAlignment align = ParagraphAlignment.CENTER;
                if (i == 0) {
                    wordColor = colorBlue;
                    align = ParagraphAlignment.LEFT;
                }
                //添加文本，9号/黑体/左对齐
                buildCell(xwpfTableCell, foot.get(i), 10, true, wordColor, align, cellWidth, null);

                //添加边框
                addBottomBorder(xwpfTableCell, 1, colorBlue, true);

                //添加背景色，与上一行不同的样式
                if (baseColor.equals(colorLightBlue)) {
                    xwpfTableCell.setColor(colorWrite);
                } else {
                    xwpfTableCell.setColor(colorLightBlue);
                }

            }
        }

    }

    public void buildTableForScaleInhibitor(XWPFDocument document, List<String> title, List<List<String>> content, Integer numColumn
            , Long width, boolean colorFlag) throws InvalidFormatException, IOException, URISyntaxException {
        int rowNum = content.size();
        int columnNum = numColumn;
        if (CollectionUtils.isNotEmpty(title)) {
            rowNum++;
        }
        XWPFTable xwpfTable = document.createTable(rowNum, columnNum);
        //表格居中显示
        CTTblPr ctTblPr = xwpfTable.getCTTbl().addNewTblPr();
        ctTblPr.addNewJc().setVal(STJc.CENTER);
        //设置表格宽度
        CTTblWidth ctTblWidth = ctTblPr.addNewTblW();
        ctTblWidth.setW(BigInteger.valueOf(width));
        //设置表格宽度为非自动
        ctTblWidth.setType(STTblWidth.DXA);

        //行对象
        XWPFTableRow xwpfTableRowTitle = xwpfTable.getRow(0);
        //行高
        xwpfTableRowTitle.setHeight(350);
        //创建标题
        for (int i = 0; i < title.size() && i < numColumn; i++) {
            String msg = title.get(i);
            //单元格对象
            XWPFTableCell xwpfTableCell = xwpfTableRowTitle.getCell(i);
            //添加文本，9号/黑体/白色/居中
            buildCell(xwpfTableCell, msg, 9, true, colorWrite, ParagraphAlignment.CENTER, null, null);
            //添加背景色，蓝色
            if (StringUtils.isNotEmpty(msg)) {
                xwpfTableCell.setColor(colorBlue);
            }
            //添加边框
            addBottomBorder(xwpfTableCell, 1, colorBlue, true);
        }

        String baseColor = colorLightBlue;
        //创建内容
        for (int j = 0; j < content.size(); j++) {
            List<String> row = content.get(j);
            //行对象
            XWPFTableRow xwpfTableRow = xwpfTable.getRow(j + 1);
            //行高
            xwpfTableRow.setHeight(350);
            //每行轮换颜色
            if (j % 2 == 0) {
                baseColor = colorLightBlue;
            } else {
                baseColor = colorWrite;
            }
            //如果没有尾行则最后一行添加边框
            boolean lastLineFlag = false;
            if (j == content.size() - 1) {
                lastLineFlag = true;
            }
            for (int i = 0; i < row.size() && i < numColumn; i++) {
                //单元格对象
                XWPFTableCell xwpfTableCell = xwpfTableRow.getCell(i);
                //第一列为蓝字，后面为黑字
                //第一列居左, 后面为居中
                String wordColor = colorBlack;
                ParagraphAlignment align = ParagraphAlignment.CENTER;
                if (i == 0) {
                    wordColor = colorBlue;
                    align = ParagraphAlignment.LEFT;
                }
                //添加文本，9号/黑体/左对齐
                buildCell(xwpfTableCell, row.get(i), 8, false, wordColor, align, null, null);

                //添加边框
                addBottomBorder(xwpfTableCell, 1, colorBlue, lastLineFlag);

                if (i >= 3 && StringUtils.isNotBlank(row.get(i)) && colorFlag) {
                    if (new BigDecimal(row.get(i)).compareTo(new BigDecimal(100)) > 0) {
                        //添加背景色，红色
                        xwpfTableCell.setColor(colorRed);
                    } else {
                        //添加背景色，绿色
                        xwpfTableCell.setColor(colorGreen);
                    }
                } else {
                    //添加背景色，蓝色
                    xwpfTableCell.setColor(baseColor);
                }
            }
        }
    }


    /**
     * 构建无标题表格
     *
     * @param document
     * @param content
     * @param tableWidths
     * @param width
     * @throws DocumentException
     * @throws InvalidFormatException
     * @throws IOException
     * @throws URISyntaxException
     */
    public void buildTable(XWPFDocument document, List<List<String>> content, Long[] tableWidths, long width)
            throws InvalidFormatException, IOException, URISyntaxException {
        int rowNum = content.size();
        int columnNum = tableWidths.length;
        XWPFTable xwpfTable = document.createTable(rowNum, columnNum);
        //表格居中显示
        CTTblPr ctTblPr = xwpfTable.getCTTbl().addNewTblPr();
        ctTblPr.addNewJc().setVal(STJc.CENTER);
        //设置表格宽度
        CTTblWidth ctTblWidth = ctTblPr.addNewTblW();
        ctTblWidth.setW(BigInteger.valueOf(width));
        //设置表格宽度为非自动
        ctTblWidth.setType(STTblWidth.DXA);

        //创建内容
        for (int j = 0; j < content.size(); j++) {
            List<String> row = content.get(j);
            //行对象
            XWPFTableRow xwpfTableRow = xwpfTable.getRow(j);
            //行高
            xwpfTableRow.setHeight(350);
            //每行轮换颜色
            boolean colorFlag = false;
            if (j % 2 == 0) {
                colorFlag = true;
            }
            boolean lastLineFlag = false;
            if (j == content.size() - 1) {
                lastLineFlag = true;
            }
            boolean firstLineFlag = false;
            if (j == 0) {
                firstLineFlag = true;
            }
            for (int i = 0; i < row.size() && i < tableWidths.length; i++) {
                //单元格对象
                XWPFTableCell xwpfTableCell = xwpfTableRow.getCell(i);
                //每三列为蓝字左对齐，后面为黑字居中对齐
                String baseColor = colorWrite;
                String wordColor = colorBlack;
                ParagraphAlignment align;
                if (i % 3 == 0) {
                    wordColor = colorBlue;
                    align = ParagraphAlignment.LEFT;
                    if (colorFlag) {
                        baseColor = colorLightBlue;
                    }
                } else {
                    wordColor = colorBlack;
                    align = ParagraphAlignment.CENTER;
                    baseColor = colorWrite;
                }
                //添加文本，9号/黑体/左对齐
                buildCell(xwpfTableCell, row.get(i), 8, false, wordColor, align, tableWidths[i], null);
                //添加边框
                if (firstLineFlag) {
                    addTopBorder(xwpfTableCell, 1, colorBlue, true);
                } else {
                    addBottomBorder(xwpfTableCell, 1, colorBlue, lastLineFlag);
                }
                //添加背景色，蓝色
                xwpfTableCell.setColor(baseColor);
            }
        }
    }


    public void buildTable(XWPFDocument document, List<Object> title, List<List<Object>> content, Long[] tableWidths
            , long width, Integer colspanCell, Integer colspan)
            throws InvalidFormatException, IOException, URISyntaxException {
        int rowNum = content.size();
        int columnNum = tableWidths.length;
        if (CollectionUtils.isNotEmpty(title)) {
            rowNum++;
        }
        XWPFTable xwpfTable = document.createTable(rowNum, columnNum);
        //表格居中显示
        CTTblPr ctTblPr = xwpfTable.getCTTbl().addNewTblPr();
        ctTblPr.addNewJc().setVal(STJc.CENTER);
        //设置表格宽度
        CTTblWidth ctTblWidth = ctTblPr.addNewTblW();
        ctTblWidth.setW(BigInteger.valueOf(width));
        //设置表格宽度为非自动
        ctTblWidth.setType(STTblWidth.DXA);

        /**第一行中第三列第四列合并单元格*/
        mergeCellsHorizontal(xwpfTable, 0, colspanCell, colspanCell + colspan - 1);

        //创建标题
        for (int i = 0; i < title.size() && i < tableWidths.length; i++) {
            //行对象
            XWPFTableRow xwpfTableRow = xwpfTable.getRow(0);
            //行高度
            xwpfTableRow.setHeight(600);
            //单元格对象
            XWPFTableCell xwpfTableCell = xwpfTableRow.getCell(i);
            buildCell(xwpfTableCell, title.get(i), 10, true, colorBlack, ParagraphAlignment.CENTER, null, true);
            //添加边框
            addTopBorder(xwpfTableCell, 2, colorBlack, true);
        }

        //创建内容
        for (int j = 0; j < content.size(); j++) {
            List<Object> row = content.get(j);
            //行对象
            XWPFTableRow xwpfTableRow = xwpfTable.getRow(j + 1);
            //设置行高度
            xwpfTableRow.setHeight(1247);

            boolean lastLineFlag = false;
            if (j == content.size() - 1) {
                lastLineFlag = true;
            }
            boolean firstLineFlag = false;
            if (j == 0) {
                firstLineFlag = true;
            }
            for (int i = 0; i < row.size() && i < tableWidths.length; i++) {
                //第一列居左
                ParagraphAlignment align = ParagraphAlignment.CENTER;
                if (i == 0) {
                    align = ParagraphAlignment.LEFT;
                }
                //最后一列要特殊处理金额为蓝色，单位为黑色
                XWPFTableCell xwpfTableCell = xwpfTableRow.getCell(i);
                //设置上下居中
                xwpfTableCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                if (i == row.size() - 1) {
                    String msg = (String) row.get(i);
                    String unit = "元/年";
                    msg = msg.replace("元/年", "");
                    XWPFParagraph xwpfParagraph = xwpfTableCell.getParagraphs().get(0);
                    XWPFRun xwpfRun = xwpfParagraph.createRun();
                    xwpfRun.setText(msg);
                    xwpfRun.setFontSize(8);
                    xwpfRun.setFontFamily("黑体");
                    xwpfRun.setBold(false);
                    xwpfRun.setColor(colorBlue);
                    XWPFRun xwpfRunUnit = xwpfParagraph.createRun();
                    xwpfRunUnit.setText(unit);
                    xwpfRunUnit.setFontSize(8);
                    xwpfRunUnit.setFontFamily("黑体");
                    xwpfRunUnit.setBold(false);
                    //段落位置
                    xwpfParagraph.setAlignment(align);
                    //设置宽度
                    CTTcPr ctTcPr = xwpfTableCell.getCTTc().addNewTcPr();
                    CTTblWidth ctTblWidthCell = ctTcPr.addNewTcW();
                    ctTblWidthCell.setType(STTblWidth.DXA);
                    ctTblWidthCell.setW(BigInteger.valueOf(tableWidths[i]));

                    xwpfTableCell.setParagraph(xwpfParagraph);

                } else {
                    buildCell(xwpfTableCell, row.get(i), 8, false, colorBlack, align, tableWidths[i], true);
                }

                //添加边框
                if (firstLineFlag) {
                    addTopBorder(xwpfTableCell, 1, colorBlack, true);
                } else {
                    addBottomBorder(xwpfTableCell, 2, colorBlack, lastLineFlag);
                }
            }
        }


    }

    /**
     * 创建单元格（指定文本、字体大小、是否加粗、颜色、宽度）
     *
     * @param cell
     * @param value
     * @param fontSize
     * @param bold
     * @param color
     */
    //public void buildCell(XWPFTableCell cell, String value, int fontSize
    //        , Boolean bold, String color, Long width, ParagraphAlignment align) {
    //    cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
    //    //单元格宽度
    //    if (width != null) {
    //        CTTcPr ctTcPr = cell.getCTTc().addNewTcPr();
    //        CTTblWidth ctTblWidthCell = ctTcPr.addNewTcW();
    //        ctTblWidthCell.setType(STTblWidth.DXA);
    //        ctTblWidthCell.setW(BigInteger.valueOf(width));
    //    }
    //    XWPFParagraph xwpfParagraph = cell.getParagraphs().get(0);
    //    buildParagraph(xwpfParagraph, value, fontSize, bold, color);
    //    xwpfParagraph.setAlignment(align);
    //    cell.setParagraph(xwpfParagraph);
    //}


    /**
     * 创建单元格 (对象，可以是String也可以是Image,指定字体，水平居...)
     *
     * @param cell
     * @param value
     * @param fontSize
     * @param bold
     * @param color
     * @param align
     * @throws IOException
     * @throws InvalidFormatException
     * @throws URISyntaxException
     */
    public void buildCell(XWPFTableCell cell, Object value, Integer fontSize, Boolean bold, String color
            , ParagraphAlignment align, Long width, Boolean mediate) throws IOException, InvalidFormatException, URISyntaxException {
        CTTc cttc = cell.getCTTc();
        CTTcPr ctPr = cttc.addNewTcPr();
        /** 水平居中 */
        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        if (mediate != null && mediate == true) {
            /** 竖直居中 */
            ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
            cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
        }
        /**单元格宽度*/
        if (width != null) {
            CTTblWidth ctTblWidthCell = ctPr.addNewTcW();
            ctTblWidthCell.setType(STTblWidth.DXA);
            ctTblWidthCell.setW(BigInteger.valueOf(width));
        }
        if (value instanceof String) {
            XWPFParagraph paragraph = cell.getParagraphs().get(0);
            buildParagraph(paragraph, align, (String) value, fontSize, bold, color);
            cell.setParagraph(paragraph);
        } else if (value instanceof StringBuffer) {
            String imageName = value.toString().substring(value.toString().lastIndexOf("/") + 1);
            XWPFRun pictureRun = cell.getParagraphs().get(0).createRun();
            FileInputStream is = null;
            try {
                // todo 111
//                is = new FileInputStream(new File(this.getClass().getResource(((StringBuffer) value).toString()).toURI()));
                is = new FileInputStream(new File(((StringBuffer) value).toString()));
                if ("image2.jpg".equals(imageName)) {
                    pictureRun.addPicture(is, Document.PICTURE_TYPE_PNG, null, Units.toEMU(37.5), Units.toEMU(15));
                } else {
                    pictureRun.addPicture(is, Document.PICTURE_TYPE_PNG, null, Units.toEMU(50), Units.toEMU(50));
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
     * 添加底边边框
     *
     * @param cell
     * @param width
     * @param baseColor
     * @param lastLineFlag
     */
    public void addBottomBorder(XWPFTableCell cell, long width, String baseColor, boolean lastLineFlag) {
        CTTc ctTc = cell.getCTTc();
        CTTcPr tcPr = ctTc.addNewTcPr();
        CTTcBorders border = tcPr.addNewTcBorders();
        if (lastLineFlag) {
            //只剩下边框
            border.addNewBottom().setVal(STBorder.SINGLE);
            border.getBottom().setColor(baseColor);
            border.getBottom().setSz(BigInteger.valueOf(width));

            border.addNewTop().setVal(STBorder.NIL);
            border.addNewRight().setVal(STBorder.NIL);
            border.addNewLeft().setVal(STBorder.NIL);
        } else {
            //隐藏边框
            border.addNewTop().setVal(STBorder.NIL);
            border.addNewRight().setVal(STBorder.NIL);
            border.addNewLeft().setVal(STBorder.NIL);
            border.addNewBottom().setVal(STBorder.NIL);
        }
    }

    /**
     * 添加顶边边框
     *
     * @param cell
     * @param width
     * @param baseColor
     * @param firstLineFlag
     */
    public void addTopBorder(XWPFTableCell cell, long width, String baseColor, boolean firstLineFlag) {
        CTTc ctTc = cell.getCTTc();
        CTTcPr tcPr = ctTc.addNewTcPr();
        CTTcBorders border = tcPr.addNewTcBorders();
        if (firstLineFlag) {

            //只剩上边框
            border.addNewTop().setVal(STBorder.SINGLE);
            border.getTop().setColor(baseColor);
            border.getTop().setSz(BigInteger.valueOf(width));

            border.addNewRight().setVal(STBorder.NIL);
            border.addNewLeft().setVal(STBorder.NIL);
            border.addNewBottom().setVal(STBorder.NIL);
        } else {
            //隐藏边框
            border.addNewTop().setVal(STBorder.NIL);
            border.addNewRight().setVal(STBorder.NIL);
            border.addNewLeft().setVal(STBorder.NIL);
            border.addNewBottom().setVal(STBorder.NIL);
        }
    }


    /**
     * 隐藏表格边框
     *
     * @param table
     */
    public void displayBorder(XWPFTable table) {
        CTTblBorders ctTblBorders = table.getCTTbl().getTblPr().addNewTblBorders();

        CTBorder leftBorder = ctTblBorders.addNewLeft();
        leftBorder.setVal(STBorder.NIL);
        ctTblBorders.setLeft(leftBorder);

        CTBorder rBorder = ctTblBorders.addNewRight();
        rBorder.setVal(STBorder.NIL);
        ctTblBorders.setRight(rBorder);

        CTBorder tBorder = ctTblBorders.addNewTop();
        tBorder.setVal(STBorder.NIL);
        ctTblBorders.setTop(tBorder);

        CTBorder bBorder = ctTblBorders.addNewBottom();
        bBorder.setVal(STBorder.NIL);
        ctTblBorders.setBottom(bBorder);

        CTBorder vBorder = ctTblBorders.addNewInsideV();
        vBorder.setVal(STBorder.NIL);
        ctTblBorders.setInsideV(vBorder);

        CTBorder hBorder = ctTblBorders.addNewInsideH();
        hBorder.setVal(STBorder.NIL);
        ctTblBorders.setInsideH(hBorder);
    }

    /**
     * 自定义边框
     * 隐藏表格边框（除上下边框）
     *
     * @param table
     */
    public void customizeBorder(XWPFTable table, String color) {
        CTTblBorders ctTblBorders = table.getCTTbl().getTblPr().addNewTblBorders();

        CTBorder leftBorder = ctTblBorders.addNewLeft();
        leftBorder.setVal(STBorder.NIL);
        ctTblBorders.setLeft(leftBorder);

        CTBorder rBorder = ctTblBorders.addNewRight();
        rBorder.setVal(STBorder.NIL);
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
        vBorder.setVal(STBorder.THICK);
        ctTblBorders.setInsideV(vBorder);

        CTBorder hBorder = ctTblBorders.addNewInsideH();
        hBorder.setVal(STBorder.THICK);
        ctTblBorders.setInsideH(hBorder);
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
    public void buildParagraph(XWPFParagraph paragraph, ParagraphAlignment align, String content, int fontSize
            , Boolean bold, String color) {
        paragraph.setAlignment(align);
        buildParagraph(paragraph, content, fontSize, bold, color);
    }

    /**
     * 构建文本、字体大小、是否加粗、颜色
     *
     * @param paragraph
     * @param content
     * @param fontSize
     * @param bold
     * @param color
     */
    public void buildParagraph(XWPFParagraph paragraph, String content, int fontSize
            , Boolean bold, String color) {
        XWPFRun xwpfRun = paragraph.createRun();
        xwpfRun.setText(content);
        xwpfRun.setFontSize(fontSize);
        xwpfRun.setFontFamily("黑体");
        if (bold != null) {
            xwpfRun.setBold(bold);
        }
        if (color != null) {
            xwpfRun.setColor(color);
        }
    }


    /**
     * 创建段落文字
     *
     * @param document
     * @param content
     * @throws DocumentException
     */
    public void buildParagraph(XWPFDocument document, String content, int fontSize) {
        blankParagraph(document);

        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun xwpfRun = paragraph.createRun();
        xwpfRun.setText("    " + content);
        xwpfRun.setFontSize(fontSize);
        xwpfRun.setFontFamily("黑体");
        xwpfRun.setColor(colorBlack);
    }


    /**
     * 分割线
     *
     * @param document
     * @param titleName
     * @param content
     */
//    public void buildParagraph(XWPFDocument document, String titleName, String content) throws InvalidFormatException, IOException, URISyntaxException {
//        try {
//            XWPFParagraph titleParagraph = document.createParagraph();
//            XWPFRun titleRun = titleParagraph.createRun();
//            titleRun.setText("");
//            titleRun.setText(titleName);
//            titleRun.setBold(true);
//            titleRun.setFontFamily("黑体");
//            titleRun.setColor(colorBlue);
//
//            XWPFRun pictureRun = document.createParagraph().createRun();
//            FileInputStream is = null;
//            try {
//                is = new FileInputStream(new File(this.getClass().getResource("/target/classes/static/image/line.png").toURI()));
//                pictureRun.addPicture(is, Document.PICTURE_TYPE_PNG, null, Units.toEMU(420), Units.toEMU(6));
//            } catch (IOException e) {
//                throw e;
//            } finally {
//                if (is != null) {
//                    is.close();
//                }
//            }
//            XWPFParagraph xwpfParagraph = document.createParagraph();
//            XWPFRun xwpfRun = xwpfParagraph.createRun();
//            xwpfRun.setText("    ");
//            xwpfRun.setText(content);
//            xwpfRun.setFontSize(9);
//            xwpfRun.setFontFamily("黑体");
//        } catch (Exception e) {
//            throw e;
//        }
//    }

    /**
     * 空白行
     *
     * @param document
     */
    public void blankParagraph(XWPFDocument document) {
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("           ");
    }


    /**
     * 分页
     *
     * @param document
     */
    public void pageBreak(XWPFDocument document) {
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
    public void mergeCellsHorizontal(XWPFTable table, int row, int fromCell, int toCell) {
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
    public void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
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
}
