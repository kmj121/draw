package com.example.draw.utils;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;
import com.itextpdf.text.pdf.draw.LineSeparator;
import org.apache.commons.lang.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.List;
import java.util.Map;

public abstract class PdfUtils {

    private static Logger logger = LoggerFactory.getLogger(PdfUtils.class);

    //构建字体黑体
    protected static BaseFont baseFont = null;
    //浅蓝色
    protected static BaseColor colorLightBlue = new BaseColor(223, 240, 250);
    //白色
    protected static BaseColor colorWrite = new BaseColor(255, 255, 255);
    //蓝色
    protected static BaseColor colorBlue = new BaseColor(0, 122, 201);
    //灰色
    protected static BaseColor colorGary = new BaseColor(239, 239, 239);
    //灰色
    protected static BaseColor colorGary2 = new BaseColor(173, 175, 175);
    //黑色
    protected static BaseColor colorBlack = new BaseColor(0, 0, 0);
    //绿色
    protected static BaseColor colorGreen = new BaseColor(112, 173, 71);
    //红色
    protected static BaseColor colorRed = new BaseColor(255, 0, 0);
    //标题白字
    protected static Font titleFontWrite = new Font(getBaseFont(), 18, Font.BOLD, colorWrite);
    //副标题白字
    protected static Font subTitleFontWrite = new Font(getBaseFont(), 9, Font.BOLD, colorWrite);
    //段落黑字
    protected static Font fontBlack = new Font(getBaseFont(), 9, Font.NORMAL);
    //副标题蓝字
    protected static Font titleFontBlue = new Font(getBaseFont(), 10.5f, Font.BOLD, colorBlue);
    //表格黑字
    protected static Font tableFontBlack = new Font(getBaseFont(), 8, Font.NORMAL);
    //表格蓝字
    protected static Font tableFontBlue = new Font(getBaseFont(), 8, Font.NORMAL, colorBlue);
    //表格加粗黑字
    protected static Font tableFontBlackBlob = new Font(getBaseFont(), 10, Font.BOLD);
    //表格加粗蓝字
    protected static Font tableFontBlueBlob = new Font(getBaseFont(), 10, Font.BOLD, colorBlue);


    public static BaseFont getBaseFont() {
        if (baseFont != null) {
            return baseFont;
        }
        try {
            baseFont = BaseFont.createFont("/static/font/simhei.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
        } catch (DocumentException e) {
            logger.error(e.getMessage());
        } catch (IOException e) {
            logger.error(e.getMessage());
        }
        return baseFont;
    }

    public void buildPdf(String filePath, Map<String, Object> params, Map<String, List<List<String>>> tableList) throws IOException, DocumentException {
        Document document = new Document();
        PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(filePath));
        document.open();
        this.generatePDF(document, params, tableList);
        document.close();
        writer.close();
    }

    public abstract void generatePDF(Document document, Map<String, Object> params, Map<String, List<List<String>>> tableList) throws DocumentException, IOException;

    /**
     * 构建主标题部分
     *
     * @param document
     * @param name
     * @throws DocumentException
     * @throws IOException
     */
    public void buildTitleTable(Document document, String name) throws DocumentException, IOException {
        PdfPTable table = new PdfPTable(new float[]{40, 82});
        //创建第一列
        PdfPCell pdfPCellImg = new PdfPCell();
        //创建第二列
        PdfPCell pdfPCell = new PdfPCell();
        //隐藏边框
        pdfPCellImg.disableBorderSide(15);
        pdfPCell.disableBorderSide(15);
        //设置列高
        pdfPCellImg.setFixedHeight(36);
        pdfPCell.setFixedHeight(36);
        //上下居中
        pdfPCellImg.setVerticalAlignment(Element.ALIGN_MIDDLE);
        pdfPCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        //添加背景色，蓝色
        pdfPCellImg.setBackgroundColor(colorBlue);
        pdfPCell.setBackgroundColor(colorBlue);
        //添加logo；左对齐
        Image image = Image.getInstance(this.getClass().getResource("/static/image/image1.png"));
        image.setAlignment(Image.ALIGN_LEFT);
//        image.scalePercent(150); //依照比例缩放
        pdfPCellImg.setImage(image);
        table.addCell(pdfPCellImg);

        pdfPCell.setHorizontalAlignment(Element.ALIGN_CENTER);
        //添加文本，字体白色/18号/加粗
        pdfPCell.setPhrase(new Phrase(name, titleFontWrite));
        table.addCell(pdfPCell);

        document.add(table);
    }

    /**
     * 构建主表格
     *
     * @param document
     * @param content
     * @param numColumn
     * @throws DocumentException
     */
    public void buildMainTable(Document document, List<List<String>> content, Integer numColumn) throws DocumentException {
        PdfPTable table = new PdfPTable(numColumn);
        for (List<String> row : content) {
            for (int i = 0; i < row.size() && i < numColumn; i++) {
                //添加文本，9号/黑体/左对齐
                PdfPCell pdfPCell = createCell(row.get(i), fontBlack, Element.ALIGN_LEFT);
                //隐藏边框
                pdfPCell.disableBorderSide(15);
                //添加背景色，蓝色
                pdfPCell.setBackgroundColor(colorGary);
                //设置列高
                pdfPCell.setFixedHeight(16);
                table.addCell(pdfPCell);
            }
        }
        document.add(table);
    }

    /**
     * 创建副标题数据
     *
     * @param document
     * @param titleName
     * @throws DocumentException
     */
    public void buildTitle(Document document, String titleName, String content) throws DocumentException {
        document.add(new Paragraph("\n"));

        //标题文字
        Paragraph paragraph = new Paragraph(titleName, titleFontBlue);
        paragraph.setIndentationLeft(50);
        paragraph.setIndentationRight(50);
        paragraph.setSpacingAfter(-5);
        document.add(paragraph);

        // 直线
        Paragraph paragraphLine = new Paragraph();
        paragraphLine.setIndentationLeft(50);
        paragraphLine.setIndentationRight(50);
        LineSeparator lineSeparator = new LineSeparator();
        lineSeparator.setLineColor(colorGary2);
        lineSeparator.setLineWidth(5);
        lineSeparator.setAlignment(Element.ALIGN_TOP);
        paragraphLine.add(new Chunk(lineSeparator));
        document.add(paragraphLine);

        buildParagraph(document, content, fontBlack);
    }

    /**
     * 创建段落文字
     * @param document
     * @param content
     * @param font
     * @throws DocumentException
     */
    public void buildParagraph(Document document, String content, Font font) throws DocumentException {
        document.add(new Paragraph("\n"));

        Paragraph paragraph = new Paragraph(content, font);
        paragraph.setIndentationLeft(50);
        paragraph.setIndentationRight(50);
        paragraph.setFirstLineIndent(20);
        document.add(paragraph);
    }

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
    public void buildTable(Document document, List<String> title, List<List<String>> content, Integer numColumn, Float width) throws DocumentException {
        buildTable(document, title, null, content, numColumn, width);
    }

    public void buildTable(Document document, List<String> title, List<String> foot, List<List<String>> content, Integer numColumn
            , Float width) throws DocumentException {
        buildTable(document, title, foot, content, numColumn, null, width);
    }

    public void buildTable(Document document, List<String> title, List<String> foot, List<List<String>> content, float[] tableWidths
            , Float width) throws DocumentException {
        buildTable(document, title, foot, content, null, tableWidths, width);
    }

    public void buildTable(Document document, List<String> title, List<String> foot, List<List<String>> content, Integer numColumn
            , float[] tableWidths, Float width) throws DocumentException {
        PdfPTable table;
        if (numColumn != null) {
            table = new PdfPTable(numColumn);
        } else {
            table = new PdfPTable(tableWidths);
            numColumn = tableWidths.length;
        }
//        int[] TableWidths = { 15, 40, 15, 20 };//按百分比分配单元格宽带
//        table.SetWidths(TableWidths);
        table.setTotalWidth(width);//设置绝对宽度  560
        table.setLockedWidth(true);//使绝对宽度模式生效
        table.setSpacingBefore(10);
        //创建标题
        for (int i = 0; i < title.size() && i < numColumn; i++) {
            String msg = title.get(i);
            //添加文本，9号/黑体/白色/居中
            PdfPCell pdfPCell = createCell(msg, subTitleFontWrite, Element.ALIGN_CENTER);
            //添加背景色，蓝色
            if (StringUtils.isNotEmpty(msg)) {
                pdfPCell.setBackgroundColor(colorBlue);
            }
            //设置列高
            pdfPCell.setFixedHeight(16);
            //添加边框
            addBottomBorder(pdfPCell, 1f, colorBlue, true);

            table.addCell(pdfPCell);
        }

        BaseColor baseColor = colorLightBlue;
        //创建内容
        for (int j = 0; j < content.size(); j++) {
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
            for (int i = 0; i < row.size() && i < numColumn; i++) {
                //第一列为蓝字，后面为黑字
                //第一列居左, 后面为居中
                Font font = tableFontBlack;
                int align = Element.ALIGN_CENTER;
                if (i == 0) {
                    font = tableFontBlue;
                    align = Element.ALIGN_LEFT;
                }
                //添加文本，9号/黑体/左对齐
                PdfPCell pdfPCell = createCell(row.get(i), font, align);

                //添加边框
                addBottomBorder(pdfPCell, 1f, colorBlue, lastLineFlag);

                //添加背景色，蓝色
                pdfPCell.setBackgroundColor(baseColor);
                //设置列高
                pdfPCell.setFixedHeight(16);
                table.addCell(pdfPCell);
            }
        }
        //创建尾行
        if (foot != null && foot.size() > 0) {
            for (int i = 0; i < foot.size() && i < numColumn; i++) {
                //第一列为蓝色加粗字，后面为黑色加粗字
                //第一列居左，后面居中
                Font font = tableFontBlackBlob;
                int align = Element.ALIGN_CENTER;
                if (i == 0) {
                    font = tableFontBlueBlob;
                    align = Element.ALIGN_LEFT;
                }
                //添加文本，9号/黑体/左对齐
                PdfPCell pdfPCell = createCell(foot.get(i), font, align);

                //添加边框
                addBottomBorder(pdfPCell, 1f, colorBlue, true);

                //添加背景色，与上一行不同的样式
                if (baseColor.equals(colorLightBlue)) {
                    pdfPCell.setBackgroundColor(colorWrite);
                } else {
                    pdfPCell.setBackgroundColor(colorLightBlue);
                }

                //设置列高
                pdfPCell.setFixedHeight(16);
                table.addCell(pdfPCell);
            }
        }

        document.add(table);
    }

    public void buildTableForScaleInhibitor(Document document, List<String> title, List<List<String>> content, Integer numColumn
            , Float width, boolean colorFlag) throws DocumentException {
        PdfPTable table = new PdfPTable(numColumn);
        table.setTotalWidth(width);//设置绝对宽度  560
        table.setLockedWidth(true);//使绝对宽度模式生效
        table.setSpacingBefore(10);
        //创建标题
        for (int i = 0; i < title.size() && i < numColumn; i++) {
            String msg = title.get(i);
            //添加文本，9号/黑体/白色/居中
            PdfPCell pdfPCell = createCell(msg, subTitleFontWrite, Element.ALIGN_CENTER);
            //添加背景色，蓝色
            if (StringUtils.isNotEmpty(msg)) {
                pdfPCell.setBackgroundColor(colorBlue);
            }
            //设置列高
            pdfPCell.setFixedHeight(32);
            //添加边框
            addBottomBorder(pdfPCell, 1f, colorBlue, true);
            table.addCell(pdfPCell);
        }

        BaseColor baseColor = colorLightBlue;
        //创建内容
        for (int j = 0; j < content.size(); j++) {
            List<String> row = content.get(j);
            //每行轮换颜色
            if (j % 2 == 0) {
                baseColor = colorLightBlue;
            } else {
                baseColor = colorWrite;
            }
            //如果没有尾行则最后一行添加边框
            boolean lastLineFlag = false;
            if (j == content.size() - 1 ) {
                lastLineFlag = true;
            }
            for (int i = 0; i < row.size() && i < numColumn; i++) {
                //第一列为蓝字，后面为黑字
                //第一列居左, 后面为居中
                Font font = tableFontBlack;
                int align = Element.ALIGN_CENTER;
                if (i == 0) {
                    font = tableFontBlue;
                    align = Element.ALIGN_LEFT;
                }
                //添加文本，9号/黑体/左对齐
                PdfPCell pdfPCell = createCell(row.get(i), font, align);

                //添加边框
                addBottomBorder(pdfPCell, 1f, colorBlue, lastLineFlag);
                if(i >= 3 && StringUtils.isNotBlank(row.get(i)) && colorFlag){
                    if(new BigDecimal(row.get(i)).compareTo(new BigDecimal(100))>0){
                        //添加背景色，红色
                        pdfPCell.setBackgroundColor(colorRed);
                    }else {
                        //添加背景色，绿色
                        pdfPCell.setBackgroundColor(colorGreen);
                    }
                } else {
                    //添加背景色，蓝色
                    pdfPCell.setBackgroundColor(baseColor);
                }

                //设置列高
                pdfPCell.setFixedHeight(32);
                table.addCell(pdfPCell);
            }
        }
        document.add(table);
    }

    /**
     * 构建无标题表格
     *
     * @param document
     * @param content
     * @param TableWidths
     * @param width
     * @throws DocumentException
     */
    public void buildTable(Document document, List<List<String>> content, float[] TableWidths, Float width)
            throws DocumentException {
        PdfPTable table = new PdfPTable(TableWidths);
//        int[] TableWidths = { 15, 40, 15, 20 };//按百分比分配单元格宽带
//        table.SetWidths(TableWidths);
        table.setTotalWidth(width);//设置绝对宽度  560
        table.setLockedWidth(true);//使绝对宽度模式生效
        table.setSpacingBefore(10);
        //创建内容
        for (int j = 0; j < content.size(); j++) {
            List<String> row = content.get(j);
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
            for (int i = 0; i < row.size() && i < TableWidths.length; i++) {
                //每三列为蓝字左对齐，后面为黑字居中对齐
                BaseColor baseColor = colorWrite;
                Font font;
                int align;
                if (i % 3 == 0) {
                    font = tableFontBlue;
                    align = Element.ALIGN_LEFT;
                    if (colorFlag) {
                        baseColor = colorLightBlue;
                    }
                } else {
                    font = tableFontBlack;
                    align = Element.ALIGN_CENTER;
                    baseColor = colorWrite;
                }
                //添加文本，9号/黑体/左对齐
                PdfPCell pdfPCell = createCell(row.get(i), font, align);
                //添加边框
                if (firstLineFlag) {
                    addTopBorder(pdfPCell, 1f, colorBlue, true);
                } else {
                    addBottomBorder(pdfPCell, 1f, colorBlue, lastLineFlag);
                }

                //添加背景色，蓝色
                pdfPCell.setBackgroundColor(baseColor);
                //设置列高
                pdfPCell.setFixedHeight(16);
                table.addCell(pdfPCell);
            }
        }
        document.add(table);
    }

    public void buildTable(Document document, List<Object> title, List<List<Object>> content, float[] TableWidths
            , Float width, Integer colspanCell, Integer colspan)
            throws DocumentException {
        PdfPTable table = new PdfPTable(TableWidths);
//        int[] TableWidths = { 15, 40, 15, 20 };//按百分比分配单元格宽带
//        table.SetWidths(TableWidths);
        table.setTotalWidth(width);//设置绝对宽度  560
        table.setLockedWidth(true);//使绝对宽度模式生效
        table.setSpacingBefore(10);

        //创建标题
        for (int i = 0; i < title.size() && i < TableWidths.length; i++) {
            PdfPCell pdfPCell = createCell(title.get(i), tableFontBlackBlob, Element.ALIGN_CENTER);
            //设置列高
            pdfPCell.setFixedHeight(32);
            //添加边框
            addTopBorder(pdfPCell, 1.5f, colorBlack, true);
            if (i == colspanCell) {
                //合并单元格
                pdfPCell.setColspan(colspan);
            }

            table.addCell(pdfPCell);
        }

        //创建内容
        for (int j = 0; j < content.size(); j++) {
            List<Object> row = content.get(j);
            boolean lastLineFlag = false;
            if (j == content.size() - 1) {
                lastLineFlag = true;
            }
            boolean firstLineFlag = false;
            if (j == 0) {
                firstLineFlag = true;
            }
            for (int i = 0; i < row.size() && i < TableWidths.length; i++) {
                //第一列居左
                int align = Element.ALIGN_CENTER;
                if (i == 0) {
                    align = Element.ALIGN_LEFT;
                }
                //最后一列要特殊处理金额为蓝色，单位为黑色
                PdfPCell pdfPCell;
                if (i == row.size() - 1) {
                    String msg = (String) row.get(i);
                    String unit = "元/年";
                    msg = msg.replace("元/年", "");
                    Chunk chunk1 = new Chunk(msg, tableFontBlue);
                    Chunk chunk2 = new Chunk(unit, tableFontBlack);
                    Paragraph paragraph = new Paragraph();
                    paragraph.add(chunk1);
                    paragraph.add(chunk2);
                    pdfPCell = new PdfPCell();
                    pdfPCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
                    pdfPCell.setHorizontalAlignment(align);
                    pdfPCell.setPhrase(paragraph);
                } else {
                    pdfPCell = createCell(row.get(i), tableFontBlack, align);
                }

                //添加边框
                if (firstLineFlag) {
                    addTopBorder(pdfPCell, 1f, colorBlack, true);
                } else {
                    addBottomBorder(pdfPCell, 1.5f, colorBlack, lastLineFlag);
                }

                //设置列高
                pdfPCell.setFixedHeight(45);
                table.addCell(pdfPCell);
            }
        }
        document.add(table);
    }

    /**
     * 创建单元格（指定字体、水平..）
     *
     * @param value
     * @param font
     * @param align
     * @return
     */
    public PdfPCell createCell(String value, Font font, int align) {
        PdfPCell cell = new PdfPCell();
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        cell.setHorizontalAlignment(align);
        cell.setPhrase(new Phrase(value, font));
        return cell;
    }

//    /**
//     * 创建单元格（指定字体、水平居..、单元格跨x列合并）
//     *
//     * @param value
//     * @param font
//     * @param align
//     * @param colspan
//     * @return
//     */
//    public PdfPCell createCell(String value, Font font, int align, int colspan) {
//        PdfPCell cell = new PdfPCell();
//        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
//        cell.setHorizontalAlignment(align);
//        cell.setColspan(colspan);
//        cell.setPhrase(new Phrase(value, font));
//        return cell;
//    }
//
//    /**
//     * 创建单元格（指定字体、水平居..、单元格跨x列合并、设置单元格内边距）
//     *
//     * @param value
//     * @param font
//     * @param align
//     * @param colspan
//     * @param boderFlag
//     * @return
//     */
//    public PdfPCell createCell(String value, Font font, int align, int colspan, boolean boderFlag) {
//        PdfPCell cell = new PdfPCell();
//        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
//        cell.setHorizontalAlignment(align);
//        cell.setColspan(colspan);
//        cell.setPhrase(new Phrase(value, font));
//        cell.setPadding(3.0f);
//        if (!boderFlag) {
//            cell.setBorder(0);
//            cell.setPaddingTop(15.0f);
//            cell.setPaddingBottom(8.0f);
//        } else if (boderFlag) {
//            cell.setBorder(0);
//            cell.setPaddingTop(0.0f);
//            cell.setPaddingBottom(15.0f);
//        }
//        return cell;
//    }
//
//    /**
//     * 创建单元格（指定字体、水平..、边框宽度：0表示无边框、内边距）
//     *
//     * @param value
//     * @param font
//     * @param align
//     * @param borderWidth
//     * @param paddingSize
//     * @param flag
//     * @return
//     */
//    public PdfPCell createCell(String value, Font font, int align, float[] borderWidth, float[] paddingSize, boolean flag) {
//        PdfPCell cell = new PdfPCell();
//        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
//        cell.setHorizontalAlignment(align);
//        cell.setPhrase(new Phrase(value, font));
//        cell.setBorderWidthLeft(borderWidth[0]);
//        cell.setBorderWidthRight(borderWidth[1]);
//        cell.setBorderWidthTop(borderWidth[2]);
//        cell.setBorderWidthBottom(borderWidth[3]);
//        cell.setPaddingTop(paddingSize[0]);
//        cell.setPaddingBottom(paddingSize[1]);
//        if (flag) {
//            cell.setColspan(2);
//        }
//        return cell;
//    }

    /**
     * 创建单元格 (对象，可以是String也可以是Image,指定字体，水平居...)
     *
     * @param value
     * @param font
     * @param align
     * @return
     */
    public PdfPCell createCell(Object value, Font font, int align) {
        PdfPCell pdfPCell = new PdfPCell();
        if (value instanceof String) {
            pdfPCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            pdfPCell.setHorizontalAlignment(align);
            pdfPCell.setPhrase(new Phrase((String) value, font));
        } else if (value instanceof Image) {
            pdfPCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            pdfPCell.setHorizontalAlignment(align);
            pdfPCell.setImage((Image) value);
        }
        return pdfPCell;
    }

    /**
     * 添加底边边框
     *
     * @param pdfPCell
     * @param width
     * @param baseColor
     * @param lastLineFlag
     */
    public void addBottomBorder(PdfPCell pdfPCell, float width, BaseColor baseColor, boolean lastLineFlag) {
        if (lastLineFlag) {
            //只剩下边框
            pdfPCell.disableBorderSide(13);
            pdfPCell.setBorderWidth(width);
            pdfPCell.setBorderColor(baseColor);
        } else {
            //隐藏边框
            pdfPCell.disableBorderSide(15);
        }
    }

    /**
     * 添加顶边边框
     *
     * @param pdfPCell
     * @param width
     * @param baseColor
     * @param firstLineFlag
     */
    public void addTopBorder(PdfPCell pdfPCell, float width, BaseColor baseColor, boolean firstLineFlag) {
        if (firstLineFlag) {
            //只剩上边框
            pdfPCell.disableBorderSide(14);
            pdfPCell.setBorderWidth(width);
            pdfPCell.setBorderColor(baseColor);
        } else {
            //隐藏边框
            pdfPCell.disableBorderSide(15);
        }
    }
}
