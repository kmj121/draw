package com.example.draw.utils;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;

import java.io.IOException;
import java.util.Date;


public class PdfHeaderFooterEvent extends PdfPageEventHelper {

    //总页码使用的模板对象
    public PdfTemplate totalNumTemplate = null;

//    private final static String FONT_PATH = "/Users/jiyi/business/zhiliao/code/zhiliaobeidiao/lab/outputdir/zaozigongfangshujianti.ttf,0";
    private final static String FONT_PATH = "/Users/kanmeijie/Workspace/draw/src/main/resources/static/font/simhei.ttf";

    private final static BaseFont BASE_FONT = initFont();

//    private final static String logoPath = "/Users/jiyi/business/zhiliao/code/zhiliaobeidiao/lab/outputdir/1682682397523.jpg";
    private final static String logoPath = "/Users/kanmeijie/Workspace/draw/src/main/resources/static/image/image3.png";

    public static BaseFont initFont() {
        try{
            return BaseFont.createFont(FONT_PATH, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
        }catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 重写页面结束时间  分别添加页眉、页脚
     */
    @Override
    public void onEndPage(PdfWriter writer, Document docment){
        try{
            this.addPageHeader(writer, docment);
        }catch(Exception e){
            System.out.println("添加页眉出错:" + e);
        }

        try{
            this.addPageFooter(writer, docment);
        }catch(Exception e){
            System.out.println("添加页脚出错" + e);
        }
    }

    /**
     * 页眉
     */
    private void addPageHeader(PdfWriter writer, Document docment) throws BadElementException, IOException {
        //创建字体
        Font textFont = new Font(BASE_FONT, 10f);

        //两列  一列logo  一列项目简称
        PdfPTable table = new PdfPTable(1);
        //设置表格宽度 A4纸宽度减去两个边距  比如我一边30  所以减去60
        table.setTotalWidth(PageSize.A4.getWidth()-60);

        //logo
        //创建图片对象
//        Image logo = null;
//        try {
//            logo = Image.getInstance(logoPath);
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//        //创建一个Phrase对象 再添加一个Chunk对象进去  Chunk里边是图片
//        Phrase logoP = new Phrase("", textFont);
//        //自己调整偏移值 主要是y轴值
//        logoP.add(new Chunk(logo, 0, -35));
//        PdfPCell logoCell = new PdfPCell(logoP);
        PdfPCell logoCell = new PdfPCell();
        Image image = Image.getInstance(this.getClass().getResource("/static/image/image3.png"));
        image.setAlignment(Image.ALIGN_RIGHT);
        logoCell.setImage(image);
        //只保留底部边框和设置高度
        logoCell.disableBorderSide(13);
        logoCell.setFixedHeight(40);
        logoCell.setVerticalAlignment(Element.ALIGN_RIGHT);
        table.addCell(logoCell);

//        Phrase nameP = new Phrase("TEST", textFont);
//        PdfPCell nameCell = new PdfPCell(nameP);
//        //只保留底部边框和设置高度 设置水平居右和垂直居中
//        nameCell.disableBorderSide(13);
//        nameCell.setFixedHeight(40);
//        nameCell.setHorizontalAlignment(Element.ALIGN_RIGHT);
//        nameCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
//        table.addCell(logoCell);

        //再把表格写到页眉处  使用绝对定位
        table.writeSelectedRows(0, -1, 30,  PageSize.A4.getHeight()-20, writer.getDirectContent());
    }

    /**
     * 页脚
     */
    private void addPageFooter(PdfWriter writer, Document docment){
        //创建字体
        Font textFont = new Font(BASE_FONT, 10f);

        //三列  一列导出人  一列页码   一列时间
        PdfPTable table = new PdfPTable(3);
        //设置表格宽度 A4纸宽度减去两个边距  比如我一边30  所以减去60
        table.setTotalWidth(PageSize.A4.getWidth()-60);
        //仅保留顶部边框
        table.getDefaultCell().disableBorderSide(14);
        table.getDefaultCell().setFixedHeight(40);
        table.getDefaultCell().setVerticalAlignment(Element.ALIGN_MIDDLE);

        //导出人
        table.addCell(new Phrase("admin", textFont));

        //页码
        //初始化总页码模板
        if(null == totalNumTemplate){
            totalNumTemplate = writer.getDirectContent().createTemplate(30, 16);
        }
        //再嵌套一个表格 一左一右  左边当前页码 右边总页码
        PdfPTable pageNumTable = new PdfPTable(2);
        try {
            pageNumTable.setTotalWidth(new float[]{80f, 80f});
        } catch (Exception e) {
            e.printStackTrace();
        }
        pageNumTable.setLockedWidth(true);
        pageNumTable.setPaddingTop(-5f);
        //第一列居右
        pageNumTable.getDefaultCell().disableBorderSide(15);
        pageNumTable.getDefaultCell().setFixedHeight(16);
        pageNumTable.getDefaultCell().setHorizontalAlignment(Element.ALIGN_RIGHT);
        pageNumTable.getDefaultCell().setVerticalAlignment(Element.ALIGN_BOTTOM);
        pageNumTable.addCell(new Phrase(writer.getPageNumber()+" / ", textFont));
        //第二列居左
        Image totalNumImg = null;
        try {
            totalNumImg = Image.getInstance(totalNumTemplate);
        } catch (Exception e) {
            e.printStackTrace();
        }
        totalNumImg.setPaddingTop(-5f);
        pageNumTable.getDefaultCell().setPaddingTop(-18f);
        pageNumTable.getDefaultCell().setHorizontalAlignment(Element.ALIGN_LEFT);
        pageNumTable.getDefaultCell().setVerticalAlignment(Element.ALIGN_TOP);
        pageNumTable.addCell(totalNumImg);
        //把页码表格添加到页脚表格
        table.addCell(pageNumTable);

        //日期
        table.addCell(new Phrase(new Date().toString(), textFont));

        //再把表格写到页脚处  使用绝对定位
        table.writeSelectedRows(0, -1, 30, 40, writer.getDirectContent());
    }
    /**
     * 文档关闭事件
     */
    @Override
    public void onCloseDocument(PdfWriter writer, Document docment){
        //创建字体
        Font textFont = new Font(BASE_FONT, 10f);
        //将最后的页码写入到总页码模板
//        String totalNum = writer.getPageNumber() + "页";
//        totalNumTemplate.beginText();
//        totalNumTemplate.setFontAndSize(BASE_FONT, 5f);
//        totalNumTemplate.showText(totalNum);
//        totalNumTemplate.setHeight(16f);
//        totalNumTemplate.endText();
        totalNumTemplate.closePath();
    }


}