package com.example.draw.utils;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import org.apache.commons.lang.StringUtils;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class ProgramsPdfUtils extends PdfUtils {

    @Override
    public void generatePDF(Document document, Map<String, Object> params, Map<String, List<List<String>>> tableMap) throws DocumentException, IOException {
        buildTitleTable(document, "方案建议书");

        buildMainTable(document
                , "111"
                , "222"
                , "333"
                , "444");

        buildHomePage(document);




//        if (params.containsKey("${lightFlag}")) {
//            buildTitle(document, "背景", "在微电子行业，膜分离技术是制备超纯水和废水回用工艺中不可替代的主流技术。工业回用水中普遍存在各种有机物、无机物和微生物。这些物质与微生物本身产生的粘液杂混在一起形成生物粘泥，极易造成膜表面污染。微生物污染积累迅速，不仅造成膜分离装置产水流量和质量降低、增加系统操作压力进而导致能耗增加等问题，还因为频繁化学清洗消耗大量人力并增加运行费用，甚至会缩短膜的使用寿命。");
//
//            buildBackendParaL(document);
//
//            buildTitle(document, "水质分析", "为进一步了解水中可能存在的污染物，对保安过滤器进水进行水质分析，结果见下：");
//
//            buildWaterTableL(document
//                    , params.get("${aluminum}").toString()
//                    , params.get("${iron}").toString()
//                    , params.get("${silica}").toString()
//                    , params.get("${copper}").toString()
//                    , params.get("${totalBacteriaCount}").toString()
//                    , params.get("${ph}").toString()
//                    , params.get("${conductivity}").toString()
//                    , params.get("${temperature}").toString()
//                    , params.get("${chemicalOxygenD}").toString()
//                    , params.get("${totalOrganicC}").toString()
//                    , params.get("${turbidity}").toString());
//            //            todo
////            buildWaterTableL(document, "", "", "", "", "", "", "", "", "", "", "");
//
//            buildTitle(document, "系统性能", "反渗透系统运行性能情况见下：");
//
//            buildFunctionTable(document
//                    , params.get("${cfrfValue}").toString()
//                    , params.get("${cipValue}").toString()
//                    , params.get("${ocfValue}").toString());
//            //            todo
////            buildFunctionTable(document, "", "", "");
//
//            document.newPage();
//
//            buildTitle(document, "解决方案", "基于对回用系统水质和运行性能的了解，对该系统进行诊断，推荐以下方案：");
//
//            buildProgramTable(document, tableMap.get("light_product_table"));
//
//            buildTitle(document, "系统性能预测", "使用推荐的化学品方案后的反渗透系统性能预测如下：");
//
//            buildForecastTable(document
//                    , params.get("${cfrfValue}").toString()
//                    , params.get("${cfrfValueNew}").toString()
//                    , params.get("${cipValue}").toString()
//                    , params.get("${cipValueNew}").toString()
//                    , params.get("${ocfValue}").toString()
//                    , params.get("${ocfValueNew}").toString());
//            //            todo
////            buildForecastTable(document, "", "", "", "", "", "");
//        } else if (params.containsKey("${heavyFlag}")) {
//
//            buildTitle(document, "背景", "工业回用水中普遍存在各种有机物、无机物和微生物。这些物质与微生物本身产生的粘液杂混在一起形成生物粘泥，极易造成膜表面污染。微生物污染积累迅速，不仅造成膜分离装置产水流量和质量降低、增加系统操作压力进而导致能耗增加等问题，还因为频繁化学清洗消耗大量人力并增加运行费用，甚至会缩短膜的使用寿命。");
//
//            buildBackendParaH(document);
//
//            buildTitle(document, "水质分析", "为进一步了解水中可能存在的污染物，对系统进水进行水质分析，结果见下。");
//
//            //重工业按杀菌剂，阻垢剂，杀菌剂+阻垢剂  分类
//            if (params.containsKey("${heavyBFlag}")) {
//                //基本表单
//                buildWaterTableHB(document
//                        , params.get("${aluminum}").toString()
//                        , params.get("${ironTotal}").toString()
//                        , params.get("${silica}").toString()
//                        , params.get("${magnesium}").toString()
//                        , params.get("${manganese}").toString()
//                        , params.get("${calcium}").toString()
//                        , params.get("${totalBacteriaCount}").toString()
//                        , params.get("${ph}").toString()
//                        , params.get("${recoveryRate}").toString()
//                        , params.get("${siltDensityIndex}").toString()
//                        , params.get("${chemicalOxygenDemand}").toString());
//
//
//                buildTitle(document, "系统性能", "反渗透系统运行性能情况见下：");
//
//                buildFunctionTable(document
//                        , params.get("${cfrfValue}").toString()
//                        , params.get("${cipValue}").toString()
//                        , params.get("${ocfValue}").toString());
//                //除硅表单
//                if (params.containsKey("${aluminum1}")) {
//                    buildParagraph(document, "除硅预处理后参数内容", fontBlack);
//
//                    buildWaterTableHB(document
//                            , params.get("${aluminum1}").toString()
//                            , params.get("${ironTotal1}").toString()
//                            , params.get("${silica1}").toString()
//                            , params.get("${magnesium1}").toString()
//                            , params.get("${manganese1}").toString()
//                            , params.get("${calcium1}").toString()
//                            , params.get("${totalBacteriaCount1}").toString()
//                            , params.get("${ph1}").toString()
//                            , params.get("${recoveryRate1}").toString()
//                            , params.get("${siltDensityIndex1}").toString()
//                            , params.get("${chemicalOxygenDemand1}").toString());
//
//
//                    buildTitle(document, "系统性能", "反渗透系统运行性能情况见下：");
//
//                    buildFunctionTable(document
//                            , params.get("${cfrfValue1}").toString()
//                            , params.get("${cipValue1}").toString()
//                            , params.get("${ocfValue1}").toString());
//                }
//                //N3108表单
//                if (params.containsKey("${aluminum2}")) {
//
//                    buildParagraph(document, "N3108预处理后参数内容", fontBlack);
//
//                    buildWaterTableHB(document
//                            , params.get("${aluminum2}").toString()
//                            , params.get("${ironTotal2}").toString()
//                            , params.get("${silica2}").toString()
//                            , params.get("${magnesium2}").toString()
//                            , params.get("${manganese2}").toString()
//                            , params.get("${calcium2}").toString()
//                            , params.get("${totalBacteriaCount2}").toString()
//                            , params.get("${ph2}").toString()
//                            , params.get("${recoveryRate2}").toString()
//                            , params.get("${siltDensityIndex2}").toString()
//                            , params.get("${chemicalOxygenDemand2}").toString());
//
//
//                    buildTitle(document, "系统性能", "反渗透系统运行性能情况见下：");
//
//                    buildFunctionTable(document
//                            , params.get("${cfrfValue2}").toString()
//                            , params.get("${cipValue2}").toString()
//                            , params.get("${ocfValue2}").toString());
//                }
//
//            } else if (params.containsKey("${heavySFlag}")) {
//                //基本表单
//                buildWaterTableHS(document
//                        , params.get("${aluminum}").toString()
//                        , params.get("${silica}").toString()
//                        , params.get("${sodium}").toString()
//                        , params.get("${magnesium}").toString()
//                        , params.get("${barium}").toString()
//                        , params.get("${kalium}").toString()
//                        , params.get("${manganese}").toString()
//                        , params.get("${strontium}").toString()
//                        , params.get("${fluorine}").toString()
//                        , params.get("${chlorine}").toString()
//                        , params.get("${bromine}").toString()
//                        , params.get("${calcium}").toString()
//                        , params.get("${sulfate}").toString()
//                        , params.get("${nitrate}").toString()
//                        , params.get("${phosphate}").toString()
//                        , params.get("${bicarbonate}").toString()
//                        , params.get("${ironTotal}").toString()
//                        , params.get("${ferricIon}").toString()
//                        , params.get("${ferrous}").toString()
//                        , params.get("${temperature}").toString()
//                        , params.get("${ph}").toString()
//                        , params.get("${influentFlow}").toString()
//                        , params.get("${recoveryRate}").toString()
//                        , params.get("${chemicalOxygenDemand}").toString()
//                        , params.get("${siltDensityIndex}").toString());
//                //除硅表单
//                if (params.containsKey("${aluminum1}")) {
//                    buildParagraph(document, "除硅预处理后参数内容", fontBlack);
//
//                    buildWaterTableHS(document
//                            , params.get("${aluminum1}").toString()
//                            , params.get("${silica1}").toString()
//                            , params.get("${sodium1}").toString()
//                            , params.get("${magnesium1}").toString()
//                            , params.get("${barium1}").toString()
//                            , params.get("${kalium1}").toString()
//                            , params.get("${manganese1}").toString()
//                            , params.get("${strontium1}").toString()
//                            , params.get("${fluorine1}").toString()
//                            , params.get("${chlorine11}").toString()
//                            , params.get("${bromine1}").toString()
//                            , params.get("${calcium1}").toString()
//                            , params.get("${sulfate1}").toString()
//                            , params.get("${nitrate1}").toString()
//                            , params.get("${phosphate1}").toString()
//                            , params.get("${bicarbonate1}").toString()
//                            , params.get("${ironTotal1}").toString()
//                            , params.get("${ferricIon1}").toString()
//                            , params.get("${ferrous1}").toString()
//                            , params.get("${temperature1}").toString()
//                            , params.get("${ph1}").toString()
//                            , params.get("${influentFlow1}").toString()
//                            , params.get("${recoveryRate1}").toString()
//                            , params.get("${chemicalOxygenDemand1}").toString()
//                            , params.get("${siltDensityIndex1}").toString());
//                }
//
//                //N3108表单
//                if (params.containsKey("${aluminum2}")) {
//                    buildParagraph(document, "N3108预处理后参数内容", fontBlack);
//
//                    buildWaterTableHS(document
//                            , params.get("${aluminum2}").toString()
//                            , params.get("${silica2}").toString()
//                            , params.get("${sodium2}").toString()
//                            , params.get("${magnesium2}").toString()
//                            , params.get("${barium2}").toString()
//                            , params.get("${kalium2}").toString()
//                            , params.get("${manganese2}").toString()
//                            , params.get("${strontium2}").toString()
//                            , params.get("${fluorine2}").toString()
//                            , params.get("${chlorine2}").toString()
//                            , params.get("${bromine2}").toString()
//                            , params.get("${calcium2}").toString()
//                            , params.get("${sulfate2}").toString()
//                            , params.get("${nitrate2}").toString()
//                            , params.get("${phosphate2}").toString()
//                            , params.get("${bicarbonate2}").toString()
//                            , params.get("${ironTotal2}").toString()
//                            , params.get("${ferricIon2}").toString()
//                            , params.get("${ferrous2}").toString()
//                            , params.get("${temperature2}").toString()
//                            , params.get("${ph2}").toString()
//                            , params.get("${influentFlow2}").toString()
//                            , params.get("${recoveryRate2}").toString()
//                            , params.get("${chemicalOxygenDemand2}").toString()
//                            , params.get("${siltDensityIndex2}").toString());
//                }
//
//            } else if (params.containsKey("${heavyBSFlag}")) {
//                //基本表单
//                buildWaterTableHBS(document
//                        , params.get("${aluminum}").toString()
//                        , params.get("${silica}").toString()
//                        , params.get("${sodium}").toString()
//                        , params.get("${magnesium}").toString()
//                        , params.get("${barium}").toString()
//                        , params.get("${kalium}").toString()
//                        , params.get("${manganese}").toString()
//                        , params.get("${strontium}").toString()
//                        , params.get("${fluorine}").toString()
//                        , params.get("${chlorine}").toString()
//                        , params.get("${bromine}").toString()
//                        , params.get("${calcium}").toString()
//                        , params.get("${sulfate}").toString()
//                        , params.get("${nitrate}").toString()
//                        , params.get("${phosphate}").toString()
//                        , params.get("${bicarbonate}").toString()
//                        , params.get("${ironTotal}").toString()
//                        , params.get("${ferricIon}").toString()
//                        , params.get("${ferrous}").toString()
//                        , params.get("${temperature}").toString()
//                        , params.get("${ph}").toString()
//                        , params.get("${influentFlow}").toString()
//                        , params.get("${recoveryRate}").toString()
//                        , params.get("${chemicalOxygenDemand}").toString()
//                        , params.get("${siltDensityIndex}").toString()
//                        , params.get("${totalBacteriaCount}").toString());
//
//
//                buildTitle(document, "系统性能", "反渗透系统运行性能情况见下：");
//
//                buildFunctionTable(document
//                        , params.get("${cfrfValue}").toString()
//                        , params.get("${cipValue}").toString()
//                        , params.get("${ocfValue}").toString());
//
//                //除硅表单
//                if (params.containsKey("${aluminum1}")) {
//                    buildParagraph(document, "除硅预处理后参数内容", fontBlack);
//
//                    buildWaterTableHBS(document
//                            , params.get("${aluminum1}").toString()
//                            , params.get("${silica1}").toString()
//                            , params.get("${sodium1}").toString()
//                            , params.get("${magnesium1}").toString()
//                            , params.get("${barium1}").toString()
//                            , params.get("${kalium1}").toString()
//                            , params.get("${manganese1}").toString()
//                            , params.get("${strontium1}").toString()
//                            , params.get("${fluorine1}").toString()
//                            , params.get("${chlorine1}").toString()
//                            , params.get("${bromine1}").toString()
//                            , params.get("${calcium1}").toString()
//                            , params.get("${sulfate1}").toString()
//                            , params.get("${nitrate1}").toString()
//                            , params.get("${phosphate1}").toString()
//                            , params.get("${bicarbonate1}").toString()
//                            , params.get("${ironTotal1}").toString()
//                            , params.get("${ferricIon1}").toString()
//                            , params.get("${ferrous1}").toString()
//                            , params.get("${temperature1}").toString()
//                            , params.get("${ph1}").toString()
//                            , params.get("${influentFlow1}").toString()
//                            , params.get("${recoveryRate1}").toString()
//                            , params.get("${chemicalOxygenDemand1}").toString()
//                            , params.get("${siltDensityIndex1}").toString()
//                            , params.get("${totalBacteriaCount1}").toString());
//
//
//                    buildTitle(document, "系统性能", "反渗透系统运行性能情况见下：");
//
//                    buildFunctionTable(document
//                            , params.get("${cfrfValue1}").toString()
//                            , params.get("${cipValue1}").toString()
//                            , params.get("${ocfValue1}").toString());
//                }
//
//                //N3108表单
//                if (params.containsKey("${aluminum2}")) {
//                    buildParagraph(document, "N3108预处理后参数内容", fontBlack);
//
//                    buildWaterTableHBS(document
//                            , params.get("${aluminum2}").toString()
//                            , params.get("${silica2}").toString()
//                            , params.get("${sodium2}").toString()
//                            , params.get("${magnesium2}").toString()
//                            , params.get("${barium2}").toString()
//                            , params.get("${kalium2}").toString()
//                            , params.get("${manganese2}").toString()
//                            , params.get("${strontium2}").toString()
//                            , params.get("${fluorine2}").toString()
//                            , params.get("${chlorine2}").toString()
//                            , params.get("${bromine2}").toString()
//                            , params.get("${calcium2}").toString()
//                            , params.get("${sulfate2}").toString()
//                            , params.get("${nitrate2}").toString()
//                            , params.get("${phosphate2}").toString()
//                            , params.get("${bicarbonate2}").toString()
//                            , params.get("${ironTotal2}").toString()
//                            , params.get("${ferricIon2}").toString()
//                            , params.get("${ferrous2}").toString()
//                            , params.get("${temperature2}").toString()
//                            , params.get("${ph2}").toString()
//                            , params.get("${influentFlow2}").toString()
//                            , params.get("${recoveryRate2}").toString()
//                            , params.get("${chemicalOxygenDemand2}").toString()
//                            , params.get("${siltDensityIndex2}").toString()
//                            , params.get("${totalBacteriaCount2}").toString());
//
//                    buildTitle(document, "系统性能", "反渗透系统运行性能情况见下：");
//
//                    buildFunctionTable(document
//                            , params.get("${cfrfValue2}").toString()
//                            , params.get("${cipValue2}").toString()
//                            , params.get("${ocfValue2}").toString());
//                }
//            }
//
//
//            document.newPage();
//
//            buildTitle(document, "解决方案", "基于对回用系统水质和运行性能的了解，对该系统进行诊断，推荐以下方案：");
//
//            if (params.containsKey("${heavy_desilication_flag}")) {
//                buildParagraph(document, "除硅预处理推荐方案内容", fontBlack);
//
//                buildDesilicationTable(document
//                        , params.containsKey("${heroSuggestions}") ? params.get("${heroSuggestions}").toString() : null
//                        , params.containsKey("${n1998SISuggestions}") ? params.get("${n1998SISuggestions}").toString() : null
//                        , params.get("${n1998SIProductName}").toString()
//                        , params.get("${n1998SIProductValue}").toString()
//                        , params.get("${n1998SIAddingPlace}").toString()
//                        , params.get("${n1998SIAddingType}").toString()
//                        , params.get("${sludgeGenerationName}").toString()
//                        , params.get("${sludgeGenerationValue}").toString()
//                        , params.get("${sludgeGenerationUse}").toString()
//                        , params.get("${sludgeGenerationExplain}").toString()
//                        , params.get("${extraCausticNeededName}").toString()
//                        , params.get("${extraCausticNeededValue}").toString()
//                        , params.get("${extraCausticNeededUse}").toString()
//                        , params.get("${extraCausticNeededExplain}").toString());
//            }
//            if (params.containsKey("${heavy_N3108_flag}")) {
//                buildParagraph(document, "N3108预处理推荐方案内容", fontBlack);
//
//                buildN3108Table(document
//                        , params.get("${n3108ProductName}").toString()
//                        , params.get("${n3108Value}").toString()
//                        , params.get("${n3108AddingPlace}").toString()
//                        , params.get("${n3108AddingType}").toString());
//            }
//            if (params.containsKey("${heavy_product_flag}")) {
//                buildParagraph(document, "杀菌剂推荐方案内容", fontBlack);
//
//                buildProgramTable(document, tableMap.get("heavy_product_table"));
//            }
//            if (params.containsKey("${heavy_scale_inhibitor_flag}")) {
//                buildParagraph(document, "阻垢剂推荐方案内容", fontBlack);
//
//                buildScaleInhibitorTable(document
//                        , params.get("${feedCalciteSrValue}") == null ? "" : params.get("${feedCalciteSrValue}").toString()
//                        , params.get("${feedLSIValue}") == null ? "" : params.get("${feedLSIValue}").toString()
//                        , params.get("${concentrationfactorValue}") == null ? "" : params.get("${concentrationfactorValue}").toString()
//                        , params.get("${pHValue}") == null ? "" : params.get("${pHValue}").toString()
//                        , params.get("${calciteSRValue}") == null ? "" : params.get("${calciteSRValue}").toString()
//                        , params.get("${concentrateLSIValue}") == null ? "" : params.get("${concentrateLSIValue}").toString()
//                        , params.get("${caValue}") == null ? "" : params.get("${caValue}").toString()
//                        , params.get("${siO2Value}") == null ? "" : params.get("${siO2Value}").toString()
//                        , params.get("${mgValue}") == null ? "" : params.get("${mgValue}").toString()
//                        , tableMap.get("heavy_scale_inhibitor_table"));
//            }
//
//            if (params.containsKey("${heavy_msg_flag}")) {
//                buildParagraph(document, params.get("${heavy_msg_flag}").toString(), fontBlack);
//            }
//
//        }
    }

    public void buildHomePage(Document document) throws DocumentException, IOException {
        for (int i=0; i<=10; i++) {
            /**
             * 雇前背景调查报告
             */
            PdfPTable table1 = new PdfPTable(new float[]{50});
            //创建第一列
            PdfPCell table1Cell1 = new PdfPCell();
            //隐藏边框
            table1Cell1.disableBorderSide(15);
            //设置列高
            table1Cell1.setFixedHeight(36);
            //上下居中
            table1Cell1.setVerticalAlignment(Element.ALIGN_MIDDLE);
            //添加文本，字体蓝色/18号/加粗
            table1Cell1.setPhrase(new Phrase("雇前背景调查报告", titleFont1));
            table1.addCell(table1Cell1);
            document.add(table1);

            /**
             * 委托日期：2023-01-01
             */
            PdfPTable table2 = new PdfPTable(new float[]{40});
            //创建第一列
            PdfPCell table2Cell1 = new PdfPCell();
            //隐藏边框
            table2Cell1.disableBorderSide(15);
            //设置列高
            table2Cell1.setFixedHeight(36);
            //上下居中
            table2Cell1.setVerticalAlignment(Element.ALIGN_MIDDLE);
            //添加背景色，蓝色
            table2Cell1.setBackgroundColor(colorBlue);
            //添加文本，字体白色/18号/加粗
            table2Cell1.setPhrase(new Phrase("委托日期：2023-01-01", titleFontWrite));
            table2.addCell(table2Cell1);
            document.add(table2);
        }
    }

    public void buildMainTable(Document document, String accountName, String recycleSystemName, String proposalName
            , String createDate) throws DocumentException {
        document.add(new Paragraph("\n"));

        List<List<String>> firstTableList = new ArrayList<>();
        List<String> row1 = new ArrayList<>();
        row1.add("客户名称：" + accountName);
        row1.add("系统名称：" + recycleSystemName);
        List<String> row2 = new ArrayList<>();
        row2.add("方案名称：" + proposalName);
        row2.add("日期：" + createDate);
        firstTableList.add(row1);
        firstTableList.add(row2);
        buildMainTable(document, firstTableList, 2);
    }

    /**
     * 添加背景描述
     *
     * @param document
     */
    public void buildBackendParaL(Document document) throws DocumentException, IOException {
        Paragraph paragraph = new Paragraph("反渗透系统总体运行效率不高，主要表现在以下几个方面：", fontBlack);
        paragraph.setIndentationLeft(50);
        paragraph.setIndentationRight(50);
        paragraph.setFirstLineIndent(20);
        document.add(paragraph);

        paragraph = new Paragraph("●   保安过滤器(超滤和反渗透膜前)滤芯更换频繁", fontBlack);
        paragraph.setIndentationLeft(50);
        paragraph.setIndentationRight(50);
        paragraph.setFirstLineIndent(20);
        document.add(paragraph);


        paragraph = new Paragraph("●   反渗透膜清洗频率高", fontBlack);
        paragraph.setIndentationLeft(50);
        paragraph.setIndentationRight(50);
        paragraph.setFirstLineIndent(20);
        document.add(paragraph);

        paragraph = new Paragraph("●   反渗透膜产水流量低，回收率低于系统设计回收率", fontBlack);
        paragraph.setIndentationLeft(50);
        paragraph.setIndentationRight(50);
        paragraph.setFirstLineIndent(20);
        document.add(paragraph);

        paragraph = new Paragraph("●   反渗透膜压差偏高", fontBlack);
        paragraph.setIndentationLeft(50);
        paragraph.setIndentationRight(50);
        paragraph.setFirstLineIndent(20);
        document.add(paragraph);
    }

    /**
     * 添加背景描述
     *
     * @param document
     */
    public void buildBackendParaH(Document document) throws DocumentException, IOException {
        Paragraph paragraph = new Paragraph("水质中的污染物成分若不进行适当处理和控制，会造成反渗透系统总体运行效率不高，主要表现在以下几个方面：", fontBlack);
        paragraph.setIndentationLeft(50);
        paragraph.setIndentationRight(50);
        paragraph.setFirstLineIndent(20);
        document.add(paragraph);

        paragraph = new Paragraph("●   保安过滤器滤芯更换频繁，增加人工劳动量和系统停机频次", fontBlack);
        paragraph.setIndentationLeft(50);
        paragraph.setIndentationRight(50);
        paragraph.setFirstLineIndent(20);
        document.add(paragraph);


        paragraph = new Paragraph("●   膜系统产水流量低，回收率低于系统设计回收率，系统效率低下", fontBlack);
        paragraph.setIndentationLeft(50);
        paragraph.setIndentationRight(50);
        paragraph.setFirstLineIndent(20);
        document.add(paragraph);

        paragraph = new Paragraph("●   膜系统压差偏高，能耗增加，同时加大了膜性能损坏的风险", fontBlack);
        paragraph.setIndentationLeft(50);
        paragraph.setIndentationRight(50);
        paragraph.setFirstLineIndent(20);
        document.add(paragraph);

        paragraph = new Paragraph("●   膜系统清洗频率高，膜寿命降低", fontBlack);
        paragraph.setIndentationLeft(50);
        paragraph.setIndentationRight(50);
        paragraph.setFirstLineIndent(20);
        document.add(paragraph);
    }

    /**
     * 添加水质分析表格-(轻工业表单)
     *
     * @param document
     * @throws DocumentException
     */
    public void buildWaterTableL(Document document, String aluminum, String iron, String silica, String copper
            , String totalBacteriaCount, String ph, String conductivity, String temperature, String chemicalOxygenD
            , String totalOrganicC, String turbidity) throws DocumentException {
        //构建标题
        List<String> rowTitle = new ArrayList<>();
        rowTitle.add("关键指标");
        rowTitle.add("单位");
        rowTitle.add("数值");
        //构建内容
        List<List<String>> tableList = new ArrayList<>();
        List<String> row1 = new ArrayList<>();
        row1.add("铝");
        row1.add("ppm");
        row1.add(aluminum);
        tableList.add(row1);

        List<String> row2 = new ArrayList<>();
        row2.add("铁");
        row2.add("ppm");
        row2.add(iron);
        tableList.add(row2);

        List<String> row3 = new ArrayList<>();
        row3.add("硅");
        row3.add("ppm");
        row3.add(silica);
        tableList.add(row3);

        List<String> row4 = new ArrayList<>();
        row4.add("铜");
        row4.add("ppm");
        row4.add(copper);
        tableList.add(row4);

        List<String> row5 = new ArrayList<>();
        row5.add("细菌总数");
        row5.add("CFU/ml");
        row5.add(totalBacteriaCount);
        tableList.add(row5);

        List<String> row6 = new ArrayList<>();
        row6.add("pH");
        row6.add("");
        row6.add(ph);
        tableList.add(row6);

        List<String> row7 = new ArrayList<>();
        row7.add("电导率");
        row7.add("μs/cm");
        row7.add(conductivity);
        tableList.add(row7);

        List<String> row8 = new ArrayList<>();
        row8.add("温度");
        row8.add("℃");
        row8.add(temperature);
        tableList.add(row8);

        List<String> row9 = new ArrayList<>();
        row9.add("总有机碳");
        row9.add("ppm");
        row9.add(chemicalOxygenD);
        tableList.add(row9);

        List<String> row10 = new ArrayList<>();
        row10.add("化学需氧量");
        row10.add("ppm");
        row10.add(totalOrganicC);
        tableList.add(row10);

        List<String> row11 = new ArrayList<>();
        row11.add("浊度");
        row11.add("NTU");
        row11.add(turbidity);
        tableList.add(row11);

        buildTable(document, rowTitle, tableList, 3, 320f);
    }

    /**
     * 添加水质分析表格-重工业-杀菌剂
     *
     * @param document
     * @throws DocumentException
     */
    public void buildWaterTableHB(Document document, String aluminum, String ironTotal, String silica, String magnesium
            , String manganese, String calcium, String totalBacteriaCount, String ph, String recoveryRate
            , String siltDensityIndex, String chemicalOxygenDemand) throws DocumentException {
        //构建标题
        List<String> rowTitle = new ArrayList<>();
        rowTitle.add("关键指标");
        rowTitle.add("单位");
        rowTitle.add("数值");
        //构建内容
        List<List<String>> tableList = new ArrayList<>();
        List<String> row1 = new ArrayList<>();
        row1.add("铝");
        row1.add("ppm");
        row1.add(aluminum);
        tableList.add(row1);

        List<String> row2 = new ArrayList<>();
        row2.add("铁");
        row2.add("ppm");
        row2.add(ironTotal);
        tableList.add(row2);

        List<String> row3 = new ArrayList<>();
        row3.add("硅");
        row3.add("ppm");
        row3.add(silica);
        tableList.add(row3);

        List<String> row4 = new ArrayList<>();
        row4.add("镁");
        row4.add("ppm");
        row4.add(magnesium);
        tableList.add(row4);

        List<String> row5 = new ArrayList<>();
        row5.add("锰");
        row5.add("ppm");
        row5.add(manganese);
        tableList.add(row5);

        List<String> row6 = new ArrayList<>();
        row6.add("钙");
        row6.add("ppm");
        row6.add(calcium);
        tableList.add(row6);

        List<String> row7 = new ArrayList<>();
        row7.add("细菌总数");
        row7.add("CFU/ml");
        row7.add(totalBacteriaCount);
        tableList.add(row7);

        List<String> row8 = new ArrayList<>();
        row8.add("pH");
        row8.add("");
        row8.add(ph);
        tableList.add(row8);

        List<String> row9 = new ArrayList<>();
        row9.add("回收率");
        row9.add("%");
        row9.add(recoveryRate);
        tableList.add(row9);

        List<String> row10 = new ArrayList<>();
        row10.add("污染密度指数");
        row10.add("");
        row10.add(siltDensityIndex);
        tableList.add(row10);

        List<String> row11 = new ArrayList<>();
        row11.add("COD");
        row11.add("ppm");
        row11.add(chemicalOxygenDemand);
        tableList.add(row11);

        buildTable(document, rowTitle, tableList, 3, 320f);
    }

    /**
     * 添加水质分析表格-重工业-阻垢剂
     *
     * @param document
     * @throws DocumentException
     */
    public void buildWaterTableHS(Document document, String aluminum, String silica, String sodium, String magnesium
            , String barium, String kalium, String manganese, String strontium, String fluorine
            , String chlorine, String bromine, String calcium, String sulfate, String nitrate, String phosphate
            , String bicarbonate, String ironTotal, String ferricIon, String ferrous, String temperature
            , String ph, String influentFlow, String recoveryRate, String chemicalOxygenDemand, String siltDensityIndex)
            throws DocumentException {
        //构建标题
        List<String> rowTitle = new ArrayList<>();
        rowTitle.add("关键指标");
        rowTitle.add("单位");
        rowTitle.add("数值");
        //构建内容
        List<List<String>> tableList = new ArrayList<>();
        List<String> row1 = new ArrayList<>();
        row1.add("铝");
        row1.add("ppm");
        row1.add(aluminum);
        tableList.add(row1);

        List<String> row2 = new ArrayList<>();
        row2.add("硅");
        row2.add("ppm");
        row2.add(silica);
        tableList.add(row2);

        List<String> row3 = new ArrayList<>();
        row3.add("钠");
        row3.add("ppm");
        row3.add(sodium);
        tableList.add(row3);

        List<String> row4 = new ArrayList<>();
        row4.add("镁");
        row4.add("ppm");
        row4.add(magnesium);
        tableList.add(row4);

        List<String> row5 = new ArrayList<>();
        row5.add("钡");
        row5.add("ppm");
        row5.add(barium);
        tableList.add(row5);

        List<String> row6 = new ArrayList<>();
        row6.add("钾");
        row6.add("ppm");
        row6.add(kalium);
        tableList.add(row6);

        List<String> row7 = new ArrayList<>();
        row7.add("锰");
        row7.add("ppm");
        row7.add(manganese);
        tableList.add(row7);

        List<String> row8 = new ArrayList<>();
        row8.add("锶");
        row8.add("ppm");
        row8.add(strontium);
        tableList.add(row8);

        List<String> row9 = new ArrayList<>();
        row9.add("氟");
        row9.add("ppm");
        row9.add(fluorine);
        tableList.add(row9);

        List<String> row10 = new ArrayList<>();
        row10.add("氯");
        row10.add("ppm");
        row10.add(chlorine);
        tableList.add(row10);

        List<String> row11 = new ArrayList<>();
        row11.add("溴");
        row11.add("ppm");
        row11.add(bromine);
        tableList.add(row11);

        List<String> row12 = new ArrayList<>();
        row12.add("钙");
        row12.add("ppm");
        row12.add(calcium);
        tableList.add(row12);

        List<String> row13 = new ArrayList<>();
        row13.add("硫酸根");
        row13.add("ppm");
        row13.add(sulfate);
        tableList.add(row13);

        List<String> row14 = new ArrayList<>();
        row14.add("硝酸根");
        row14.add("ppm");
        row14.add(nitrate);
        tableList.add(row14);

        List<String> row15 = new ArrayList<>();
        row15.add("磷酸根");
        row15.add("ppm");
        row15.add(phosphate);
        tableList.add(row15);

        List<String> row16 = new ArrayList<>();
        row16.add("碳酸氢根");
        row16.add("ppm");
        row16.add(bicarbonate);
        tableList.add(row16);

        List<String> row17 = new ArrayList<>();
        row17.add("总铁量");
        row17.add("ppm");
        row17.add(ironTotal);
        tableList.add(row17);

        List<String> row18 = new ArrayList<>();
        row18.add("三价铁");
        row18.add("ppm");
        row18.add(ferricIon);
        tableList.add(row18);

        List<String> row19 = new ArrayList<>();
        row19.add("二价铁");
        row19.add("ppm");
        row19.add(ferrous);
        tableList.add(row19);

        List<String> row20 = new ArrayList<>();
        row20.add("温度");
        row20.add("℃");
        row20.add(temperature);
        tableList.add(row20);

        List<String> row21 = new ArrayList<>();
        row21.add("pH");
        row21.add("");
        row21.add(ph);
        tableList.add(row21);

        List<String> row22 = new ArrayList<>();
        row22.add("进水流量");
        row22.add("m3/h");
        row22.add(influentFlow);
        tableList.add(row22);

        List<String> row23 = new ArrayList<>();
        row23.add("回收率");
        row23.add("%");
        row23.add(recoveryRate);
        tableList.add(row23);

        List<String> row24 = new ArrayList<>();
        row24.add("COD");
        row24.add("pmm");
        row24.add(chemicalOxygenDemand);
        tableList.add(row24);

        List<String> row25 = new ArrayList<>();
        row25.add("污染密度指数");
        row25.add("");
        row25.add(siltDensityIndex);
        tableList.add(row25);

        buildTable(document, rowTitle, tableList, 3, 320f);
    }

    /**
     * 添加水质分析表格-重工业-杀菌剂+阻垢剂
     *
     * @param document
     * @throws DocumentException
     */
    public void buildWaterTableHBS(Document document, String aluminum, String silica, String sodium, String magnesium
            , String barium, String kalium, String manganese, String strontium, String fluorine
            , String chlorine, String bromine, String calcium, String sulfate, String nitrate, String phosphate
            , String bicarbonate, String ironTotal, String ferricIon, String ferrous, String temperature
            , String ph, String influentFlow, String recoveryRate, String chemicalOxygenDemand, String siltDensityIndex
            , String totalBacteriaCount) throws DocumentException {
        //构建标题
        List<String> rowTitle = new ArrayList<>();
        rowTitle.add("关键指标");
        rowTitle.add("单位");
        rowTitle.add("数值");
        //构建内容
        //构建内容
        //构建内容
        List<List<String>> tableList = new ArrayList<>();
        List<String> row1 = new ArrayList<>();
        row1.add("铝");
        row1.add("ppm");
        row1.add(aluminum);
        tableList.add(row1);

        List<String> row2 = new ArrayList<>();
        row2.add("硅");
        row2.add("ppm");
        row2.add(silica);
        tableList.add(row2);

        List<String> row3 = new ArrayList<>();
        row3.add("钠");
        row3.add("ppm");
        row3.add(sodium);
        tableList.add(row3);

        List<String> row4 = new ArrayList<>();
        row4.add("镁");
        row4.add("ppm");
        row4.add(magnesium);
        tableList.add(row4);

        List<String> row5 = new ArrayList<>();
        row5.add("钡");
        row5.add("ppm");
        row5.add(barium);
        tableList.add(row5);

        List<String> row6 = new ArrayList<>();
        row6.add("钾");
        row6.add("ppm");
        row6.add(kalium);
        tableList.add(row6);

        List<String> row7 = new ArrayList<>();
        row7.add("锰");
        row7.add("ppm");
        row7.add(manganese);
        tableList.add(row7);

        List<String> row8 = new ArrayList<>();
        row8.add("锶");
        row8.add("ppm");
        row8.add(strontium);
        tableList.add(row8);

        List<String> row9 = new ArrayList<>();
        row9.add("氟");
        row9.add("ppm");
        row9.add(fluorine);
        tableList.add(row9);

        List<String> row10 = new ArrayList<>();
        row10.add("氯");
        row10.add("ppm");
        row10.add(chlorine);
        tableList.add(row10);

        List<String> row11 = new ArrayList<>();
        row11.add("溴");
        row11.add("ppm");
        row11.add(bromine);
        tableList.add(row11);

        List<String> row12 = new ArrayList<>();
        row12.add("钙");
        row12.add("ppm");
        row12.add(calcium);
        tableList.add(row12);

        List<String> row13 = new ArrayList<>();
        row13.add("硫酸根");
        row13.add("ppm");
        row13.add(sulfate);
        tableList.add(row13);

        List<String> row14 = new ArrayList<>();
        row14.add("硝酸根");
        row14.add("ppm");
        row14.add(nitrate);
        tableList.add(row14);

        List<String> row15 = new ArrayList<>();
        row15.add("磷酸根");
        row15.add("ppm");
        row15.add(phosphate);
        tableList.add(row15);

        List<String> row16 = new ArrayList<>();
        row16.add("碳酸氢根");
        row16.add("ppm");
        row16.add(bicarbonate);
        tableList.add(row16);

        List<String> row17 = new ArrayList<>();
        row17.add("总铁量");
        row17.add("ppm");
        row17.add(ironTotal);
        tableList.add(row17);

        List<String> row18 = new ArrayList<>();
        row18.add("三价铁");
        row18.add("ppm");
        row18.add(ferricIon);
        tableList.add(row18);

        List<String> row19 = new ArrayList<>();
        row19.add("二价铁");
        row19.add("ppm");
        row19.add(ferrous);
        tableList.add(row19);

        List<String> row20 = new ArrayList<>();
        row20.add("温度");
        row20.add("℃");
        row20.add(temperature);
        tableList.add(row20);

        List<String> row21 = new ArrayList<>();
        row21.add("pH");
        row21.add("");
        row21.add(ph);
        tableList.add(row21);

        List<String> row22 = new ArrayList<>();
        row22.add("进水流量");
        row22.add("m3/h");
        row22.add(influentFlow);
        tableList.add(row22);

        List<String> row23 = new ArrayList<>();
        row23.add("回收率");
        row23.add("%");
        row23.add(recoveryRate);
        tableList.add(row23);

        List<String> row24 = new ArrayList<>();
        row24.add("COD");
        row24.add("pmm");
        row24.add(chemicalOxygenDemand);
        tableList.add(row24);

        List<String> row25 = new ArrayList<>();
        row25.add("污染密度指数");
        row25.add("");
        row25.add(siltDensityIndex);
        tableList.add(row25);

        List<String> row26 = new ArrayList<>();
        row26.add("细菌总数");
        row26.add("CFU/ml");
        row26.add(totalBacteriaCount);
        tableList.add(row26);

        buildTable(document, rowTitle, tableList, 3, 320f);
    }

    /**
     * 添加系统性能表格
     *
     * @param document
     * @throws DocumentException
     */
    public void buildFunctionTable(Document document, String cfrfValue, String cipValue, String ocfValue) throws DocumentException {
        //构建标题
        List<String> rowTitle = new ArrayList<>();
        rowTitle.add("关键指标");
        rowTitle.add("单位");
        rowTitle.add("数值");
        //构建内容
        List<List<String>> tableList = new ArrayList<>();
        List<String> row1 = new ArrayList<>();
        row1.add("保安过滤器滤芯更换周期");
        row1.add("天");
        row1.add(cfrfValue);
        tableList.add(row1);

        List<String> row2 = new ArrayList<>();
        row2.add("反渗透系统在线清洗周期");
        row2.add("天");
        row2.add(cipValue);
        tableList.add(row2);

        List<String> row3 = new ArrayList<>();
        row3.add("反渗透系统离线清洗周期");
        row3.add("天");
        row3.add(ocfValue);
        tableList.add(row3);

        buildTable(document, rowTitle, tableList, 3, 320f);
    }

    /**
     * 添加解决方案表格
     *
     * @param document
     * @throws DocumentException
     */
    public void buildProgramTable(Document document, List<List<String>> tableList) throws DocumentException {
        //构建标题
        List<String> rowTitle = new ArrayList<>();
        rowTitle.add("产品名称");
        rowTitle.add("加药量 ppm");
        rowTitle.add("加药方式");
        rowTitle.add("加药位置");
//        //构建尾行
//        List<String> foot = tableList.get(tableList.size()-1);
//        tableList.remove(tableList.size()-1);
        buildTable(document, rowTitle, tableList, 4, 380f);
    }

    public void buildDesilicationTable(Document document, String heroSuggestions, String n1998SISuggestions
            , String n1998SIProductName, String n1998SIProductValue
            , String n1998SIAddingPlace, String n1998SIAddingType, String sludgeGenerationName
            , String sludgeGenerationValue, String sludgeGenerationUse, String sludgeGenerationExplain
            , String extraCausticNeededName, String extraCausticNeededValue, String extraCausticNeededUse
            , String extraCausticNeededExplain) throws DocumentException {
        if (StringUtils.isNotBlank(heroSuggestions)) {
            buildParagraph(document, heroSuggestions, fontBlack);
        } else {
            if (StringUtils.isNotBlank(n1998SISuggestions)) {
                buildParagraph(document, n1998SISuggestions, fontBlack);
            }

            List<String> rowTitle1 = new ArrayList<>();
            rowTitle1.add("推荐药剂");
            rowTitle1.add("推荐剂量 ppm");
            rowTitle1.add("投加位置");
            rowTitle1.add("投加方式");
            List<List<String>> tableList1 = new ArrayList<>();
            List<String> valueList1 = new ArrayList<>();
            valueList1.add(n1998SIProductName);
            valueList1.add(n1998SIProductValue);
            valueList1.add(n1998SIAddingPlace);
            valueList1.add(n1998SIAddingType);
            tableList1.add(valueList1);
            buildTable(document, rowTitle1, tableList1, 4, 380f);

            List<String> rowTitle2 = new ArrayList<>();
            rowTitle2.add("其他");
            rowTitle2.add("预估值,kg/m³");
            rowTitle2.add("用途");
            rowTitle2.add("说明");
            List<List<String>> tableList2 = new ArrayList<>();
            List<String> valueList2 = new ArrayList<>();
            valueList2.add(sludgeGenerationName);
            valueList2.add(sludgeGenerationValue);
            valueList2.add(sludgeGenerationUse);
            valueList2.add(sludgeGenerationExplain);
            tableList2.add(valueList2);
            List<String> valueList3 = new ArrayList<>();
            valueList3.add(extraCausticNeededName);
            valueList3.add(extraCausticNeededValue);
            valueList3.add(extraCausticNeededUse);
            valueList3.add(extraCausticNeededExplain);
            tableList2.add(valueList3);
            buildTable(document, rowTitle2, tableList2, 4, 380f);
        }

    }

    public void buildN3108Table(Document document, String n3108ProductName, String n3108Value, String n3108AddingPlace
            , String n3108AddingType) throws DocumentException {
        //构建标题
        List<String> rowTitle = new ArrayList<>();
        rowTitle.add("推荐药剂");
        rowTitle.add("推荐剂量 ppm");
        rowTitle.add("投加位置");
        rowTitle.add("投加方式");
        List<List<String>> tableList = new ArrayList<>();
        List<String> valueList = new ArrayList<>();
        valueList.add(n3108ProductName);
        valueList.add(n3108Value);
        valueList.add(n3108AddingPlace);
        valueList.add(n3108AddingType);
        tableList.add(valueList);
        buildTable(document, rowTitle, tableList, 4, 380f);
    }

    public void buildScaleInhibitorTable(Document document, String feedCalciteSrValue, String feedLSIValue
            , String concentrationfactorValue, String pHValue, String calciteSRValue, String concentrateLSIValue
            , String caValue, String siO2Value, String mgValue
            , List<List<String>> tableList) throws DocumentException {
        List<String> rowTitle1 = new ArrayList<>();
        rowTitle1.add("进水碳酸钙过饱和度");
        rowTitle1.add("进水郎利尔过饱和指数");
        rowTitle1.add("浓缩倍数");
        rowTitle1.add("pH");
        rowTitle1.add("碳酸钙过饱和度");
        rowTitle1.add("浓水郎利尔过饱和指数");
        rowTitle1.add("钙");
        rowTitle1.add("二氧化硅");
        rowTitle1.add("镁");
        List<List<String>> tableList1 = new ArrayList<>();
        List<String> valueList1 = new ArrayList<>();
        valueList1.add(feedCalciteSrValue);
        valueList1.add(feedLSIValue);
        valueList1.add(concentrationfactorValue);
        valueList1.add(pHValue);
        valueList1.add(calciteSRValue);
        valueList1.add(concentrateLSIValue);
        valueList1.add(caValue);
        valueList1.add(siO2Value);
        valueList1.add(mgValue);
        tableList1.add(valueList1);
        buildTableForScaleInhibitor(document, rowTitle1, tableList1, 9, 500f, false);


        List<String> rowTitle2 = new ArrayList<>();
        rowTitle2.add("选取产品");
        rowTitle2.add("浓水加药量");
        rowTitle2.add("进水加药量");
        rowTitle2.add("碳酸钙");
        rowTitle2.add("硫酸钙");
        rowTitle2.add("磷酸钙");
        rowTitle2.add("二氧化硅");
        rowTitle2.add("铝");
        rowTitle2.add("铁");
        rowTitle2.add("锰");
        buildTableForScaleInhibitor(document, rowTitle2, tableList, 10, 500f, true);
    }


    /**
     * 添加系统性能预测表格
     *
     * @param document
     * @throws DocumentException
     */
    public void buildForecastTable(Document document, String cfrfValue, String cfrfValueNew, String cipValue
            , String cipValueNew, String ocfValue, String ocfValueNew) throws DocumentException {
        //构建标题
        List<String> rowTitle = new ArrayList<>();
        rowTitle.add("");
        rowTitle.add("");
        rowTitle.add("使用原方案");
        rowTitle.add("使用新方案");
        //构建内容
        List<List<String>> tableList = new ArrayList<>();
        List<String> row1 = new ArrayList<>();
        row1.add("保安过滤器滤芯更换周期");
        row1.add("(天)");
        row1.add(cfrfValue);
        row1.add(cfrfValueNew);
        tableList.add(row1);

        List<String> row2 = new ArrayList<>();
        row2.add("反渗透在线清洗周期");
        row2.add("(天)");
        row2.add(cipValue);
        row2.add(cipValueNew);
        tableList.add(row2);

        List<String> row3 = new ArrayList<>();
        row3.add("反渗透离线清洗周期");
        row3.add("(天)");
        row3.add(ocfValue);
        row3.add(ocfValueNew);
        tableList.add(row3);

        buildTable(document, rowTitle, tableList, 4, 380f);
    }

}
