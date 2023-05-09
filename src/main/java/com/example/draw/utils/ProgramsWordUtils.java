package com.example.draw.utils;

import com.itextpdf.text.DocumentException;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Internal;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.net.URISyntaxException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

/**
 * @Description
 * @Author Roger
 * @Date 2020/9/21
 */
public class ProgramsWordUtils extends WordUtils {

    @Override
    public void generateWord(XWPFDocument document, Map<String, Object> params, Map<String, List<List<String>>> tableMap) throws Exception {
        // 页眉页脚
        createHeaderAndFooter(document);

        // 报告封面
        buildHomePage(document);

        // 标题
        buildTitle(document, " 委托信息 / 报告概要");

        // 候选人、委托日期等信息的表格
        buildMainTable(document);

//        buildTitleTable(document, "方案建议书");

        for (int i = 0; i <= 20; i++) {
            buildMainTable(document
                    , "111"
                    , "222"
                    , "333"
                    , "444");
        }


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
//            pageBreak(document);
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
////            todo
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
//                    buildParagraph(document, "除硅预处理后参数内容", 9);
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
//                    buildParagraph(document, "N3108预处理后参数内容", 9);
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
//                    buildParagraph(document, "除硅预处理后参数内容", 9);
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
//                    buildParagraph(document, "N3108预处理后参数内容", 9);
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
//                    buildParagraph(document, "除硅预处理后参数内容", 9);
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
//                    buildParagraph(document, "N3108预处理后参数内容", 9);
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
//            pageBreak(document);
//
//            buildTitle(document, "解决方案", "基于对回用系统水质和运行性能的了解，对该系统进行诊断，推荐以下方案：");
//
//            if (params.containsKey("${heavy_desilication_flag}")) {
//                buildParagraph(document, "除硅预处理推荐方案内容", 9);
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
//                buildParagraph(document, "N3108预处理推荐方案内容", 9);
//
//                buildN3108Table(document
//                        , params.get("${n3108ProductName}").toString()
//                        , params.get("${n3108Value}").toString()
//                        , params.get("${n3108AddingPlace}").toString()
//                        , params.get("${n3108AddingType}").toString());
//            }
//            if (params.containsKey("${heavy_product_flag}")) {
//                buildParagraph(document, "杀菌剂推荐方案内容", 9);
//
//                buildProgramTable(document, tableMap.get("heavy_product_table"));
//            }
//            if (params.containsKey("${heavy_scale_inhibitor_flag}")) {
//                buildParagraph(document, "阻垢剂推荐方案内容", 9);
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
//                buildParagraph(document, params.get("${heavy_msg_flag}").toString(), "");
//            }
//
//        }
    }

    /**
     * 候选人、委托日期等信息的表格
     */
    public void buildMainTable(XWPFDocument document) throws InvalidFormatException, IOException, URISyntaxException {
        blankParagraph(document);

        XWPFTable table = document.createTable(2, 5);

        //设置表格宽度
        CTTblPr tablePr = table.getCTTbl().addNewTblPr();
        //表格宽度
        CTTblWidth tableWidth = tablePr.addNewTblW();
        tableWidth.setW(BigInteger.valueOf(8310));
        //设置表格宽度为非自动
        tableWidth.setType(STTblWidth.DXA);
        //设置边框
        displayBorder(table);


        XWPFTableRow xwpfTableRow1 = table.getRow(0);
        XWPFTableCell cell11 = xwpfTableRow1.getCell(0);
        builderCell(cell11, "姓  名：", colorGary, ParagraphAlignment.LEFT, 1000, 8, null);
        XWPFTableCell cell12 = xwpfTableRow1.getCell(1);
        builderCell(cell12, "张三", colorGary, ParagraphAlignment.LEFT, 3000, 8, null);
        XWPFTableCell cell13 = xwpfTableRow1.getCell(2);
        builderCell(cell13, "", colorWrite, ParagraphAlignment.LEFT, 310, 8, null);
        XWPFTableCell cell14 = xwpfTableRow1.getCell(3);
        builderCell(cell14, "委托日期：", colorPink, ParagraphAlignment.LEFT, 1000, 8, null);
        XWPFTableCell cell15 = xwpfTableRow1.getCell(4);
        builderCell(cell15, "2023年1月1日", colorPink, ParagraphAlignment.LEFT, 3000, 8, null);


        XWPFTableRow xwpfTableRow2 = table.getRow(1);
        XWPFTableCell cell21 = xwpfTableRow2.getCell(0);
        builderCell(cell21, "交付类型：", colorGary, ParagraphAlignment.LEFT, 1000, 8, null);
        XWPFTableCell cell22 = xwpfTableRow2.getCell(1);
        builderCell(cell22, "终版报告", colorGary, ParagraphAlignment.LEFT, 3000, 8, null);
        XWPFTableCell cell23 = xwpfTableRow2.getCell(2);
        builderCell(cell23, "", colorWrite, ParagraphAlignment.LEFT, 310, 8, null);
        XWPFTableCell cell24 = xwpfTableRow2.getCell(3);
        builderCell(cell24, "交付日期：", colorPink, ParagraphAlignment.LEFT, 1000, 8, null);
        XWPFTableCell cell25 = xwpfTableRow2.getCell(4);
        builderCell(cell25, "2023年1月3日", colorPink, ParagraphAlignment.LEFT, 3000, 8, null);

//        XWPFTableRow xwpfTableRow3 = table.getRow(2);

    }

    public void builderCell(XWPFTableCell cell, String content, String backgroundColor, ParagraphAlignment paragraphAlign, Integer width, Integer fontSize, String filePath) throws IOException, InvalidFormatException, URISyntaxException {
        XWPFParagraph paragraph = cell.getParagraphs().get(0);
        if(StringUtils.isNotBlank(filePath)) {
            //图片
            XWPFRun pictureRun = paragraph.createRun();
            FileInputStream is = null;
            try {
                // filePath = /static/image/image1.png
                is = new FileInputStream(new File(this.getClass().getResource(filePath).toURI()));
                pictureRun.addPicture(is, Document.PICTURE_TYPE_JPEG, "c1.png", Units.toEMU(120), Units.toEMU(30));
            } catch (IOException | InvalidFormatException | URISyntaxException e) {
                throw e;
            } finally {
                if (is != null) {
                    is.close();
                }
            }
        } else {
            CTTc cttc = cell.getCTTc();
            CTTcPr ctPr = cttc.addNewTcPr();
            /** 背景色 */
            cell.setColor(backgroundColor);
            /** 水平居中 */
            cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
            /** 竖直居中 */
            ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
            cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
            /**单元格宽度*/
            CTTblWidth ctTblWidthCell = ctPr.addNewTcW();
            ctTblWidthCell.setType(STTblWidth.DXA);
            ctTblWidthCell.setW(BigInteger.valueOf(width));

            XWPFParagraph paragraph1 = cell.getParagraphs().get(0);
            buildParagraph(paragraph1, paragraphAlign, content, fontSize, null, colorBlack);
        }

        cell.setParagraph(paragraph);
    }

    /**
     * 报告首页
     */

    public void buildHomePage(XWPFDocument document) throws IOException, InvalidFormatException, URISyntaxException {
        /**
         * 雇前背景调查报告
         */
        List<List<String>> content = new ArrayList<>();
        content.add(Arrays.asList("雇 前 背 景 调 查 报 告"));
        buildTableSpecial(document, null, content,1, new Long[]{8310L}, new Integer[]{600}, 8130L, true, null, null, colorBlue, 28, true);

        // 空一行
        blankParagraph(document);

        /**
         * 委托日期
         */
        List<List<String>> content2 = new ArrayList<>();
        content2.add(Arrays.asList("委托日期：2023-01-01"));
        // 背景色 浅蓝色
        buildTableSpecial(document, null, content2,1, new Long[]{4000L}, new Integer[]{400}, 4000L, true, null, colorBlue2, colorWrite, 14, true);

        // 空一行
        blankParagraph(document);
        /**
         * 公司名称、委托方名称、报告编号
         */
        List<String> strings1 = new ArrayList<>();
        strings1.add("北京字节跳动网络技术有限公司");
        List<String> strings2 = new ArrayList<>();
        strings2.add("张三");
        List<String> strings3 = new ArrayList<>();
        strings3.add("报告编号：BJDC20230101000001");
        List<List<String>> content3 = new ArrayList();
        content3.add(strings1);
        content3.add(strings2);
        content3.add(strings3);
        buildTableSpecial(document, null, content3,1, new Long[]{8310L}, new Integer[]{400, 400, 400}, 8310L, true, null, null, colorBlack, 14, false);

        /**
         * 内部保密文件
         */

        List<List<String>> content4 = new ArrayList<>();
        content4.add(Arrays.asList("<内部保密文件>"));
        buildTableSpecial(document, null, content4,1, new Long[]{8310L}, new Integer[]{600}, 8310L, true, null, null, colorRed, 12, false);

        /**
         * L4级机密、禁止分享、限期删除
         */
        List<String> stringList = new ArrayList<>();
        stringList.add("");
        stringList.add("L4级机密");
        stringList.add("");
        stringList.add("禁止分享");
        stringList.add("");
        stringList.add("限期删除");
        stringList.add("");
        List<List<String>> content5 = new ArrayList();
        content5.add(stringList);
//        buildTable(document, null, null, content, 3, new Long[]{800L, 800L, 800L}, 2400L);
        buildTableSpecial2(document, null, content5,7, new Long[]{1655L, 1400L, 400L, 1400L, 400L, 1400L, 1655L}, new Integer[]{200}, 8310L, true, null, colorPink, colorRed, 10, true);

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
    public void buildTableSpecial(XWPFDocument document
            , List<String> title
            , List<List<String>> content
            , Integer numColumn
            , Long[] tableWidths
            , Integer[] rowsHeight
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
        System.out.println("rowNum=" + rowNum + ", columnNum=" + columnNum);
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
            xwpfTableRowTitle.setHeight(rowsHeight[0]);
            for (int i = 0; i < title.size() && i < columnNum; i++) {
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                String msg = title.get(i);
                //单元格对象
                XWPFTableCell xwpfTableCell = xwpfTableRowTitle.getCell(i);
                //添加文本，9号/黑体/黑色/居中
                buildCellSpecial(xwpfTableCell, msg, 9, true, colorBlack, ParagraphAlignment.CENTER, cellWidth, null, titleBackground);
                // todo 添加边框
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
                xwpfTableRow.setHeight(rowsHeight[j + 1]);
            } else {
                // 行对象
                xwpfTableRow = xwpfTable.getRow(j);
                //行高
                xwpfTableRow.setHeight(rowsHeight[j]);
            }
            List<String> row = content.get(j);
            for (int i = 0; i < row.size() && i < columnNum; i++) {
                //单元格对象
                XWPFTableCell xwpfTableCell = xwpfTableRow.getCell(i);
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                //黑字 居中
//                String wordColor = colorBlack;
                ParagraphAlignment align = ParagraphAlignment.CENTER;
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
            , List<List<String>> content
            , Integer numColumn
            , Long[] tableWidths
            , Integer[] rowsHeight
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
        System.out.println("rowNum=" + rowNum + ", columnNum=" + columnNum);
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
            xwpfTableRowTitle.setHeight(rowsHeight[0]);
            for (int i = 0; i < title.size() && i < columnNum; i++) {
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                String msg = title.get(i);
                //单元格对象
                XWPFTableCell xwpfTableCell = xwpfTableRowTitle.getCell(i);
                //添加文本，9号/黑体/黑色/居中
                buildCellSpecial(xwpfTableCell, msg, 9, true, colorBlack, ParagraphAlignment.CENTER, cellWidth, null, titleBackground);
                // todo 添加边框
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
                xwpfTableRow.setHeight(rowsHeight[j + 1]);
            } else {
                // 行对象
                xwpfTableRow = xwpfTable.getRow(j);
                //行高
                xwpfTableRow.setHeight(rowsHeight[j]);
            }
            List<String> row = content.get(j);
            for (int i = 0; i < row.size() && i < columnNum; i++) {
                //单元格对象
                XWPFTableCell xwpfTableCell = xwpfTableRow.getCell(i);
                if (tableWidths != null) {
                    cellWidth = tableWidths[i];
                }
                //黑字 居中
//                String wordColor = colorBlack;
                ParagraphAlignment align = ParagraphAlignment.CENTER;
                if (i % 2 == 1) {
                    buildCellSpecial(xwpfTableCell, row.get(i), fontSize, bold, wordColor, align, cellWidth, null, tableBackground);
                } else {
                    buildCellSpecial(xwpfTableCell, row.get(i), fontSize, bold, wordColor, align, cellWidth, null, null);
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
            // todo
            String imageName = value.toString().substring(value.toString().lastIndexOf("/") + 1);
            XWPFRun pictureRun = cell.getParagraphs().get(0).createRun();
            FileInputStream is = null;
            try {
                is = new FileInputStream(new File(this.getClass().getResource(((StringBuffer) value).toString()).toURI()));
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
        xwpfRun.setFontFamily("黑体");
        if (bold != null) {
            xwpfRun.setBold(bold);
        }
        if (color != null) {
            xwpfRun.setColor(color);
        }
    }
//    public void buildTable2(XWPFDocument document, List<String> title, List<String> foot, List<List<String>> content, Integer numColumn
//            , Long[] tableWidths, Long width) throws InvalidFormatException, IOException, URISyntaxException {
//        int rowNum = content.size();
//        int columnNum = numColumn == null ? tableWidths.length : numColumn;
//        if (CollectionUtils.isNotEmpty(title)) {
//            rowNum++;
//        }
//        if (CollectionUtils.isNotEmpty(foot)) {
//            rowNum++;
//        }
//        System.out.println("rowNum=" + rowNum + ", columnNum=" + columnNum);
//        XWPFTable xwpfTable = document.createTable(rowNum, columnNum);
//        //表格居中显示
//        CTTblPr ctTblPr = xwpfTable.getCTTbl().addNewTblPr();
//        ctTblPr.addNewJc().setVal(STJc.CENTER);
//        //设置表格宽度
//        CTTblWidth ctTblWidth = ctTblPr.addNewTblW();
//        ctTblWidth.setW(BigInteger.valueOf(width));
//        //设置表格宽度为非自动
//        ctTblWidth.setType(STTblWidth.DXA);
//        Long cellWidth = null;
//        // 设置无边框
//        displayBorder(xwpfTable);
//
//        //行对象
//        XWPFTableRow xwpfTableRowTitle = xwpfTable.getRow(0);
//        //行高
//        xwpfTableRowTitle.setHeight(350);
//        //创建标题
//        if (CollectionUtils.isNotEmpty(title)) {
//            for (int i = 0; i < title.size() && i < columnNum; i++) {
//                if (tableWidths != null) {
//                    cellWidth = tableWidths[i];
//                }
//                String msg = title.get(i);
//                //单元格对象
//                XWPFTableCell xwpfTableCell = xwpfTableRowTitle.getCell(i);
//                //添加文本，9号/黑体/白色/居中
//                buildCell(xwpfTableCell, msg, 9, true, colorWrite, ParagraphAlignment.CENTER, cellWidth, null);
//                //添加背景色，蓝色
//                if (StringUtils.isNotEmpty(msg)) {
//                    xwpfTableCell.setColor(colorBlue);
//                }
//                //添加边框
//                addBottomBorder(xwpfTableCell, 1, colorBlue, true);
//            }
//        }
//
//        //创建内容
//        for (int j = 0; j < content.size(); j++) {
//            //行对象 有标题时和无标题时取值位置不同
//            XWPFTableRow xwpfTableRow;
//            if (CollectionUtils.isNotEmpty(title)) {
//                xwpfTableRow = xwpfTable.getRow(j + 1);
//            } else {
//                xwpfTableRow = xwpfTable.getRow(j);
//            }
//            //行高
//            xwpfTableRow.setHeight(500);
//
//            List<String> row = content.get(j);
//            for (int i = 0; i < row.size() && i < columnNum; i++) {
//                //单元格对象
//                XWPFTableCell xwpfTableCell = xwpfTableRow.getCell(i);
//                if (tableWidths != null) {
//                    cellWidth = tableWidths[i];
//                }
//                //黑字 居中
//                String wordColor = colorBlack;
//                ParagraphAlignment align = ParagraphAlignment.CENTER;
//                //添加文本，9号/黑体/左对齐
//                buildCell(xwpfTableCell, row.get(i), 14, false, wordColor, align, cellWidth, null);
//            }
//        }
//    }


    /**
     * 标题
     */
    public void buildTitle(XWPFDocument document, String titleName) throws IOException, InvalidFormatException, URISyntaxException {
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
//        CTTc cttc2 = xwpfTableCell1.getCTTc();
//        CTTcPr ctPr2 = cttc2.addNewTcPr();
//        // 设置宽度
//        CTTblWidth ctTblWidthCell2 = ctPr2.addNewTcW();
//        ctTblWidthCell2.setType(STTblWidth.DXA);
//        ctTblWidthCell2.setW(BigInteger.valueOf(4000));
//        XWPFParagraph paragraph2 = xwpfTableCell2.getParagraphs().get(0);
//        XWPFRun run2 = paragraph2.createRun();
//        run2.setText("    ");
//        run2.setText(titleName);
//        run2.setBold(true);
//        run2.setFontFamily("黑体");
//        run2.setColor(colorBlack);
        buildCell(xwpfTableCell2, titleName, 12, true, colorBlack, ParagraphAlignment.LEFT, 8270L, true);
    }

    /**
     * 页眉页脚
     */
    public static void createHeaderAndFooter( XWPFDocument document) throws Exception {
        // 页眉
        // Appends and returns a new empty "sectPr" element
        CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();

        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(document, sectPr);

        XWPFHeader header = headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);
        XWPFParagraph paragraph = header.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.RIGHT);
        // paragraph.setBorderBottom(Borders.THICK);
        XWPFRun run = paragraph.createRun();

        String imgFile = "/Users/kanmeijie/Workspace/draw/src/main/resources/static/image/image2.jpg";
        File file = new File( imgFile );
        InputStream is = new FileInputStream( file );
        XWPFPicture picture = run.addPicture( is, XWPFDocument.PICTURE_TYPE_JPEG, imgFile, Units.toEMU( 80 ), Units.toEMU( 45 ) );
        String blipID = "";
        for( XWPFPictureData picturedata : header.getAllPackagePictures() ) { // 这段必须有，不然打开的logo图片不显示
            blipID = header.getRelationId( picturedata );
            picture.getCTPicture().getBlipFill().getBlip().setEmbed( blipID );
        }
        // 添加tab
        // run.addTab();
        is.close();

        // 页脚
        XWPFFooter footer = headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);
        XWPFParagraph paragraph1 = footer.createParagraph();
        paragraph1.setAlignment(ParagraphAlignment.RIGHT);
        // paragraph1.setBorderBottom(Borders.THICK);

        paragraph1.getCTP().addNewFldSimple().setInstr("PAGE \\* MERGEFORMAT");
        XWPFRun runFooter1 = paragraph1.createRun();
        runFooter1.setText(" / ");
        paragraph1.getCTP().addNewFldSimple().setInstr("NUMPAGES \\* MERGEFORMAT");
    }

    public void buildMainTable(XWPFDocument document, String accountName, String recycleSystemName, String proposalName
            , String createDate) throws InvalidFormatException, IOException, URISyntaxException {
        blankParagraph(document);

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
//    public void buildBackendParaL(XWPFDocument document) throws DocumentException {
//        buildParagraph(document, "反渗透系统总体运行效率不高，主要表现在以下几个方面：", 9);
//
//        buildParagraph(document, "●   保安过滤器(超滤和反渗透膜前)滤芯更换频繁", 9);
//
//        buildParagraph(document, "●   反渗透膜清洗频率高", 9);
//
//        buildParagraph(document, "●   反渗透膜产水流量低，回收率低于系统设计回收率", 9);
//
//        buildParagraph(document, "●   反渗透膜压差偏高", 9);
//
//    }

    /**
     * 添加背景描述
     *
     * @param document
     */
//    public void buildBackendParaH(XWPFDocument document) throws DocumentException {
//        buildParagraph(document, "水质中的污染物成分若不进行适当处理和控制，会造成反渗透系统总体运行效率不高，主要表现在以下几个方面：", 9);
//
//        buildParagraph(document, "●   保安过滤器滤芯更换频繁，增加人工劳动量和系统停机频次", 9);
//
//        buildParagraph(document, "●   膜系统产水流量低，回收率低于系统设计回收率，系统效率低下", 9);
//
//        buildParagraph(document, "●   膜系统压差偏高，能耗增加，同时加大了膜性能损坏的风险", 9);
//
//        buildParagraph(document, "●   膜系统清洗频率高，膜寿命降低", 9);
//    }

    /**
     * 添加水质分析表格-(轻工业表单)
     *
     * @param document
     * @throws DocumentException
     */
//    public void buildWaterTableL(XWPFDocument document, String aluminum, String iron, String silica, String copper
//            , String totalBacteriaCount, String ph, String conductivity, String temperature, String chemicalOxygenD
//            , String totalOrganicC, String turbidity) throws DocumentException, InvalidFormatException, IOException, URISyntaxException {
//        //构建标题
//        List<String> rowTitle = new ArrayList<>();
//        rowTitle.add("关键指标");
//        rowTitle.add("单位");
//        rowTitle.add("数值");
//        //构建内容
//        List<List<String>> tableList = new ArrayList<>();
//        List<String> row1 = new ArrayList<>();
//        row1.add("铝");
//        row1.add("ppm");
//        row1.add(aluminum);
//        tableList.add(row1);
//
//        List<String> row2 = new ArrayList<>();
//        row2.add("铁");
//        row2.add("ppm");
//        row2.add(iron);
//        tableList.add(row2);
//
//        List<String> row3 = new ArrayList<>();
//        row3.add("硅");
//        row3.add("ppm");
//        row3.add(silica);
//        tableList.add(row3);
//
//        List<String> row4 = new ArrayList<>();
//        row4.add("铜");
//        row4.add("ppm");
//        row4.add(copper);
//        tableList.add(row4);
//
//        List<String> row5 = new ArrayList<>();
//        row5.add("细菌总数");
//        row5.add("CFU/ml");
//        row5.add(totalBacteriaCount);
//        tableList.add(row5);
//
//        List<String> row6 = new ArrayList<>();
//        row6.add("pH");
//        row6.add("");
//        row6.add(ph);
//        tableList.add(row6);
//
//        List<String> row7 = new ArrayList<>();
//        row7.add("电导率");
//        row7.add("μs/cm");
//        row7.add(conductivity);
//        tableList.add(row7);
//
//        List<String> row8 = new ArrayList<>();
//        row8.add("温度");
//        row8.add("℃");
//        row8.add(temperature);
//        tableList.add(row8);
//
//        List<String> row9 = new ArrayList<>();
//        row9.add("总有机碳");
//        row9.add("ppm");
//        row9.add(chemicalOxygenD);
//        tableList.add(row9);
//
//        List<String> row10 = new ArrayList<>();
//        row10.add("化学需氧量");
//        row10.add("ppm");
//        row10.add(totalOrganicC);
//        tableList.add(row10);
//
//        List<String> row11 = new ArrayList<>();
//        row11.add("浊度");
//        row11.add("NTU");
//        row11.add(turbidity);
//        tableList.add(row11);
//
//        buildTable(document, rowTitle, tableList, 3, 8310L);
//    }

    /**
     * 添加水质分析表格-重工业-杀菌剂
     *
     * @param document
     * @throws DocumentException
     */
//    public void buildWaterTableHB(XWPFDocument document, String aluminum, String ironTotal, String silica, String magnesium
//            , String manganese, String calcium, String totalBacteriaCount, String ph, String recoveryRate
//            , String siltDensityIndex, String chemicalOxygenDemand) throws DocumentException, InvalidFormatException, IOException, URISyntaxException {
//        //构建标题
//        List<String> rowTitle = new ArrayList<>();
//        rowTitle.add("关键指标");
//        rowTitle.add("单位");
//        rowTitle.add("数值");
//        //构建内容
//        List<List<String>> tableList = new ArrayList<>();
//        List<String> row1 = new ArrayList<>();
//        row1.add("铝");
//        row1.add("ppm");
//        row1.add(aluminum);
//        tableList.add(row1);
//
//        List<String> row2 = new ArrayList<>();
//        row2.add("铁");
//        row2.add("ppm");
//        row2.add(ironTotal);
//        tableList.add(row2);
//
//        List<String> row3 = new ArrayList<>();
//        row3.add("硅");
//        row3.add("ppm");
//        row3.add(silica);
//        tableList.add(row3);
//
//        List<String> row4 = new ArrayList<>();
//        row4.add("镁");
//        row4.add("ppm");
//        row4.add(magnesium);
//        tableList.add(row4);
//
//        List<String> row5 = new ArrayList<>();
//        row5.add("锰");
//        row5.add("ppm");
//        row5.add(manganese);
//        tableList.add(row5);
//
//        List<String> row6 = new ArrayList<>();
//        row6.add("钙");
//        row6.add("ppm");
//        row6.add(calcium);
//        tableList.add(row6);
//
//        List<String> row7 = new ArrayList<>();
//        row7.add("细菌总数");
//        row7.add("CFU/ml");
//        row7.add(totalBacteriaCount);
//        tableList.add(row7);
//
//        List<String> row8 = new ArrayList<>();
//        row8.add("pH");
//        row8.add("");
//        row8.add(ph);
//        tableList.add(row8);
//
//        List<String> row9 = new ArrayList<>();
//        row9.add("回收率");
//        row9.add("%");
//        row9.add(recoveryRate);
//        tableList.add(row9);
//
//        List<String> row10 = new ArrayList<>();
//        row10.add("污染密度指数");
//        row10.add("");
//        row10.add(siltDensityIndex);
//        tableList.add(row10);
//
//        List<String> row11 = new ArrayList<>();
//        row11.add("COD");
//        row11.add("ppm");
//        row11.add(chemicalOxygenDemand);
//        tableList.add(row11);
//
//        buildTable(document, rowTitle, tableList, 3, 8310L);
//    }

    /**
     * 添加水质分析表格-重工业-阻垢剂
     *
     * @param document
     * @throws DocumentException
     */
//    public void buildWaterTableHS(XWPFDocument document, String aluminum, String silica, String sodium, String magnesium
//            , String barium, String kalium, String manganese, String strontium, String fluorine
//            , String chlorine, String bromine, String calcium, String sulfate, String nitrate, String phosphate
//            , String bicarbonate, String ironTotal, String ferricIon, String ferrous, String temperature
//            , String ph, String influentFlow, String recoveryRate, String chemicalOxygenDemand, String siltDensityIndex)
//            throws DocumentException, InvalidFormatException, IOException, URISyntaxException {
//        //构建标题
//        List<String> rowTitle = new ArrayList<>();
//        rowTitle.add("关键指标");
//        rowTitle.add("单位");
//        rowTitle.add("数值");
//        //构建内容
//        List<List<String>> tableList = new ArrayList<>();
//        List<String> row1 = new ArrayList<>();
//        row1.add("铝");
//        row1.add("ppm");
//        row1.add(aluminum);
//        tableList.add(row1);
//
//        List<String> row2 = new ArrayList<>();
//        row2.add("硅");
//        row2.add("ppm");
//        row2.add(silica);
//        tableList.add(row2);
//
//        List<String> row3 = new ArrayList<>();
//        row3.add("钠");
//        row3.add("ppm");
//        row3.add(sodium);
//        tableList.add(row3);
//
//        List<String> row4 = new ArrayList<>();
//        row4.add("镁");
//        row4.add("ppm");
//        row4.add(magnesium);
//        tableList.add(row4);
//
//        List<String> row5 = new ArrayList<>();
//        row5.add("钡");
//        row5.add("ppm");
//        row5.add(barium);
//        tableList.add(row5);
//
//        List<String> row6 = new ArrayList<>();
//        row6.add("钾");
//        row6.add("ppm");
//        row6.add(kalium);
//        tableList.add(row6);
//
//        List<String> row7 = new ArrayList<>();
//        row7.add("锰");
//        row7.add("ppm");
//        row7.add(manganese);
//        tableList.add(row7);
//
//        List<String> row8 = new ArrayList<>();
//        row8.add("锶");
//        row8.add("ppm");
//        row8.add(strontium);
//        tableList.add(row8);
//
//        List<String> row9 = new ArrayList<>();
//        row9.add("氟");
//        row9.add("ppm");
//        row9.add(fluorine);
//        tableList.add(row9);
//
//        List<String> row10 = new ArrayList<>();
//        row10.add("氯");
//        row10.add("ppm");
//        row10.add(chlorine);
//        tableList.add(row10);
//
//        List<String> row11 = new ArrayList<>();
//        row11.add("溴");
//        row11.add("ppm");
//        row11.add(bromine);
//        tableList.add(row11);
//
//        List<String> row12 = new ArrayList<>();
//        row12.add("钙");
//        row12.add("ppm");
//        row12.add(calcium);
//        tableList.add(row12);
//
//        List<String> row13 = new ArrayList<>();
//        row13.add("硫酸根");
//        row13.add("ppm");
//        row13.add(sulfate);
//        tableList.add(row13);
//
//        List<String> row14 = new ArrayList<>();
//        row14.add("硝酸根");
//        row14.add("ppm");
//        row14.add(nitrate);
//        tableList.add(row14);
//
//        List<String> row15 = new ArrayList<>();
//        row15.add("磷酸根");
//        row15.add("ppm");
//        row15.add(phosphate);
//        tableList.add(row15);
//
//        List<String> row16 = new ArrayList<>();
//        row16.add("碳酸氢根");
//        row16.add("ppm");
//        row16.add(bicarbonate);
//        tableList.add(row16);
//
//        List<String> row17 = new ArrayList<>();
//        row17.add("总铁量");
//        row17.add("ppm");
//        row17.add(ironTotal);
//        tableList.add(row17);
//
//        List<String> row18 = new ArrayList<>();
//        row18.add("三价铁");
//        row18.add("ppm");
//        row18.add(ferricIon);
//        tableList.add(row18);
//
//        List<String> row19 = new ArrayList<>();
//        row19.add("二价铁");
//        row19.add("ppm");
//        row19.add(ferrous);
//        tableList.add(row19);
//
//        List<String> row20 = new ArrayList<>();
//        row20.add("温度");
//        row20.add("℃");
//        row20.add(temperature);
//        tableList.add(row20);
//
//        List<String> row21 = new ArrayList<>();
//        row21.add("pH");
//        row21.add("");
//        row21.add(ph);
//        tableList.add(row21);
//
//        List<String> row22 = new ArrayList<>();
//        row22.add("进水流量");
//        row22.add("m3/h");
//        row22.add(influentFlow);
//        tableList.add(row22);
//
//        List<String> row23 = new ArrayList<>();
//        row23.add("回收率");
//        row23.add("%");
//        row23.add(recoveryRate);
//        tableList.add(row23);
//
//        List<String> row24 = new ArrayList<>();
//        row24.add("COD");
//        row24.add("pmm");
//        row24.add(chemicalOxygenDemand);
//        tableList.add(row24);
//
//        List<String> row25 = new ArrayList<>();
//        row25.add("污染密度指数");
//        row25.add("");
//        row25.add(siltDensityIndex);
//        tableList.add(row25);
//
//        buildTable(document, rowTitle, tableList, 3, 8310L);
//    }

    /**
     * 添加水质分析表格-重工业-杀菌剂+阻垢剂
     *
     * @param document
     * @throws DocumentException
     */
//    public void buildWaterTableHBS(XWPFDocument document, String aluminum, String silica, String sodium, String magnesium
//            , String barium, String kalium, String manganese, String strontium, String fluorine
//            , String chlorine, String bromine, String calcium, String sulfate, String nitrate, String phosphate
//            , String bicarbonate, String ironTotal, String ferricIon, String ferrous, String temperature
//            , String ph, String influentFlow, String recoveryRate, String chemicalOxygenDemand, String siltDensityIndex
//            , String totalBacteriaCount) throws DocumentException, InvalidFormatException, IOException, URISyntaxException {
//        //构建标题
//        List<String> rowTitle = new ArrayList<>();
//        rowTitle.add("关键指标");
//        rowTitle.add("单位");
//        rowTitle.add("数值");
//        //构建内容
//        //构建内容
//        //构建内容
//        List<List<String>> tableList = new ArrayList<>();
//        List<String> row1 = new ArrayList<>();
//        row1.add("铝");
//        row1.add("ppm");
//        row1.add(aluminum);
//        tableList.add(row1);
//
//        List<String> row2 = new ArrayList<>();
//        row2.add("硅");
//        row2.add("ppm");
//        row2.add(silica);
//        tableList.add(row2);
//
//        List<String> row3 = new ArrayList<>();
//        row3.add("钠");
//        row3.add("ppm");
//        row3.add(sodium);
//        tableList.add(row3);
//
//        List<String> row4 = new ArrayList<>();
//        row4.add("镁");
//        row4.add("ppm");
//        row4.add(magnesium);
//        tableList.add(row4);
//
//        List<String> row5 = new ArrayList<>();
//        row5.add("钡");
//        row5.add("ppm");
//        row5.add(barium);
//        tableList.add(row5);
//
//        List<String> row6 = new ArrayList<>();
//        row6.add("钾");
//        row6.add("ppm");
//        row6.add(kalium);
//        tableList.add(row6);
//
//        List<String> row7 = new ArrayList<>();
//        row7.add("锰");
//        row7.add("ppm");
//        row7.add(manganese);
//        tableList.add(row7);
//
//        List<String> row8 = new ArrayList<>();
//        row8.add("锶");
//        row8.add("ppm");
//        row8.add(strontium);
//        tableList.add(row8);
//
//        List<String> row9 = new ArrayList<>();
//        row9.add("氟");
//        row9.add("ppm");
//        row9.add(fluorine);
//        tableList.add(row9);
//
//        List<String> row10 = new ArrayList<>();
//        row10.add("氯");
//        row10.add("ppm");
//        row10.add(chlorine);
//        tableList.add(row10);
//
//        List<String> row11 = new ArrayList<>();
//        row11.add("溴");
//        row11.add("ppm");
//        row11.add(bromine);
//        tableList.add(row11);
//
//        List<String> row12 = new ArrayList<>();
//        row12.add("钙");
//        row12.add("ppm");
//        row12.add(calcium);
//        tableList.add(row12);
//
//        List<String> row13 = new ArrayList<>();
//        row13.add("硫酸根");
//        row13.add("ppm");
//        row13.add(sulfate);
//        tableList.add(row13);
//
//        List<String> row14 = new ArrayList<>();
//        row14.add("硝酸根");
//        row14.add("ppm");
//        row14.add(nitrate);
//        tableList.add(row14);
//
//        List<String> row15 = new ArrayList<>();
//        row15.add("磷酸根");
//        row15.add("ppm");
//        row15.add(phosphate);
//        tableList.add(row15);
//
//        List<String> row16 = new ArrayList<>();
//        row16.add("碳酸氢根");
//        row16.add("ppm");
//        row16.add(bicarbonate);
//        tableList.add(row16);
//
//        List<String> row17 = new ArrayList<>();
//        row17.add("总铁量");
//        row17.add("ppm");
//        row17.add(ironTotal);
//        tableList.add(row17);
//
//        List<String> row18 = new ArrayList<>();
//        row18.add("三价铁");
//        row18.add("ppm");
//        row18.add(ferricIon);
//        tableList.add(row18);
//
//        List<String> row19 = new ArrayList<>();
//        row19.add("二价铁");
//        row19.add("ppm");
//        row19.add(ferrous);
//        tableList.add(row19);
//
//        List<String> row20 = new ArrayList<>();
//        row20.add("温度");
//        row20.add("℃");
//        row20.add(temperature);
//        tableList.add(row20);
//
//        List<String> row21 = new ArrayList<>();
//        row21.add("pH");
//        row21.add("");
//        row21.add(ph);
//        tableList.add(row21);
//
//        List<String> row22 = new ArrayList<>();
//        row22.add("进水流量");
//        row22.add("m3/h");
//        row22.add(influentFlow);
//        tableList.add(row22);
//
//        List<String> row23 = new ArrayList<>();
//        row23.add("回收率");
//        row23.add("%");
//        row23.add(recoveryRate);
//        tableList.add(row23);
//
//        List<String> row24 = new ArrayList<>();
//        row24.add("COD");
//        row24.add("pmm");
//        row24.add(chemicalOxygenDemand);
//        tableList.add(row24);
//
//        List<String> row25 = new ArrayList<>();
//        row25.add("污染密度指数");
//        row25.add("");
//        row25.add(siltDensityIndex);
//        tableList.add(row25);
//
//        List<String> row26 = new ArrayList<>();
//        row26.add("细菌总数");
//        row26.add("CFU/ml");
//        row26.add(totalBacteriaCount);
//        tableList.add(row26);
//
//        buildTable(document, rowTitle, tableList, 3, 8310L);
//    }

    /**
     * 添加系统性能表格
     *
     * @param document
     * @throws DocumentException
     */
//    public void buildFunctionTable(XWPFDocument document, String cfrfValue, String cipValue, String ocfValue) throws DocumentException, InvalidFormatException, IOException, URISyntaxException {
//        //构建标题
//        List<String> rowTitle = new ArrayList<>();
//        rowTitle.add("关键指标");
//        rowTitle.add("单位");
//        rowTitle.add("数值");
//        //构建内容
//        List<List<String>> tableList = new ArrayList<>();
//        List<String> row1 = new ArrayList<>();
//        row1.add("保安过滤器滤芯更换周期");
//        row1.add("天");
//        row1.add(cfrfValue);
//        tableList.add(row1);
//
//        List<String> row2 = new ArrayList<>();
//        row2.add("反渗透系统在线清洗周期");
//        row2.add("天");
//        row2.add(cipValue);
//        tableList.add(row2);
//
//        List<String> row3 = new ArrayList<>();
//        row3.add("反渗透系统离线清洗周期");
//        row3.add("天");
//        row3.add(ocfValue);
//        tableList.add(row3);
//
//        buildTable(document, rowTitle, tableList, 3, 8310L);
//    }

    /**
     * 添加解决方案表格
     *
     * @param document
     */
//    public void buildProgramTable(XWPFDocument document, List<List<String>> tableList) throws DocumentException, InvalidFormatException, IOException, URISyntaxException {
//        //构建标题
//        List<String> rowTitle = new ArrayList<>();
//        rowTitle.add("产品名称");
//        rowTitle.add("加药量 ppm");
//        rowTitle.add("加药方式");
//        rowTitle.add("加药位置");
//
//        buildTable(document, rowTitle, tableList, 4, 8310L);
//    }

//    public void buildDesilicationTable(XWPFDocument document, String heroSuggestions, String n1998SISuggestions
//            , String n1998SIProductName, String n1998SIProductValue
//            , String n1998SIAddingPlace, String n1998SIAddingType, String sludgeGenerationName
//            , String sludgeGenerationValue, String sludgeGenerationUse, String sludgeGenerationExplain
//            , String extraCausticNeededName, String extraCausticNeededValue, String extraCausticNeededUse
//            , String extraCausticNeededExplain) throws DocumentException, InvalidFormatException, IOException, URISyntaxException {
//        if (StringUtils.isNotBlank(heroSuggestions)) {
//            buildParagraph(document, heroSuggestions, 9);
//        } else {
//            if (StringUtils.isNotBlank(n1998SISuggestions)) {
//                buildParagraph(document, n1998SISuggestions, 9);
//            }
//
//            List<String> rowTitle1 = new ArrayList<>();
//            rowTitle1.add("推荐药剂");
//            rowTitle1.add("推荐剂量 ppm");
//            rowTitle1.add("投加位置");
//            rowTitle1.add("投加方式");
//            List<List<String>> tableList1 = new ArrayList<>();
//            List<String> valueList1 = new ArrayList<>();
//            valueList1.add(n1998SIProductName);
//            valueList1.add(n1998SIProductValue);
//            valueList1.add(n1998SIAddingPlace);
//            valueList1.add(n1998SIAddingType);
//            tableList1.add(valueList1);
//            buildTable(document, rowTitle1, tableList1, 4, 8310L);
//
//            List<String> rowTitle2 = new ArrayList<>();
//            rowTitle2.add("其他");
//            rowTitle2.add("预估值,kg/m³");
//            rowTitle2.add("用途");
//            rowTitle2.add("说明");
//            List<List<String>> tableList2 = new ArrayList<>();
//            List<String> valueList2 = new ArrayList<>();
//            valueList2.add(sludgeGenerationName);
//            valueList2.add(sludgeGenerationValue);
//            valueList2.add(sludgeGenerationUse);
//            valueList2.add(sludgeGenerationExplain);
//            tableList2.add(valueList2);
//            List<String> valueList3 = new ArrayList<>();
//            valueList3.add(extraCausticNeededName);
//            valueList3.add(extraCausticNeededValue);
//            valueList3.add(extraCausticNeededUse);
//            valueList3.add(extraCausticNeededExplain);
//            tableList2.add(valueList3);
//            buildTable(document, rowTitle2, tableList2, 4, 8310L);
//        }
//
//    }

//    public void buildN3108Table(XWPFDocument document, String n3108ProductName, String n3108Value, String n3108AddingPlace
//            , String n3108AddingType) throws DocumentException, InvalidFormatException, IOException, URISyntaxException {
//        //构建标题
//        List<String> rowTitle = new ArrayList<>();
//        rowTitle.add("推荐药剂");
//        rowTitle.add("推荐剂量 ppm");
//        rowTitle.add("投加位置");
//        rowTitle.add("投加方式");
//        List<List<String>> tableList = new ArrayList<>();
//        List<String> valueList = new ArrayList<>();
//        valueList.add(n3108ProductName);
//        valueList.add(n3108Value);
//        valueList.add(n3108AddingPlace);
//        valueList.add(n3108AddingType);
//        tableList.add(valueList);
//        buildTable(document, rowTitle, tableList, 4, 8310L);
//    }

//    public void buildScaleInhibitorTable(XWPFDocument document, String feedCalciteSrValue, String feedLSIValue
//            , String concentrationfactorValue, String pHValue, String calciteSRValue, String concentrateLSIValue
//            , String caValue, String siO2Value, String mgValue
//            , List<List<String>> tableList) throws DocumentException, InvalidFormatException, IOException, URISyntaxException {
//        List<String> rowTitle1 = new ArrayList<>();
//        rowTitle1.add("进水碳酸钙过饱和度");
//        rowTitle1.add("进水郎利尔过饱和指数");
//        rowTitle1.add("浓缩倍数");
//        rowTitle1.add("pH");
//        rowTitle1.add("碳酸钙过饱和度");
//        rowTitle1.add("浓水郎利尔过饱和指数");
//        rowTitle1.add("钙");
//        rowTitle1.add("二氧化硅");
//        rowTitle1.add("镁");
//        List<List<String>> tableList1 = new ArrayList<>();
//        List<String> valueList1 = new ArrayList<>();
//        valueList1.add(feedCalciteSrValue);
//        valueList1.add(feedLSIValue);
//        valueList1.add(concentrationfactorValue);
//        valueList1.add(pHValue);
//        valueList1.add(calciteSRValue);
//        valueList1.add(concentrateLSIValue);
//        valueList1.add(caValue);
//        valueList1.add(siO2Value);
//        valueList1.add(mgValue);
//        tableList1.add(valueList1);
//        buildTableForScaleInhibitor(document, rowTitle1, tableList1, 9, 8310L, true);
//
//        blankParagraph(document);
//
//        List<String> rowTitle2 = new ArrayList<>();
//        rowTitle2.add("选取产品");
//        rowTitle2.add("浓水加药量");
//        rowTitle2.add("进水加药量");
//        rowTitle2.add("碳酸钙");
//        rowTitle2.add("硫酸钙");
//        rowTitle2.add("磷酸钙");
//        rowTitle2.add("二氧化硅");
//        rowTitle2.add("铝");
//        rowTitle2.add("铁");
//        rowTitle2.add("锰");
//        buildTableForScaleInhibitor(document, rowTitle2, tableList, 10, 8310L, true);
//    }


    /**
     * 添加系统性能预测表格
     *
     * @param document
     * @throws DocumentException
     */
//    public void buildForecastTable(XWPFDocument document, String cfrfValue, String cfrfValueNew, String cipValue
//            , String cipValueNew, String ocfValue, String ocfValueNew) throws DocumentException, InvalidFormatException, IOException, URISyntaxException {
//        //构建标题
//        List<String> rowTitle = new ArrayList<>();
//        rowTitle.add("");
//        rowTitle.add("");
//        rowTitle.add("使用原方案");
//        rowTitle.add("使用新方案");
//        //构建内容
//        List<List<String>> tableList = new ArrayList<>();
//        List<String> row1 = new ArrayList<>();
//        row1.add("保安过滤器滤芯更换周期");
//        row1.add("(天)");
//        row1.add(cfrfValue);
//        row1.add(cfrfValueNew);
//        tableList.add(row1);
//
//        List<String> row2 = new ArrayList<>();
//        row2.add("反渗透在线清洗周期");
//        row2.add("(天)");
//        row2.add(cipValue);
//        row2.add(cipValueNew);
//        tableList.add(row2);
//
//        List<String> row3 = new ArrayList<>();
//        row3.add("反渗透离线清洗周期");
//        row3.add("(天)");
//        row3.add(ocfValue);
//        row3.add(ocfValueNew);
//        tableList.add(row3);
//
//        buildTable(document, rowTitle, tableList, 4, 8310L);
//    }
}
