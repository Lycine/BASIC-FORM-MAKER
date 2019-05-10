package org.jozif;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.yaml.snakeyaml.Yaml;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Set;
import java.util.concurrent.ArrayBlockingQueue;

public class Helper {

    private static final Logger logger = LoggerFactory.getLogger(Helper.class);

    public static Sheet getSheetFormExcelByPathAndName(String filePath, String sheetName) throws Exception {
        Workbook wb = getWorkbookFormExcelByPath(filePath);
        Sheet sheet = wb.getSheet(sheetName);
        if (null == sheet) {
            logger.error("sheet不正确");
            throw new Exception("sheet不正确");
        }
        return sheet;
    }

    public static Workbook getWorkbookFormExcelByPath(String filePath) throws Exception {
        InputStream is = null;
        logger.info("filePath: " + filePath);
        is = Helper.class.getClassLoader().getResourceAsStream(filePath);
        Workbook wb = null;
        try {
            wb = new XSSFWorkbook(is);
            is.close();
        } catch (IOException e) {
            logger.error(e.getMessage(), e);
        }
        if (null == wb) {
            logger.error("excel文件不正确");
            throw new Exception("excel文件不正确");
        }
        return wb;
    }

    /**
     * 符合不规则动词变化表的词 1
     * <p>
     * 不符合自定义规则的词 2
     * <p>
     * 符合自定义规则且字典能查出来的词 3.1
     * 符合自定义规则且字典不能查处来的词 3.2
     *
     * @param taskUnitQueue
     * @param sheetName
     * @param TYPE
     * @throws IOException
     */
    public static Workbook taskUnitQueueWriteExcel(ArrayBlockingQueue<TaskUnit> taskUnitQueue, Workbook wb, String sheetName, int TYPE) throws Exception {

        Sheet sheet = wb.getSheet(sheetName);

        XSSFCellStyle backgroundStyle = (XSSFCellStyle) wb.createCellStyle();
        backgroundStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //符合不规则动词变化表的词的类型代码
        if (IRREGULAR_VERBS_MATCHED_TASK_TYPE == TYPE) {
            //符合不规则动词变化表的词 color irregularVerbsMatchedTaskColorInRGB
            backgroundStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(IRREGULAR_VERBS_MATCHED_TASK_COLOR_IN_RGB)));
            for (TaskUnit taskUnit : taskUnitQueue) {
                Row row = sheet.getRow(taskUnit.getExcelRowNumber() - 1);
                Cell cell = row.createCell(1);
                cell.setCellValue(taskUnit.getValue());
                cell.setCellStyle(backgroundStyle);
            }
        }

        //不符合自定义规则的词
        if (CUSTOMIZE_RULE_NOT_MATCHED_TASK_TYPE == TYPE) {
            //不符合自定义规则的词 color customizeRuleNotMatchedTaskColorInRGB
            backgroundStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(CUSTOMIZE_RULE_NOT_MATCHED_TASK_COLOR_IN_RGB)));
            for (TaskUnit taskUnit : taskUnitQueue) {
                Row row = sheet.getRow(taskUnit.getExcelRowNumber() - 1);
                Cell cell = row.createCell(1);
                cell.setCellValue(taskUnit.getValue());
                cell.setCellStyle(backgroundStyle);
            }
        }

        //超时错误的词
        if (SOCKET_TIMEOUT_TASK_TYPE == TYPE) {
            //超时错误的词 color socketTimeoutTaskColorInRGB
            backgroundStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(SOCKET_TIMEOUT_TASK_COLOR_IN_RGB)));
            for (TaskUnit taskUnit : taskUnitQueue) {
                Row row = sheet.getRow(taskUnit.getExcelRowNumber() - 1);
                Cell cell = row.createCell(1);
                cell.setCellValue(taskUnit.getValue());
                cell.setCellStyle(backgroundStyle);
            }
        }
        //符合自定义规则的词
        if (CUSTOMIZE_RULE_MATCHED_TASK_TYPE == TYPE) {
            for (TaskUnit taskUnit : taskUnitQueue) {
                Row row = sheet.getRow(taskUnit.getExcelRowNumber() - 1);
                Set resultValueSet = taskUnit.getResultValuesSet();
                Iterator<String> it = resultValueSet.iterator();
                int colCount = 1;

                if (!it.hasNext()) {
                    //符合自定义规则且字典查不出来结果的词 color customizeRuleMatchedNotFoundInDictTaskColorInRGB
                    backgroundStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(CUSTOMIZE_RULE_MATCHED_NOT_FOUND_IN_DICT_TASK_COLOR_IN_RGB)));
                    Cell cell = row.createCell(colCount);
                    cell.setCellValue(taskUnit.getValue());
                    cell.setCellStyle(backgroundStyle);
                } else {
                    //符合自定义规则且字典能查出来多个结果的词 color customizeRuleMatchedFoundInDictTaskColorInRGB
                    backgroundStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(Helper.CUSTOMIZE_RULE_MATCHED_FOUND_IN_DICT_TASK_COLOR_IN_RGB)));
                    if (IS_SHOW_MULTIPLE_RESULT) {
                        while (it.hasNext()) {
                            Cell cell = row.createCell(colCount);
                            cell.setCellValue(it.next());
                            cell.setCellStyle(backgroundStyle);
                            colCount += 1;
                        }
                        if (colCount > 2) {
                            logger.warn("task have multiple result value: " + taskUnit.toString());
                        }
                    } else {
                        Cell cell = row.createCell(colCount);
                        cell.setCellValue(it.next());
                        cell.setCellStyle(backgroundStyle);
                        if (colCount > 1) {
                            logger.warn("task have multiple result value， but not show, value: " + taskUnit.toString());
                        }
                    }
                }
            }
        }
        return wb;
    }

    public static void workBookWriteToFile(Workbook wb, String resultFileName) throws IOException {
        LocalDateTime localDateTime = LocalDateTime.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyymmddHHmmss");
        String timestampString = formatter.format(localDateTime);
        String fullFilePath = RESULT_FILE_PATH + timestampString + "_" + resultFileName;

        //写入文件
        FileOutputStream file = null;
        try {
            file = new FileOutputStream(fullFilePath);
            wb.write(file);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            file.close();
        }
    }

    public static Workbook polishWorkBookBeforeGenerate(Workbook wb) {

        Sheet oldSheet = wb.getSheet(TASK_EXCEL_SHEET_NAME);
        int oldSheetIndex = wb.getSheetIndex(oldSheet);
        wb.setSheetName(oldSheetIndex, RESULT_EXCEL_SHEET_NAME);

        Sheet legendSheet = wb.createSheet();

        int lastRowNum = oldSheet.getLastRowNum();
        logger.info("lastRowNum: " + lastRowNum);

        Row row0 = legendSheet.createRow(0);
        XSSFCellStyle backgroundStyle0 = (XSSFCellStyle) wb.createCellStyle();
        backgroundStyle0.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        backgroundStyle0.setFillForegroundColor(new XSSFColor(new java.awt.Color(CUSTOMIZE_RULE_MATCHED_NOT_FOUND_IN_DICT_TASK_COLOR_IN_RGB)));
        Cell cell0 = row0.createCell(0);
        cell0.setCellValue("符合自定义规则且字典不能查出来的词的颜色");
        cell0.setCellStyle(backgroundStyle0);

        Row row1 = legendSheet.createRow(0 + 1);
        XSSFCellStyle backgroundStyle1 = (XSSFCellStyle) wb.createCellStyle();
        backgroundStyle1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        backgroundStyle1.setFillForegroundColor(new XSSFColor(new java.awt.Color(CUSTOMIZE_RULE_MATCHED_FOUND_IN_DICT_TASK_COLOR_IN_RGB)));
        Cell cell1 = row1.createCell(0);
        cell1.setCellValue("符合自定义规则且字典能查出来的词的颜色");
        cell1.setCellStyle(backgroundStyle1);

        Row row2 = legendSheet.createRow(0 + 2);
        XSSFCellStyle backgroundStyle2 = (XSSFCellStyle) wb.createCellStyle();
        backgroundStyle2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        backgroundStyle2.setFillForegroundColor(new XSSFColor(new java.awt.Color(SOCKET_TIMEOUT_TASK_COLOR_IN_RGB)));
        Cell cell2 = row2.createCell(0);
        cell2.setCellValue("超时错误的词的颜色");
        cell2.setCellStyle(backgroundStyle2);

        Row row3 = legendSheet.createRow(0 + 3);
        XSSFCellStyle backgroundStyle3 = (XSSFCellStyle) wb.createCellStyle();
        backgroundStyle3.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        backgroundStyle3.setFillForegroundColor(new XSSFColor(new java.awt.Color(CUSTOMIZE_RULE_NOT_MATCHED_TASK_COLOR_IN_RGB)));
        Cell cell3 = row3.createCell(0);
        cell3.setCellStyle(backgroundStyle3);
        cell3.setCellValue("不符合自定义规则且字典能查出来的词的颜色");


        Row row4 = legendSheet.createRow(0 + 4);
        XSSFCellStyle backgroundStyle4 = (XSSFCellStyle) wb.createCellStyle();
        backgroundStyle4.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        backgroundStyle4.setFillForegroundColor(new XSSFColor(new java.awt.Color(IRREGULAR_VERBS_MATCHED_TASK_COLOR_IN_RGB)));
        Cell cell4 = row4.createCell(0);
        cell4.setCellValue("符合不规则动词变化表的词的颜色");
        cell4.setCellStyle(backgroundStyle4);

        Row row5 = legendSheet.createRow(0 + 5);
        XSSFCellStyle backgroundStyle5 = (XSSFCellStyle) wb.createCellStyle();
        backgroundStyle5.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        backgroundStyle5.setFillForegroundColor(new XSSFColor(new java.awt.Color(UNKNOWN_TASK_COLOR_IN_RGB)));
        Cell cell5 = row5.createCell(0);
        cell5.setCellValue("其他错误的颜色");
        cell5.setCellStyle(backgroundStyle5);

        legendSheet.autoSizeColumn(0, true);
        int legendSheetIndex = wb.getSheetIndex(legendSheet);
        wb.setSheetName(legendSheetIndex, "legend");

        return wb;
    }

    //秒转成时间
    public static String timeAdapter(long second) {
        long hour = second / 3600;
        second = second % 3600;
        long minute = second / 60;
        second = second % 60;
        String result = String.format("%0" + 2 + "d", hour) + ":" + String.format("%0" + 2 + "d", minute) + ":" + String.format("%0" + 2 + "d", second);
        return result;
    }

    //读YAML配置文件
    public static LinkedHashMap readYAML() {
        InputStream input = Helper.class.getClassLoader().getResourceAsStream("application.yml");
        Yaml yaml = new Yaml();
        return yaml.load(input);
    }

    public static void loadProperties() {
        LinkedHashMap yamlProperties = readYAML();

        //读是否显示多个结果相关配置信息
        IS_SHOW_MULTIPLE_RESULT = (boolean) yamlProperties.get("IS_SHOW_MULTIPLE_RESULT");
        logger.info("IS_SHOW_MULTIPLE_RESULT: " + IS_SHOW_MULTIPLE_RESULT);

        //符合自定义规则且字典不能查出来的词的颜色
        CUSTOMIZE_RULE_MATCHED_NOT_FOUND_IN_DICT_TASK_COLOR_IN_RGB = (Integer) yamlProperties.get("CUSTOMIZE_RULE_MATCHED_NOT_FOUND_IN_DICT_TASK_COLOR_IN_RGB");
        logger.info("CUSTOMIZE_RULE_MATCHED_NOT_FOUND_IN_DICT_TASK_COLOR_IN_RGB: " + CUSTOMIZE_RULE_MATCHED_NOT_FOUND_IN_DICT_TASK_COLOR_IN_RGB);

        //符合自定义规则且字典能查出来的词的颜色
        CUSTOMIZE_RULE_MATCHED_FOUND_IN_DICT_TASK_COLOR_IN_RGB = (Integer) yamlProperties.get("CUSTOMIZE_RULE_MATCHED_FOUND_IN_DICT_TASK_COLOR_IN_RGB");
        logger.info("CUSTOMIZE_RULE_MATCHED_FOUND_IN_DICT_TASK_COLOR_IN_RGB: " + CUSTOMIZE_RULE_MATCHED_FOUND_IN_DICT_TASK_COLOR_IN_RGB);

        //不符合自定义规则且字典能查出来的词的颜色
        CUSTOMIZE_RULE_NOT_MATCHED_TASK_COLOR_IN_RGB = (Integer) yamlProperties.get("CUSTOMIZE_RULE_NOT_MATCHED_TASK_COLOR_IN_RGB");
        logger.info("CUSTOMIZE_RULE_NOT_MATCHED_TASK_COLOR_IN_RGB: " + CUSTOMIZE_RULE_NOT_MATCHED_TASK_COLOR_IN_RGB);

        //符合不规则动词变化表的词
        IRREGULAR_VERBS_MATCHED_TASK_COLOR_IN_RGB = (Integer) yamlProperties.get("IRREGULAR_VERBS_MATCHED_TASK_COLOR_IN_RGB");
        logger.info("IRREGULAR_VERBS_MATCHED_TASK_COLOR_IN_RGB: " + IRREGULAR_VERBS_MATCHED_TASK_COLOR_IN_RGB);

        //超时错误的词
        SOCKET_TIMEOUT_TASK_COLOR_IN_RGB = (Integer) yamlProperties.get("SOCKET_TIMEOUT_TASK_COLOR_IN_RGB");
        logger.info("SOCKET_TIMEOUT_TASK_COLOR_IN_RGB: " + SOCKET_TIMEOUT_TASK_COLOR_IN_RGB);

        //其他错误的颜色
        UNKNOWN_TASK_COLOR_IN_RGB = (Integer) yamlProperties.get("UNKNOWN_TASK_COLOR_IN_RGB");
        logger.info("UNKNOWN_TASK_COLOR_IN_RGB: " + UNKNOWN_TASK_COLOR_IN_RGB);

        //jsoup超时时间（秒）
        TIMEOUT_SECOND = (Integer) yamlProperties.get("TIMEOUT_SECOND");
        logger.info("TIMEOUT_SECOND: " + TIMEOUT_SECOND);

        //读启用线程数相关配置信息
        WORKER_SIZE = (Integer) yamlProperties.get("WORKER_SIZE");
        logger.info("WORKER_SIZE: " + WORKER_SIZE);

        //读动词不规则变化表相关配置信息
        LinkedHashMap irregularVerbsExcelInfo = (LinkedHashMap) yamlProperties.get("IRREGULAR_VERBS_EXCEL");
        IRREGULAR_VERBS_EXCEL_NAME = (String) irregularVerbsExcelInfo.get("EXCEL_NAME");
        logger.info("IRREGULAR_VERBS_EXCEL_NAME: " + IRREGULAR_VERBS_EXCEL_NAME);
        IRREGULAR_VERBS_EXCEL_SHEET_NAME = (String) irregularVerbsExcelInfo.get("SHEET_NAME");
        logger.info("IRREGULAR_VERBS_EXCEL_SHEET_NAME: " + IRREGULAR_VERBS_EXCEL_SHEET_NAME);

        //读自定义规则相关配置信息
        LinkedHashMap customizeRulesExcelInfo = (LinkedHashMap) yamlProperties.get("CUSTOMIZE_RULES_EXCEL");
        CUSTOMIZE_RULES_EXCEL_NAME = (String) customizeRulesExcelInfo.get("EXCEL_NAME");
        logger.info("CUSTOMIZE_RULES_EXCEL_NAME: " + CUSTOMIZE_RULES_EXCEL_NAME);
        CUSTOMIZE_RULES_EXCEL_SHEET_NAME = (String) customizeRulesExcelInfo.get("SHEET_NAME");
        logger.info("CUSTOMIZE_RULES_EXCEL_SHEET_NAME: " + CUSTOMIZE_RULES_EXCEL_SHEET_NAME);

        //读待处理任务单词相关配置信息
        LinkedHashMap taskExcelInfo = (LinkedHashMap) yamlProperties.get("TASK_EXCEL");
        TASK_EXCEL_NAME = (String) taskExcelInfo.get("EXCEL_NAME");
        logger.info("TASK_EXCEL_NAME: " + TASK_EXCEL_NAME);
        TASK_EXCEL_SHEET_NAME = (String) taskExcelInfo.get("SHEET_NAME");
        logger.info("TASK_EXCEL_SHEET_NAME: " + TASK_EXCEL_SHEET_NAME);

        //读结果excel sheet 名称
        RESULT_EXCEL_SHEET_NAME = (String) yamlProperties.get("RESULT_EXCEL_SHEET_NAME");
        logger.info("RESULT_EXCEL_SHEET_NAME: " + RESULT_EXCEL_SHEET_NAME);

        //读结果excel名称
        RESULT_EXCEL_NAME = (String) yamlProperties.get("RESULT_EXCEL_NAME");
        logger.info("RESULT_EXCEL_NAME: " + RESULT_EXCEL_NAME);
    }

    //结果excel名称
    public static String RESULT_EXCEL_NAME = "";

    //结果excel sheet名称
    public static String RESULT_EXCEL_SHEET_NAME = "";

    public static String RESULT_FILE_PATH = "";

    //待处理任务单词相关配置信息
    public static String TASK_EXCEL_SHEET_NAME = "word";

    public static String TASK_EXCEL_NAME = "word.xlsx";

    //自定义规则相关配置信息
    public static String CUSTOMIZE_RULES_EXCEL_SHEET_NAME = "customizeRules";

    public static String CUSTOMIZE_RULES_EXCEL_NAME = "customizeRules.xlsx";

    //动词不规则变化表相关配置信息
    public static String IRREGULAR_VERBS_EXCEL_SHEET_NAME = "irregularVerbs";

    public static String IRREGULAR_VERBS_EXCEL_NAME = "irregularVerbs.xlsx";

    // 线程数
    public static int WORKER_SIZE = 1;

    //jsoup超时时间（秒）
    public static int TIMEOUT_SECOND = 10;

    //不符合自定义规则且字典能查出来的词的颜色
    public static int CUSTOMIZE_RULE_NOT_MATCHED_TASK_COLOR_IN_RGB = 0x000000;

    //符合自定义规则且字典不能查出来的词的颜色
    public static int CUSTOMIZE_RULE_MATCHED_NOT_FOUND_IN_DICT_TASK_COLOR_IN_RGB = 0x000000;

    //符合自定义规则且字典能查出来的词的颜色
    public static int CUSTOMIZE_RULE_MATCHED_FOUND_IN_DICT_TASK_COLOR_IN_RGB = 0x000000;

    //符合不规则动词变化表的词
    public static int IRREGULAR_VERBS_MATCHED_TASK_COLOR_IN_RGB = 0x000000;

    //超时错误的词
    public static int SOCKET_TIMEOUT_TASK_COLOR_IN_RGB = 0x000000;

    //其他错误的颜色
    public static int UNKNOWN_TASK_COLOR_IN_RGB = 0x000000;

    //读是否显示多个结果相关配置信息
    public static boolean IS_SHOW_MULTIPLE_RESULT = true;

    //符合自定义规则的词的类型代码
    public static final int CUSTOMIZE_RULE_MATCHED_TASK_TYPE = 1;

    //不符合自定义规则的词的类型代码
    public static final int CUSTOMIZE_RULE_NOT_MATCHED_TASK_TYPE = 2;

    //符合不规则动词变化表的词的类型代码
    public static final int IRREGULAR_VERBS_MATCHED_TASK_TYPE = 3;

    //超时错误的词的类型代码
    public static final int SOCKET_TIMEOUT_TASK_TYPE = 4;
}
