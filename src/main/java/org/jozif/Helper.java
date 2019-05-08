package org.jozif;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.yaml.snakeyaml.Yaml;

import java.io.*;
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

//        //是待处理的excel则复制
//        LinkedHashMap yamlProperties = Helper.readYAML();
//        //读待处理任务单词相关配置信息
//        LinkedHashMap taskExcelInfo = (LinkedHashMap) yamlProperties.get("taskExcel");
//        String taskExcelPath = (String) taskExcelInfo.get("excelName");
//        logger.info("taskExcelPath: " + taskExcelPath);
//        if (StringUtils.equals(filePath,taskExcelPath)){
//            //写入文件
//            FileOutputStream file = null;
//            try {
//                file = new FileOutputStream(filePath);
//                wb.write(file);
//            } catch (Exception e) {
//                e.printStackTrace();
//            } finally {
//                file.close();
//            }
//        }
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
     * @param filePath
     * @param sheetName
     * @param TYPE
     * @throws IOException
     */
    public static Workbook taskUnitQueueWriteExcel(ArrayBlockingQueue<TaskUnit> taskUnitQueue, Workbook wb, String sheetName, int TYPE) throws Exception {
        LinkedHashMap yamlProperties = Helper.readYAML();

        //读是否显示多个结果相关配置信息
        boolean isshowMultiple = true;
        isshowMultiple = (boolean) yamlProperties.get("isshowMultiple");
        logger.info("isshowMultiple: " + isshowMultiple);

        //符合自定义规则且字典不能查处来的词的颜色
        int customizeRuleMatchedNotFoundInDictTaskColorInRGB = 0x000000;
        customizeRuleMatchedNotFoundInDictTaskColorInRGB = (Integer) yamlProperties.get("customizeRuleMatchedNotFoundInDictTaskColorInRGB");
        logger.info("customizeRuleMatchedNotFoundInDictTaskColorInRGB: " + customizeRuleMatchedNotFoundInDictTaskColorInRGB);

        //符合自定义规则且字典能查出来的词的颜色
        int customizeRuleMatchedFoundInDictTaskColorInRGB = 0x000000;
        customizeRuleMatchedFoundInDictTaskColorInRGB = (Integer) yamlProperties.get("customizeRuleMatchedFoundInDictTaskColorInRGB");
        logger.info("customizeRuleMatchedFoundInDictTaskColorInRGB: " + customizeRuleMatchedFoundInDictTaskColorInRGB);

        //不符合自定义规则且字典能查出来的词的颜色
        int customizeRuleNotMatchedTaskColorInRGB = 0x000000;
        customizeRuleNotMatchedTaskColorInRGB = (Integer) yamlProperties.get("customizeRuleNotMatchedTaskColorInRGB");
        logger.info("customizeRuleNotMatchedTaskColorInRGB: " + customizeRuleNotMatchedTaskColorInRGB);

        //符合不规则动词变化表的词
        int irregularVerbsMatchedTaskColorInRGB = 0x000000;
        irregularVerbsMatchedTaskColorInRGB = (Integer) yamlProperties.get("irregularVerbsMatchedTaskColorInRGB");
        logger.info("irregularVerbsMatchedTaskColorInRGB: " + irregularVerbsMatchedTaskColorInRGB);

        //超时错误的词
        int socketTimeoutTaskColorInRGB = 0x000000;
        socketTimeoutTaskColorInRGB = (Integer) yamlProperties.get("socketTimeoutTaskColorInRGB");
        logger.info("socketTimeoutTaskColorInRGB: " + socketTimeoutTaskColorInRGB);
        
        Sheet sheet = wb.getSheet(sheetName);

        XSSFCellStyle backgroundStyle = (XSSFCellStyle) wb.createCellStyle();
        backgroundStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //符合不规则动词变化表的词的类型代码
        if (IRREGULAR_VERBS_MATCHED_TASK_TYPE == TYPE) {
            //符合不规则动词变化表的词 color irregularVerbsMatchedTaskColorInRGB
            backgroundStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(irregularVerbsMatchedTaskColorInRGB)));
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
            backgroundStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(customizeRuleNotMatchedTaskColorInRGB)));
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
            backgroundStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(socketTimeoutTaskColorInRGB)));
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
                    backgroundStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(customizeRuleMatchedNotFoundInDictTaskColorInRGB)));
                    Cell cell = row.createCell(colCount);
                    cell.setCellValue(taskUnit.getValue());
                    cell.setCellStyle(backgroundStyle);
                } else {
                    //符合自定义规则且字典能查出来多个结果的词 color customizeRuleMatchedFoundInDictTaskColorInRGB
                    backgroundStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(customizeRuleMatchedFoundInDictTaskColorInRGB)));
                    if (isshowMultiple) {
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

    public static void workBookWriteToFile(Workbook wb, String filePath) throws IOException {
        //写入文件
        FileOutputStream file = null;
        try {
            file = new FileOutputStream(filePath);
            wb.write(file);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            file.close();
        }
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
    public static LinkedHashMap readYAML() throws FileNotFoundException {
        InputStream input = Helper.class.getClassLoader().getResourceAsStream("application.yml");
        Yaml yaml = new Yaml();
        return yaml.load(input);
    }


    //符合自定义规则的词的类型代码
    public static final int CUSTOMIZE_RULE_MATCHED_TASK_TYPE = 1;

    //不符合自定义规则的词的类型代码
    public static final int CUSTOMIZE_RULE_NOT_MATCHED_TASK_TYPE = 2;

    //符合不规则动词变化表的词的类型代码
    public static final int IRREGULAR_VERBS_MATCHED_TASK_TYPE = 3;

    //超时错误的词的类型代码
    public static final int SOCKET_TIMEOUT_TASK_TYPE = 4;
}
