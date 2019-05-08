package org.jozif;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.net.URL;
import java.util.*;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;

import static org.jozif.Helper.*;

public class App {

    private static final Logger logger = LoggerFactory.getLogger(App.class);

    public static void main(String[] args) throws Exception {
        //开始时间
        final long startTime = System.currentTimeMillis();

        //读YAML配置文件
        LinkedHashMap yamlProperties = Helper.readYAML();

        //读启用线程数相关配置信息
        Integer workerSize = 1;
        workerSize = (Integer) yamlProperties.get("workersSize");
        logger.info("workersSize: " + workerSize);

        //读动词不规则变化表相关配置信息
        LinkedHashMap irregularVerbsExcelInfo = (LinkedHashMap) yamlProperties.get("irregularVerbsExcel");
        String irregularVerbsExcelPath = (String) irregularVerbsExcelInfo.get("excelName");
        logger.info("irregularVerbsExcelPath: " + irregularVerbsExcelPath);
        String irregularVerbsExcelSheetName = (String) irregularVerbsExcelInfo.get("sheetName");
        logger.info("irregularVerbsExcelSheetName: " + irregularVerbsExcelSheetName);

        //读自定义规则相关配置信息
        LinkedHashMap customizeRulesExcelInfo = (LinkedHashMap) yamlProperties.get("customizeRulesExcel");

        URL url = App.class.getClassLoader().getResource("conf.properties");
        String customizeRulesExcelPath = (String) customizeRulesExcelInfo.get("excelName");

        logger.info("customizeRulesExcelPath: " + customizeRulesExcelPath);
        String customizeRulesExcelSheetName = (String) customizeRulesExcelInfo.get("sheetName");
        logger.info("customizeRulesExcelSheetName: " + customizeRulesExcelSheetName);

        //读待处理任务单词相关配置信息
        LinkedHashMap taskExcelInfo = (LinkedHashMap) yamlProperties.get("taskExcel");
        String taskExcelPath = (String) taskExcelInfo.get("excelName");
        logger.info("taskExcelPath: " + taskExcelPath);
        String taskExcelSheetName = (String) taskExcelInfo.get("sheetName");
        logger.info("taskExcelSheetName: " + taskExcelSheetName);

        //读不规则变化表excel入内存
        Sheet irregularVerbsSheet = Helper.getSheetFormExcelByPathAndName(irregularVerbsExcelPath, irregularVerbsExcelSheetName);
        Map<String, String> irregularVerbsMap = new HashMap<>();
        String irregularVerbsForm;
        String basicFOrm;
        String refinedIrregularVerbsForm;
        String refinedBasicFOrm;
        Row irregularVerbsSheetRow;
        for (int i = 0; i < irregularVerbsSheet.getLastRowNum() - 1; i++) {
            irregularVerbsSheetRow = irregularVerbsSheet.getRow(i);
            if (irregularVerbsSheetRow == null) {
                continue;
            }
            //读第一列,第i行单元格内容
            irregularVerbsSheetRow.getCell(0).setCellType(CellType.STRING);//设置读取转String类型
            irregularVerbsForm = irregularVerbsSheetRow.getCell(0).getStringCellValue();//不规则变化后的形式
            basicFOrm = irregularVerbsSheetRow.getCell(1).getStringCellValue();//对应原形
            refinedIrregularVerbsForm = irregularVerbsForm.trim().toLowerCase();
            refinedBasicFOrm = basicFOrm.trim().toLowerCase();
            logger.debug("excel row number: " + (i + 1) + ", refinedIrregularVerbsForm: " + refinedIrregularVerbsForm + ", refinedBasicFOrm: " + refinedBasicFOrm);
            //不规则动词table
            irregularVerbsMap.put(refinedIrregularVerbsForm, refinedBasicFOrm);
        }
        logger.info("irregular verbs list loaded successfully! irregular verbs list size: " + irregularVerbsMap.size());

        //读取自定义规则excel入内存
        Sheet customizeRulesSheet = Helper.getSheetFormExcelByPathAndName(customizeRulesExcelPath, customizeRulesExcelSheetName);
        List<CustomizeRuleUnit> customizeRuleUnitList = new ArrayList<>();
        Row customizeRulesSheetRow;
        String suffix;
        String newSuffix;
        for (int i = 1; i < customizeRulesSheet.getLastRowNum(); i++) {
            customizeRulesSheetRow = customizeRulesSheet.getRow(i);
            if (customizeRulesSheetRow == null) {
                continue;
            }
            if (customizeRulesSheetRow.getCell(0) == null) {
                continue;
            }
            customizeRulesSheetRow.getCell(0).setCellType(CellType.STRING);
            customizeRulesSheetRow.getCell(1).setCellType(CellType.STRING);

            suffix = customizeRulesSheetRow.getCell(0).getStringCellValue();
            newSuffix = customizeRulesSheetRow.getCell(1).getStringCellValue();

            logger.info("suffix: " + suffix + ", newSuffix: " + newSuffix);

            if (StringUtils.isBlank(suffix)) {
                continue;
            }
            if (StringUtils.isBlank(newSuffix)) {
                continue;
            }
            if (suffix.equals("0")) {
                suffix = "";
            }
            if (newSuffix.equals("0")) {
                newSuffix = "";
            }
            CustomizeRuleUnit customizeRuleUnit = new CustomizeRuleUnit(suffix, newSuffix);
            customizeRuleUnitList.add(customizeRuleUnit);
        }
        logger.info("customize rule list loaded successfully! customize rule list size: " + customizeRuleUnitList.size());

        //读取待处理单词excel入内存
        Sheet taskSheet = Helper.getSheetFormExcelByPathAndName(taskExcelPath, taskExcelSheetName);
        List<TaskUnit> taskUnitList = new ArrayList<>();
        Row taskSheetRow;
        String value;
        for (int i = 0; i < taskSheet.getLastRowNum() - 1; i++) {
            taskSheetRow = taskSheet.getRow(i);
            if (taskSheetRow == null) {
                continue;
            }
            if (taskSheetRow.getCell(0) == null) {
                continue;
            }
            taskSheetRow.getCell(0).setCellType(CellType.STRING);
            value = taskSheetRow.getCell(0).getStringCellValue();
            if (StringUtils.isBlank(value)) {
                continue;
            }


            logger.debug("taskExcelSheetName: " + taskExcelSheetName + ", rowNumber: " + (i + 1) + ", cell value: " + value);
            TaskUnit taskUnit = new TaskUnit(i + 1, value);
            taskUnitList.add(taskUnit);
        }
        logger.info("Task unit list loaded successfully! Task unit list size: " + taskUnitList.size());

        //预处理 任务单词 （删除字符串中非字母字符，若与原字符串不同，加入失败sheet）
        String refinedValue;
        for (TaskUnit taskUnit : taskUnitList) {
            value = taskUnit.getValue();
            refinedValue = value.replaceAll("[^a-zA-Z]", "");
            logger.debug("refinedValue: " + refinedValue);
            taskUnit.setRefinedValues(refinedValue);
        }

        //匹配上不规则动词变化表的list
        ArrayBlockingQueue<TaskUnit> irregularVerbsMapMatchedTaskUnitList = initTaskQueue(taskUnitList, false);
        String irregularVerbsMapMatchedValue = null;
        Iterator<TaskUnit> it = taskUnitList.iterator();
        while (it.hasNext()) {
            TaskUnit taskUnit = it.next();
            value = taskUnit.getValue();
            if (StringUtils.isNotBlank(value)) {
                irregularVerbsMapMatchedValue = irregularVerbsMap.get(value);
                if (StringUtils.isNotBlank(irregularVerbsMapMatchedValue)) {
                    //匹配上的词放入单独的list
                    TaskUnit irregularVerbsMapMatchedTaskUnit = new TaskUnit(taskUnit.getExcelRowNumber(), irregularVerbsMap.get(value));
                    irregularVerbsMapMatchedTaskUnitList.add(irregularVerbsMapMatchedTaskUnit);
                    it.remove();
                }
            }
        }
        logger.info("irregular verbs map matched: " + irregularVerbsMapMatchedTaskUnitList.size() + ", remained word: " + taskUnitList.size());

        ArrayBlockingQueue<TaskUnit> customizeRuleUnitNotMatchedTaskUnidList = initTaskQueue(taskUnitList, false);
        //对比 自定义规则excel 变换，查询
        it = taskUnitList.iterator();
        while (it.hasNext()) {
            TaskUnit taskUnit = it.next();
            value = taskUnit.getValue();
            Set<String> translatedValuesSet = new HashSet<>();
            for (int j = 0; j < customizeRuleUnitList.size(); j++) {
                CustomizeRuleUnit customizeRuleUnit = customizeRuleUnitList.get(j);
                suffix = customizeRuleUnit.getSuffix();
                newSuffix = customizeRuleUnit.getNewSuffix();
                if (value.endsWith(suffix)) {
                    String translatedValue = value;
                    if (StringUtils.isNotEmpty(suffix)) {
                        int index = translatedValue.lastIndexOf(suffix);
                        translatedValue = translatedValue.substring(0, index);
                        translatedValue += newSuffix;
                    } else {
                        translatedValue += newSuffix;
                    }
                    translatedValuesSet.add(translatedValue);
                    logger.debug("[" + taskUnit.getValue() + "], statisfied with suffix: [" + suffix + "], translatedValue: [" + translatedValue + "], taskUnit: " + taskUnit.toString());
                } else {
                    logger.debug("[" + taskUnit.getValue() + "], not statisfied with suffix: [" + suffix + "]");
                }
            }
            taskUnit.setTranslatedValuesSet(translatedValuesSet);

            //不满足自定义规则的词放入单独的list
            if (translatedValuesSet.size() == 0) {
                customizeRuleUnitNotMatchedTaskUnidList.add(taskUnit);
                it.remove();
            }
        }

        ArrayBlockingQueue<TaskUnit> taskUnitQueue = initTaskQueue(taskUnitList, true);
        ArrayBlockingQueue<TaskUnit> resultTaskUnitQueue = initTaskQueue(taskUnitList, false);
        ArrayBlockingQueue<TaskUnit> failureTaskUnitQueue = initTaskQueue(taskUnitList, false);
        //定义了1个核心线程数，最大线程数1个，队列长度2个s
        ThreadPoolExecutor executor = new ThreadPoolExecutor(
                workerSize,
                workerSize,
                200,
                TimeUnit.SECONDS,
                new ArrayBlockingQueue<Runnable>(workerSize),
                new ThreadPoolExecutor.AbortPolicy() //创建线程大于上限，抛出RejectedExecutionException异常
        );

        //创建线程
        final long startConcurrencyTime = System.currentTimeMillis();
        int taskSize = taskUnitQueue.size();
        for (int i = 0; i < workerSize; i++) {
//            Thread.sleep(200);
            executor.submit(new Worker(taskUnitQueue, resultTaskUnitQueue, failureTaskUnitQueue, startConcurrencyTime, taskSize));
        }
        executor.shutdown();

        //阻塞等待完成任务
        while (true) {
            if (executor.isTerminated()) {
                long endTime = System.currentTimeMillis();
                logger.info("all task completed! used time: " + Helper.timeAdapter((endTime - startTime) / 1000));
                break;
            }
            Thread.sleep(1000);
        }

        //符合自定义规则的词写入workbook
        Workbook wb = Helper.taskUnitQueueWriteExcel(resultTaskUnitQueue, getWorkbookFormExcelByPath(taskExcelPath), taskExcelSheetName, CUSTOMIZE_RULE_MATCHED_TASK_TYPE);

        //符合不规则动词变化表的词写入workbook
        wb = Helper.taskUnitQueueWriteExcel(irregularVerbsMapMatchedTaskUnitList, wb, taskExcelSheetName, IRREGULAR_VERBS_MATCHED_TASK_TYPE);

        //不符合符合自定义规则的词写入workbook
        wb = Helper.taskUnitQueueWriteExcel(customizeRuleUnitNotMatchedTaskUnidList, wb, taskExcelSheetName, CUSTOMIZE_RULE_NOT_MATCHED_TASK_TYPE);

        //超时错误的词写入workbook
        wb = Helper.taskUnitQueueWriteExcel(failureTaskUnitQueue, wb, taskExcelSheetName, SOCKET_TIMEOUT_TASK_TYPE);

        //写入excel
        Helper.workBookWriteToFile(wb, taskExcelPath);

        logger.info("result excel generated successfully! used time: " + Helper.timeAdapter((System.currentTimeMillis() - startTime) / 1000));
    }

    public static ArrayBlockingQueue<TaskUnit> initTaskQueue(List<TaskUnit> taskUnitList, boolean isLoad) {
        logger.info("start initialize task queue");

        //initialize task
        ArrayBlockingQueue<TaskUnit> tasks = new ArrayBlockingQueue<>(taskUnitList.size());
        if (isLoad) {
            for (TaskUnit taskUnit : taskUnitList) {
                tasks.add(taskUnit);
            }
            logger.info("task queue loaded success!");
        }

        logger.info("task queue has been initialized, taskSize: " + tasks.size());
        return tasks;
    }
}
