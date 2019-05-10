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

        //读配置文件
        Helper.loadProperties();

        //读YAML配置文件
        LinkedHashMap yamlProperties = Helper.readYAML();

        //读不规则变化表excel入内存
        Sheet irregularVerbsSheet = Helper.getSheetFormExcelByPathAndName(IRREGULAR_VERBS_EXCEL_NAME, IRREGULAR_VERBS_EXCEL_SHEET_NAME);
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
        Sheet customizeRulesSheet = Helper.getSheetFormExcelByPathAndName(CUSTOMIZE_RULES_EXCEL_NAME, CUSTOMIZE_RULES_EXCEL_SHEET_NAME);
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

            //自定义excel读取到的内容
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
        Sheet taskSheet = Helper.getSheetFormExcelByPathAndName(TASK_EXCEL_NAME, TASK_EXCEL_SHEET_NAME);
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


            logger.debug("taskExcelSheetName: " + TASK_EXCEL_SHEET_NAME + ", rowNumber: " + (i + 1) + ", cell value: " + value);
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
                    logger.debug("[" + taskUnit.getValue() + "], statisfied with suffix: [" + suffix + "]");
                } else {
                    logger.debug("[" + taskUnit.getValue() + "], not statisfied with suffix: [" + suffix + "]");
                }
            }
            taskUnit.setTranslatedValuesSet(translatedValuesSet);
            logger.debug("[" + taskUnit.getValue() + "], translatedValuesSet: [" + translatedValuesSet + "], taskUnit: " + taskUnit.toString());
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
                WORKER_SIZE,
                WORKER_SIZE,
                200,
                TimeUnit.SECONDS,
                new ArrayBlockingQueue<Runnable>(WORKER_SIZE),
                new ThreadPoolExecutor.AbortPolicy() //创建线程大于上限，抛出RejectedExecutionException异常
        );

        //创建线程
        final long startConcurrencyTime = System.currentTimeMillis();
        int taskSize = taskUnitQueue.size();
        for (int i = 0; i < WORKER_SIZE; i++) {
//            Thread.sleep(200);
            executor.submit(new Worker(taskUnitQueue, resultTaskUnitQueue, failureTaskUnitQueue));
        }
        executor.shutdown();
        //阻塞等待完成任务
        while (true) {
            if (executor.isTerminated()) {
                long endTime = System.currentTimeMillis();
                logger.info("all task completed! used time: " + Helper.timeAdapter((endTime - startTime) / 1000));
                break;
            } else {
                int activeCount = executor.getActiveCount();
                final long endTime = System.currentTimeMillis();
                long usedSeond = (endTime - startConcurrencyTime) / 1000;
                int finishedTask = taskSize - taskUnitQueue.size();
//                long finishedTask = executor.getCompletedTaskCount();
                double timePerTask = 1.0 * usedSeond / finishedTask;
                double etaSecond = taskUnitQueue.size() * timePerTask;
                String etaTime = Helper.timeAdapter(new Double(etaSecond).longValue());
                Double progressRate = 1.0 * finishedTask / taskSize;
                logger.info("[usedTime: "
                        + Helper.timeAdapter(usedSeond)
                        + "], [finished/all: " + finishedTask + "/" + taskSize
                        + "], [eta: " + etaTime
                        + "],[activeWorkerSize/workerSize:" + activeCount + "/" + WORKER_SIZE
                        + "], [progressRate: " + String.format("%.2f", progressRate * 100) + "%]");
            }
            Thread.sleep(1000);
        }

        //符合自定义规则的词写入workbook
        Workbook wb = Helper.taskUnitQueueWriteExcel(resultTaskUnitQueue, getWorkbookFormExcelByPath(TASK_EXCEL_NAME), TASK_EXCEL_SHEET_NAME, CUSTOMIZE_RULE_MATCHED_TASK_TYPE);

        //符合不规则动词变化表的词写入workbook
        wb = Helper.taskUnitQueueWriteExcel(irregularVerbsMapMatchedTaskUnitList, wb, TASK_EXCEL_SHEET_NAME, IRREGULAR_VERBS_MATCHED_TASK_TYPE);

        //不符合符合自定义规则的词写入workbook
        wb = Helper.taskUnitQueueWriteExcel(customizeRuleUnitNotMatchedTaskUnidList, wb, TASK_EXCEL_SHEET_NAME, CUSTOMIZE_RULE_NOT_MATCHED_TASK_TYPE);

        //超时错误的词写入workbook
        wb = Helper.taskUnitQueueWriteExcel(failureTaskUnitQueue, wb, TASK_EXCEL_SHEET_NAME, SOCKET_TIMEOUT_TASK_TYPE);

        //写入excel
        Helper.workBookWriteToFile(wb, TASK_EXCEL_NAME);

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
