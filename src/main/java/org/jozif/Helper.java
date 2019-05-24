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
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;

public class Helper {

    private static final Logger logger = LoggerFactory.getLogger(Helper.class);

    //读取文件
    public static Set<File> readFile(String filePath) {
        Set<File> fileSet = new HashSet<>();
        File file = new File(filePath);
        if (!file.isDirectory()) {
            logger.info("name = " + file.getName() + " is file, absolute path = " + file.getAbsolutePath());
            fileSet.add(file);

        } else if (file.isDirectory()) {
            String[] fileArray = file.list();
            logger.info("name = " + file.getName() + " is directory, absolute path = " + file.getAbsolutePath() + ", fileArray: " + Arrays.toString(fileArray));
            for (int i = 0; i < fileArray.length; i++) {
                File subFile = new File(filePath + System.getProperty("file.separator") + fileArray[i]);
                if (!subFile.isDirectory()) {
                    logger.info("name = " + subFile.getName() + " is file, absolute path = " + subFile.getAbsolutePath());
                    fileSet.add(subFile);
                } else if (subFile.isDirectory()) {
                    fileSet.addAll(readFile(filePath + System.getProperty("file.separator") + fileArray[i]));
                }
            }

        }
        return fileSet;
    }

    //筛选出要批量处理的所有单词文件
    public static Set<File> getPendingTaskFile(Set<File> fileSet) {
        logger.info("fileSet siez: " + fileSet.size());
        try {
            Iterator<File> it = fileSet.iterator();
            if (!it.hasNext()) {
                File file = it.next();
                if (!file.getName().endsWith(TASK_EXCEL_NAME)) {
                    it.remove();
                }
            }
        } catch (NoSuchElementException nsee) {
            logger.error("符合条件的待处理文件数量为0");
            System.exit(0);
        }
        return fileSet;
    }

    public static Workbook process() throws Exception {

        // 本机处理

        //预处理 任务单词 （删除字符串中非字母字符，若与原字符串不同，加入失败sheet）
        taskUnitList = preprocessed(taskUnitList);

        irregularVerbsMapMatchedTaskUnitList = initTaskQueue(taskUnitList, false);
        //匹配上不规则动词变化表的list
        taskUnitList = compareWithIrregularVerbsSheet(taskUnitList);

        customizeRuleUnitNotMatchedTaskUnidList = initTaskQueue(taskUnitList, false);
        //对比 自定义规则excel 变换，查询
        taskUnitList = compareWithCustomizeVerbsSheet(taskUnitList);

        //联网处理

        ArrayBlockingQueue<TaskUnit> taskUnitQueue = initTaskQueue(taskUnitList, true);
        ArrayBlockingQueue<TaskUnit> resultTaskUnitQueue = initTaskQueue(taskUnitList, false);
        ArrayBlockingQueue<TaskUnit> failureTaskUnitQueue = initTaskQueue(taskUnitList, false);

        logger.debug("taskUnitQueueSize: " + taskUnitQueue.size() + ", taskUnitQueue: " + taskUnitQueue.toString());

        //创建线程池
        //定义了1个核心线程数，最大线程数1个，队列长度2个s
        ThreadPoolExecutor executor = new ThreadPoolExecutor(
                WORKER_SIZE,
                WORKER_SIZE,
                200,
                TimeUnit.SECONDS,
                new ArrayBlockingQueue<Runnable>(WORKER_SIZE),
                new ThreadPoolExecutor.AbortPolicy() //创建线程大于上限，抛出RejectedExecutionException异常
        );

        //分配任务
        final long startConcurrencyTime = System.currentTimeMillis();
        int taskSize = taskUnitQueue.size();
        for (int i = 0; i < WORKER_SIZE; i++) {
            executor.submit(new Worker(taskUnitQueue, resultTaskUnitQueue, failureTaskUnitQueue));
        }
        executor.shutdown();

        long usedSeond = (System.currentTimeMillis() - startConcurrencyTime) / 1000;
        //阻塞等待完成任务
        while (usedSeond < APP_TIMEOUT_MINUTE * 60) {
            if (executor.isTerminated()) {
                long endTime = System.currentTimeMillis();
                logger.info("all task completed! used time: " + Helper.timeAdapter((endTime - startTime) / 1000));
                break;
            } else {
                int activeCount = executor.getActiveCount();
                usedSeond = (System.currentTimeMillis() - startConcurrencyTime) / 1000;
                int finishedTask = taskSize - taskUnitQueue.size();
                double timePerTask = 1.0 * usedSeond / finishedTask;
                double etaSecond = taskUnitQueue.size() * timePerTask;
                String etaTime = Helper.timeAdapter(new Double(etaSecond).longValue());
                Double progressRate = 1.0 * finishedTask / taskSize;
                logger.info("[usedTime: "
                        + Helper.timeAdapter(usedSeond)
                        + "], [finished/all: " + finishedTask + "/" + taskSize
                        + "], [current segment eta: " + etaTime
                        + "],[activeWorkerSize/workerSize:" + activeCount + "/" + WORKER_SIZE
                        + "], [progressRate: " + String.format("%.2f", progressRate * 100) + "%]");
            }
            Thread.sleep(1000);
        }
        //线程中的任务还没有结束
        if (!executor.isTerminated()) {
            //强制关闭所有正在进行的进程
            executor.shutdownNow();
            logger.warn("Use time has exceeded the APP_TIMEOUT_MINUTE(" + APP_TIMEOUT_MINUTE + "), terminated all thread, completed task write to excel");
        }
        logger.info("resultTaskUnitQueueSize: " + resultTaskUnitQueue.size());
        //符合自定义规则的词写入workbook
//        Workbook wb = Helper.taskUnitQueueWriteExcel(resultTaskUnitQueue, getWorkbookFormExcelByPath(TASK_EXCEL_NAME), TASK_EXCEL_SHEET_NAME, CUSTOMIZE_RULE_MATCHED_TASK_TYPE);
        Workbook wb = Helper.taskUnitQueueWriteExcel(resultTaskUnitQueue, getWorkbookFormExcelByPath(CURRENT_TASK_EXCEL_NAME), TASK_EXCEL_SHEET_NAME, CUSTOMIZE_RULE_MATCHED_TASK_TYPE);

        //符合不规则动词变化表的词写入workbook
        wb = Helper.taskUnitQueueWriteExcel(irregularVerbsMapMatchedTaskUnitList, wb, TASK_EXCEL_SHEET_NAME, IRREGULAR_VERBS_MATCHED_TASK_TYPE);

        //不符合符合自定义规则的词写入workbook
        wb = Helper.taskUnitQueueWriteExcel(customizeRuleUnitNotMatchedTaskUnidList, wb, TASK_EXCEL_SHEET_NAME, CUSTOMIZE_RULE_NOT_MATCHED_TASK_TYPE);

        //超时错误的词写入workbook
        wb = Helper.taskUnitQueueWriteExcel(failureTaskUnitQueue, wb, TASK_EXCEL_SHEET_NAME, SOCKET_TIMEOUT_TASK_TYPE);

        //规定时间内没有执行完的任务写入workbook
        wb = Helper.taskUnitQueueWriteExcel(taskUnitQueue, wb, TASK_EXCEL_SHEET_NAME, APP_TIMEOUT_TASK_TYPE);


        //清空taskUnitQueue
        while (taskUnitQueue.size() > 0) {
            TaskUnit taskUnit = taskUnitQueue.poll();
            logger.warn("unfinished task: " + taskUnit.toString());
        }

        return wb;
    }


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
        if (null == is) {
            is = new FileInputStream(new File(filePath));
        }
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

        //规定时间内没有执行完的任务
        if (APP_TIMEOUT_TASK_TYPE == TYPE) {
            //规定时间内没有执行完的任务 color socketTimeoutTaskColorInRGB
            backgroundStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(APP_TIMEOUT_TASK_COLOR_IN_RGB)));
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
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM.dd_HH.mm.ss");
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

        Row row1 = legendSheet.createRow(1);
        XSSFCellStyle backgroundStyle1 = (XSSFCellStyle) wb.createCellStyle();
        backgroundStyle1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        backgroundStyle1.setFillForegroundColor(new XSSFColor(new java.awt.Color(CUSTOMIZE_RULE_MATCHED_FOUND_IN_DICT_TASK_COLOR_IN_RGB)));
        Cell cell1 = row1.createCell(0);
        cell1.setCellValue("符合自定义规则且字典能查出来的词的颜色");
        cell1.setCellStyle(backgroundStyle1);

        Row row2 = legendSheet.createRow(2);
        XSSFCellStyle backgroundStyle2 = (XSSFCellStyle) wb.createCellStyle();
        backgroundStyle2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        backgroundStyle2.setFillForegroundColor(new XSSFColor(new java.awt.Color(SOCKET_TIMEOUT_TASK_COLOR_IN_RGB)));
        Cell cell2 = row2.createCell(0);
        cell2.setCellValue("超时错误的词的颜色");
        cell2.setCellStyle(backgroundStyle2);

        Row row3 = legendSheet.createRow(3);
        XSSFCellStyle backgroundStyle3 = (XSSFCellStyle) wb.createCellStyle();
        backgroundStyle3.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        backgroundStyle3.setFillForegroundColor(new XSSFColor(new java.awt.Color(CUSTOMIZE_RULE_NOT_MATCHED_TASK_COLOR_IN_RGB)));
        Cell cell3 = row3.createCell(0);
        cell3.setCellStyle(backgroundStyle3);
        cell3.setCellValue("不符合自定义规则且字典能查出来的词的颜色");

        Row row4 = legendSheet.createRow(4);
        XSSFCellStyle backgroundStyle4 = (XSSFCellStyle) wb.createCellStyle();
        backgroundStyle4.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        backgroundStyle4.setFillForegroundColor(new XSSFColor(new java.awt.Color(IRREGULAR_VERBS_MATCHED_TASK_COLOR_IN_RGB)));
        Cell cell4 = row4.createCell(0);
        cell4.setCellValue("符合不规则动词变化表的词的颜色");
        cell4.setCellStyle(backgroundStyle4);

        Row row5 = legendSheet.createRow(5);
        XSSFCellStyle backgroundStyle5 = (XSSFCellStyle) wb.createCellStyle();
        backgroundStyle5.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        backgroundStyle5.setFillForegroundColor(new XSSFColor(new java.awt.Color(UNKNOWN_TASK_COLOR_IN_RGB)));
        Cell cell5 = row5.createCell(0);
        cell5.setCellValue("其他错误的颜色");
        cell5.setCellStyle(backgroundStyle5);

        Row row6 = legendSheet.createRow(6);
        XSSFCellStyle backgroundStyle6 = (XSSFCellStyle) wb.createCellStyle();
        backgroundStyle6.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        backgroundStyle6.setFillForegroundColor(new XSSFColor(new java.awt.Color(APP_TIMEOUT_TASK_COLOR_IN_RGB)));
        Cell cell6 = row6.createCell(0);
        cell6.setCellValue("规定时间内没有执行完的任务的颜色");
        cell6.setCellStyle(backgroundStyle6);

        legendSheet.autoSizeColumn(0, true);
        int legendSheetIndex = wb.getSheetIndex(legendSheet);
        try {
            wb.setSheetName(legendSheetIndex, "legend");
        } catch (Exception e) {
            e.printStackTrace();
        }

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

    //读不规则变化表excel入内存
    public static Map<String, String> loadIrregularVerbsSheet() throws Exception {
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
        return irregularVerbsMap;
    }

    //读取自定义规则excel入内存
    public static List<CustomizeRuleUnit> loadCustomizeRulesSheet() throws Exception {
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
        return customizeRuleUnitList;
    }

    public static List<TaskUnit> loadTaskSheet() throws Exception {
        return loadTaskSheet(TASK_EXCEL_NAME, TASK_EXCEL_SHEET_NAME);
    }

    public static List<TaskUnit> loadTaskSheet(Sheet taskSheet, int CellNum) throws Exception {
        List<TaskUnit> taskUnitList = new ArrayList<>();
        Row taskSheetRow;
        String value;
        for (int i = 0; i < taskSheet.getLastRowNum() - 1; i++) {
            taskSheetRow = taskSheet.getRow(i);
            if (taskSheetRow == null) {
                continue;
            }
            if (taskSheetRow.getCell(CellNum) == null) {
                continue;
            }
            taskSheetRow.getCell(CellNum).setCellType(CellType.STRING);
            value = taskSheetRow.getCell(CellNum).getStringCellValue();
            if (StringUtils.isBlank(value)) {
                continue;
            }


            logger.debug("taskExcelSheetName: " + TASK_EXCEL_SHEET_NAME + ", rowNumber: " + (i + 1) + ", cell value: " + value);
            TaskUnit taskUnit = new TaskUnit(i + 1, value);
            taskUnitList.add(taskUnit);
        }
        logger.info("Task unit list loaded successfully! Task unit list size: " + taskUnitList.size());
        return taskUnitList;
    }

    public static List<TaskUnit> loadTaskSheet(Sheet taskSheet) throws Exception {
        return loadTaskSheet(taskSheet, 0);
    }


    //读取待处理单词excel入内存
    public static List<TaskUnit> loadTaskSheet(String taskExcelName, String taskExcelSheetName) throws Exception {
        Sheet taskSheet;
        if (StringUtils.isNotBlank(taskExcelName) && StringUtils.isNotBlank(taskExcelSheetName)) {
            taskSheet = Helper.getSheetFormExcelByPathAndName(taskExcelName, taskExcelSheetName);
        } else {
            logger.error("taskExcelName or taskExcelSheetName is empty");
            throw new Exception("读取待处理单词sheet失败，检查excel文件名和sheet名");
        }
        return loadTaskSheet(taskSheet);
    }

    //预处理 任务单词 （删除字符串中非字母字符，若与原字符串不同，加入失败sheet）
    public static List<TaskUnit> preprocessed(List<TaskUnit> taskUnitList) {
        String refinedValue;
        for (TaskUnit taskUnit : taskUnitList) {
            String value = taskUnit.getValue();
            refinedValue = value.replaceAll("[^a-zA-Z]", "").trim();
            logger.debug("refinedValue: " + refinedValue);
            taskUnit.setRefinedValues(refinedValue);
        }
        return taskUnitList;
    }

    //匹配上不规则动词变化表的list
    public static List<TaskUnit> compareWithIrregularVerbsSheet(List<TaskUnit> taskUnitList) {
        String irregularVerbsMapMatchedValue = null;
        Iterator<TaskUnit> it = taskUnitList.iterator();
        while (it.hasNext()) {
            TaskUnit taskUnit = it.next();
            String value = taskUnit.getValue();
            if (StringUtils.isNotBlank(value)) {
                irregularVerbsMapMatchedValue = irregularVerbsMap.get(value);
                if (StringUtils.isNotBlank(irregularVerbsMapMatchedValue)) {
                    //匹配上的词放入单独的list
                    logger.debug("[" + taskUnit.getValue() + "] not statisfied with irregular verbs sheet, " + taskUnit.toString());
                    TaskUnit irregularVerbsMapMatchedTaskUnit = new TaskUnit(taskUnit.getExcelRowNumber(), irregularVerbsMap.get(value));
                    irregularVerbsMapMatchedTaskUnitList.add(irregularVerbsMapMatchedTaskUnit);
                    it.remove();
                } else {
                    logger.debug("[" + taskUnit.getValue() + "] statisfied with irregular verbs sheet, " + taskUnit.toString());
                }
            }
        }
        logger.info("irregular verbs map matched: " + irregularVerbsMapMatchedTaskUnitList.size() + ", remained word: " + taskUnitList.size());
        return taskUnitList;
    }

    //对比 自定义规则excel 变换，查询
    public static List<TaskUnit> compareWithCustomizeVerbsSheet(List<TaskUnit> taskUnitList) {
        Iterator<TaskUnit> it = taskUnitList.iterator();
        while (it.hasNext()) {
            TaskUnit taskUnit = it.next();
            String value = taskUnit.getValue();
            Set<String> translatedValuesSet = new HashSet<>();
            for (int j = 0; j < customizeRuleUnitList.size(); j++) {
                CustomizeRuleUnit customizeRuleUnit = customizeRuleUnitList.get(j);
                String suffix = customizeRuleUnit.getSuffix();
                String newSuffix = customizeRuleUnit.getNewSuffix();
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
                logger.debug("[" + taskUnit.getValue() + "] not statisfied with customize verbs sheet, " + taskUnit.toString());
                customizeRuleUnitNotMatchedTaskUnidList.add(taskUnit);
                it.remove();
            } else {
                logger.debug("[" + taskUnit.getValue() + "] statisfied with customize verbs sheet, " + taskUnit.toString());
            }
        }
        return taskUnitList;
    }


    //读YAML配置文件
    public static LinkedHashMap readYAML() {
        InputStream input = Helper.class.getClassLoader().getResourceAsStream("application.yml");
        Yaml yaml = new Yaml();
        return yaml.load(input);
    }

    public static void loadProperties() {
        LinkedHashMap yamlProperties = readYAML();

        VERSION = (String) yamlProperties.get("VERSION");
        logger.info("VERSION: " + VERSION);

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

        //规定时间内没有执行完的任务的颜色
        APP_TIMEOUT_TASK_COLOR_IN_RGB = (Integer) yamlProperties.get("APP_TIMEOUT_TASK_COLOR_IN_RGB");
        logger.info("APP_TIMEOUT_TASK_COLOR_IN_RGB: " + APP_TIMEOUT_TASK_COLOR_IN_RGB);

        //jsoup超时时间（秒）
        JSOUP_TIMEOUT_SECOND = (Integer) yamlProperties.get("JSOUP_TIMEOUT_SECOND");
        logger.info("JSOUP_TIMEOUT_SECOND: " + JSOUP_TIMEOUT_SECOND);

        //读启用线程数相关配置信息
        WORKER_SIZE = (Integer) yamlProperties.get("WORKER_SIZE");
        logger.info("WORKER_SIZE: " + WORKER_SIZE);

        //程序运行超时时间（分钟）
        APP_TIMEOUT_MINUTE = (Integer) yamlProperties.get("APP_TIMEOUT_MINUTE");
        logger.info("APP_TIMEOUT_MINUTE: " + APP_TIMEOUT_MINUTE);

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

        //读待处理任务单词相关配置信息
        LinkedHashMap resultExcelInfo = (LinkedHashMap) yamlProperties.get("RESULT_EXCEL");
        RESULT_EXCEL_NAME = (String) resultExcelInfo.get("EXCEL_NAME");
        logger.info("RESULT_EXCEL_NAME: " + RESULT_EXCEL_NAME);
        RESULT_EXCEL_SHEET_NAME = (String) resultExcelInfo.get("SHEET_NAME");
        logger.info("RESULT_EXCEL_SHEET_NAME: " + RESULT_EXCEL_SHEET_NAME);

    }

    //版本号
    public static String VERSION = "empty";

    //程序运行超时时间（分钟）
    public static int APP_TIMEOUT_MINUTE = 10;

    //结果excel名称
    public static String RESULT_EXCEL_NAME = "";

    //结果excel sheet名称
    public static String RESULT_EXCEL_SHEET_NAME = "";

    public static String RESULT_FILE_PATH = "";

    //待处理任务单词相关配置信息
    public static String TASK_EXCEL_SHEET_NAME = "word";

    public static String TASK_EXCEL_NAME = "word.xlsx";

    public static String CURRENT_TASK_EXCEL_NAME = "word.xlsx";

    //自定义规则相关配置信息
    public static String CUSTOMIZE_RULES_EXCEL_SHEET_NAME = "customizeRules";

    public static String CUSTOMIZE_RULES_EXCEL_NAME = "customizeRules.xlsx";

    //动词不规则变化表相关配置信息
    public static String IRREGULAR_VERBS_EXCEL_SHEET_NAME = "irregularVerbs";

    public static String IRREGULAR_VERBS_EXCEL_NAME = "irregularVerbs.xlsx";

    // 线程数
    public static int WORKER_SIZE = 1;

    //jsoup超时时间（秒）
    public static int JSOUP_TIMEOUT_SECOND = 10;

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

    //规定时间内没有执行完的任务的颜色
    public static int APP_TIMEOUT_TASK_COLOR_IN_RGB = 0x000000;

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

    //规定时间内没有执行完的任务的类型代码
    public static final int APP_TIMEOUT_TASK_TYPE = 5;

    public static Map<String, String> irregularVerbsMap = new HashMap<>();

    public static List<CustomizeRuleUnit> customizeRuleUnitList = new ArrayList<>();

    public static ArrayBlockingQueue<TaskUnit> irregularVerbsMapMatchedTaskUnitList;

    public static ArrayBlockingQueue<TaskUnit> customizeRuleUnitNotMatchedTaskUnidList;

    public static List<TaskUnit> taskUnitList;
    //开始时间
    public static final long startTime = System.currentTimeMillis();

    public static int LOOP_TIMES = 2;
}
