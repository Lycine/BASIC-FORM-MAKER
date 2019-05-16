package org.jozif;

import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;

import static org.jozif.Helper.*;

public class App {

    private static final Logger logger = LoggerFactory.getLogger(App.class);

    public static void main(String[] args) throws Exception {
        //开始时间
        final long startTime = System.currentTimeMillis();

        // 准备阶段

        //读配置文件
        Helper.loadProperties();

        //读不规则变化表excel入内存
        irregularVerbsMap = loadIrregularVerbsSheet();

        //读取自定义规则excel入内存
        customizeRuleUnitList = loadCustomizeRulesSheet();

        //读取待处理单词excel入内存
        List<TaskUnit> taskUnitList = loadTaskSheet();

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
                        + "], [eta: " + etaTime
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


        //符合自定义规则的词写入workbook
        Workbook wb = Helper.taskUnitQueueWriteExcel(resultTaskUnitQueue, getWorkbookFormExcelByPath(TASK_EXCEL_NAME), TASK_EXCEL_SHEET_NAME, CUSTOMIZE_RULE_MATCHED_TASK_TYPE);

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

        //增加图例
        wb = polishWorkBookBeforeGenerate(wb);

        //写入excel
        Helper.workBookWriteToFile(wb, RESULT_EXCEL_NAME);

        logger.info("result excel generated successfully! used time: " + Helper.timeAdapter((System.currentTimeMillis() - startTime) / 1000));
        System.exit(0);
    }
}
