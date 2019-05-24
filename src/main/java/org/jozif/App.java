package org.jozif;

import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.util.*;

import static org.jozif.Helper.*;

public class App {

    private static final Logger logger = LoggerFactory.getLogger(App.class);

    public static void main(String[] args) throws Exception {

        // 准备阶段

        //读配置文件
        Helper.loadProperties();

        //读不规则变化表excel入内存
        irregularVerbsMap = loadIrregularVerbsSheet();

        //读取自定义规则excel入内存
        customizeRuleUnitList = loadCustomizeRulesSheet();

        //所有待处理任务所在的绝对路径
        String filePath = "/home/user/Documents/pending";
        logger.info("filePath: " + filePath);
        Set<File> pendingSet = readFile(filePath);
        pendingSet = getPendingTaskFile(pendingSet);
        int pendingFileTotalSize = pendingSet.size();
        Iterator<File> it = pendingSet.iterator();
        while (it.hasNext()) {
            File file = it.next();
            taskUnitList = loadTaskSheet(file.getAbsolutePath(), TASK_EXCEL_SHEET_NAME);
            CURRENT_TASK_EXCEL_NAME = file.getAbsolutePath();
            //循环多次
            Workbook wb = null;
            for (int i = 0; i < LOOP_TIMES; i++) {
                logger.info("current file: " + file.getName()
                        + ", start loop: " + (i + 1)
                        + ", taskUnitListSize: " + taskUnitList.size());
                //读取待处理单词excel入内存
                wb = process();
                taskUnitList = loadTaskSheet(wb.getSheetAt(0), 1);

                Set result = new HashSet(taskUnitList);
                taskUnitList = new ArrayList<>(result);

                //增加图例
                wb = polishWorkBookBeforeGenerate(wb);

                //写入excel
                Helper.workBookWriteToFile(wb, file.getName() + "_LOOP_" + (i + 1) + "_RESULT_LENGTH_" + taskUnitList.size() + "_" + RESULT_EXCEL_NAME);
                
            }
            if (null == wb) {
                logger.error("wb is null, exit");
                System.exit(0);
            }
            it.remove();
            long usedSeond = (System.currentTimeMillis() - startTime) / 1000;
            int finishedFile = pendingFileTotalSize - pendingSet.size();
            double timePerFile = 1.0 * usedSeond / finishedFile;
            double etaSecond = pendingSet.size() * timePerFile;
            String etaTime = Helper.timeAdapter(new Double(etaSecond).longValue());
            Double progressRate = 1.0 * finishedFile / pendingFileTotalSize;
            logger.info("[usedTime: "
                    + Helper.timeAdapter(usedSeond)
                    + "], [finishedFile/totalFile: " + finishedFile + "/" + pendingFileTotalSize
                    + "], [total eta: " + etaTime
                    + "], [progressRate: " + String.format("%.2f", progressRate * 100)
                    + "%]");
        }

        logger.info("result excel generated successfully! used time: " + Helper.timeAdapter((System.currentTimeMillis() - startTime) / 1000));
    }
}
