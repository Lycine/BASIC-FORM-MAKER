package org.jozif;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.formula.functions.T;
import org.jsoup.HttpStatusException;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.net.SocketTimeoutException;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.concurrent.ArrayBlockingQueue;

public class Worker implements Runnable {

    private static final Logger logger = LoggerFactory.getLogger(App.class);

    private ArrayBlockingQueue<TaskUnit> taskUnitQueue;

    private ArrayBlockingQueue<TaskUnit> resultTaskUnitQueue;

    private ArrayBlockingQueue<TaskUnit> failureTaskUnitQueue;

    private long startConcurrencyTime;

    private int taskUnitQueueSize;

    public Worker(ArrayBlockingQueue<TaskUnit> taskUnitQueue, ArrayBlockingQueue<TaskUnit> resultTaskUnitQueue, ArrayBlockingQueue<TaskUnit> failureTaskUnitQueue, long startConcurrencyTime, int taskUnitQueueSize) {
        this.taskUnitQueue = taskUnitQueue;
        this.resultTaskUnitQueue = resultTaskUnitQueue;
        this.failureTaskUnitQueue = failureTaskUnitQueue;
        this.startConcurrencyTime = startConcurrencyTime;
        this.taskUnitQueueSize = taskUnitQueueSize;
    }

    @Override
    public void run() {
        logger.info("start working");
        while (taskUnitQueue.size() > 0) {
            logger.info("remained taskUnitQueue size : " + taskUnitQueue.size());
            //fetch task
            TaskUnit taskUnit = taskUnitQueue.poll();
            if (null == taskUnit) {
                logger.info("task is null, taskUnitQueue size: " + taskUnitQueue.size());
                continue;
            }
            Set<String> translatedValuesSet = taskUnit.getTranslatedValuesSet();
            Set<String> resultValuesSet = new HashSet<>();
            for (String translatedValue : translatedValuesSet) {
                //find in online dictionary
                String url = "https://www.merriam-webster.com/dictionary/" + translatedValue;
                logger.info("request url: " + url);
                String ua = "User-Agent: Mozilla/5.0 (iPhone; CPU iPhone OS 11_0 like Mac OS X) AppleWebKit/604.1.38 (KHTML, like Gecko) Version/11.0 Mobile/15A372 Safari/604.1";
                try {
                    Document doc = Jsoup.connect(url)
                            .userAgent(ua)
                            .timeout(10 * 1000)
                            .get();
                    if (doc.text().contains("mispelled-word")) {
                        //word not exist
                        logger.debug("[" + translatedValue + "] not find in webster dict");
                    } else {
                        //word exist, add translatedValue into result list
                        resultValuesSet.add(translatedValue);
                        logger.debug("[" + translatedValue + "] find in webster dict, add it into result set. taskUnit:" + taskUnit.toString());
                    }
                } catch (HttpStatusException hse) {
                    //404错误，字典不存在该单词
                    if (404 == hse.getStatusCode()) {
                        logger.info("[" + translatedValue + "] not find in webster dict, status code 404");
                    } else {
                        hse.printStackTrace();
                    }
                } catch (SocketTimeoutException ste) {
                    //超时错误
                    if (StringUtils.equals("Read timed out", ste.getMessage())) {
                        failureTaskUnitQueue.add(taskUnit);
                        logger.error("[" + translatedValue + "] request timeout, add to failure.xlsx, request url " + url);
                    }
                } catch (IOException ioe) {
                    ioe.printStackTrace();
                }
            }
            taskUnit.setResultValuesSet(resultValuesSet);
            if (resultValuesSet.size() > 1) {
                String logMessage = "task have multiple result value: " + taskUnit.toString();

                //reverse customize rule to find basic form
//                Iterator<String> it = resultValuesSet.iterator();
//                while (it.hasNext()) {
//                    String value = it.next();
//                    for (CustomizeRuleUnit customizeRuleUnit : customizeRuleUnitList) {
//                        String newSuffix = customizeRuleUnit.getNewSuffix();
//                        if (value.endsWith(newSuffix)) {
//                            it.remove();
//                            logger.debug("[" + value + "], matched new suffix, redundant value, remove");
//                            break;
//                        }
//                    }
//                }

//                logger.warn(logMessage + ", the final result is: " + resultValuesSet.toString());
                logger.warn(logMessage);
            }
            resultTaskUnitQueue.add(taskUnit);
            final long endTime = System.currentTimeMillis();
            long usedSeond = (endTime - startConcurrencyTime) / 1000;
            int finishedTask = taskUnitQueueSize - taskUnitQueue.size();
            double timePerTask = 1.0 * usedSeond / finishedTask;
            double etaSecond = taskUnitQueue.size() * timePerTask;
            String etaTime = Helper.timeAdapter(new Double(etaSecond).longValue());
            Double progressRate = 1.0 * finishedTask / taskUnitQueueSize;
            logger.info("[usedTime: " + Helper.timeAdapter(usedSeond) + "], [finished/all: " + finishedTask + "/" + taskUnitQueueSize + "], [eta: " + etaTime + "], [progressRate: " + String.format("%.2f", progressRate * 100) + "%]");

        }
        logger.info("finished!");
    }
}
