# BASIC FORM MAKER

为了将不规则形式的英语单词通过一系列规则变为原形的小程序

## Requirements

- JAVA 8

## TODO

- 增加配置项来决定筛选程序执行几遍;

## Changelog

- 0.0.9
    1. 优化代码
    
- 0.0.8
    1. 增加日志 解决有些设备中translatedValuesSet的不是每个词都执行;
    
- 0.0.7 
    1. 增加readme;
    1. 增加配置项 配置本程序运行超时时间，到超时时间将已完成的任务写入结果excel，避免遇到死锁或其他未知问题导致运行半天没有输出;

- 0.0.6 
    1. 增加windows下便于使用的批处理脚本;

- 0.0.5 
    1. 增加配置项 配置结果excel名和sheet名;
    1. 结果excel增加图例sheet; 

- 0.0.4 
    1. 增加配置项 配置jsoup超时时间;
    1. 优化读取配置文件相关代码;

- 0.0.3 
    1. 增加日志 友好的看到当前运行状态（用时，剩余时间，剩余任务等）;

- 0.0.2 
    1. 增加日志 进一步查看在其他环境中，自定义规则执行过程中有些执行，有些没有执行;

- 0.0.1 
    1. init;