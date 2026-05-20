package com.example.aiexcel.config;

import ch.qos.logback.classic.Level;
import ch.qos.logback.classic.spi.ILoggingEvent;
import ch.qos.logback.core.pattern.CompositeConverter;

import java.util.Map;

/**
 * 自定义日志级别颜色转换器
 * ERROR=红, WARN=黄, INFO=蓝, DEBUG=白, TRACE=紫
 */
public class ColoredLevelConverter extends CompositeConverter<ILoggingEvent> {

    private static final Map<Integer, String> LEVEL_COLORS = Map.of(
            Level.ERROR_INTEGER, "\033[31m",   // 红色
            Level.WARN_INTEGER, "\033[33m",    // 黄色
            Level.INFO_INTEGER, "\033[34m",    // 蓝色
            Level.DEBUG_INTEGER, "\033[37m",   // 白色
            Level.TRACE_INTEGER, "\033[35m"    // 紫色
    );

    private static final String RESET = "\033[0m";

    @Override
    protected String transform(ILoggingEvent event, String in) {
        String color = LEVEL_COLORS.getOrDefault(event.getLevel().toInteger(), "\033[37m");
        return color + in + RESET;
    }
}
