package com.example.aiexcel.util;

import org.springframework.http.HttpStatus;

import java.time.LocalDateTime;
import java.util.Map;

/**
 * 统一响应工具类
 * 提供标准化的API响应格式
 */
public class ApiResponseUtil {

    /**
     * 成功响应
     */
    public static Map<String, Object> success(Object data) {
        return Map.of(
            "success", true,
            "data", data,
            "timestamp", LocalDateTime.now(),
            "error", null
        );
    }

    /**
     * 成功响应（带消息）
     */
    public static Map<String, Object> success(Object data, String message) {
        return Map.of(
            "success", true,
            "data", data,
            "message", message,
            "timestamp", LocalDateTime.now(),
            "error", null
        );
    }

    /**
     * 失败响应
     */
    public static Map<String, Object> error(String errorMessage) {
        return Map.of(
            "success", false,
            "data", null,
            "message", errorMessage,
            "timestamp", LocalDateTime.now(),
            "error", errorMessage
        );
    }

    /**
     * 失败响应（带HTTP状态码）
     */
    public static Map<String, Object> error(String errorMessage, HttpStatus status) {
        return Map.of(
            "success", false,
            "data", null,
            "message", errorMessage,
            "status", status.value(),
            "timestamp", LocalDateTime.now(),
            "error", errorMessage
        );
    }

    /**
     * 失败响应（带详细错误信息）
     */
    public static Map<String, Object> error(String errorMessage, String errorDetails) {
        return Map.of(
            "success", false,
            "data", null,
            "message", errorMessage,
            "errorDetails", errorDetails,
            "timestamp", LocalDateTime.now(),
            "error", errorMessage
        );
    }
}