package com.example.aiexcel.controller;

import com.example.aiexcel.service.AiSuggestionService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.Map;

/**
 * AI智能建议控制器
 * 提供基于AI的数据分析和建议功能
 */
@RestController
@RequestMapping("/api/ai/suggestions")
public class AiSuggestionController {

    @Autowired
    private AiSuggestionService aiSuggestionService;

    private static final Logger logger = LoggerFactory.getLogger(AiSuggestionController.class);

    /**
     * 获取AI智能建议
     */
    @PostMapping("/general")
    public ResponseEntity<Map<String, Object>> getGeneralSuggestions(
            @RequestParam("file") MultipartFile file,
            @RequestParam(value = "context", required = false) String context) {
        logger.info("Received request for general AI suggestions for file: {}", file.getOriginalFilename());

        try {
            // 验证文件
            if (file.isEmpty()) {
                logger.error("File is empty");
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "File is empty"
                );
                return ResponseEntity.badRequest().body(response);
            }

            // 获取建议
            Map<String, Object> suggestions = aiSuggestionService.getSuggestions(file, context);

            logger.info("Successfully generated general AI suggestions for file: {}", file.getOriginalFilename());
            return ResponseEntity.ok(suggestions);
        } catch (Exception e) {
            logger.error("Error generating general AI suggestions for file: {}", file.getOriginalFilename(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error generating AI suggestions: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 获取数据类型分析和建议
     */
    @PostMapping("/data-types")
    public ResponseEntity<Map<String, Object>> getDataTypeSuggestions(@RequestParam("file") MultipartFile file) {
        logger.info("Received request for data type AI suggestions for file: {}", file.getOriginalFilename());

        try {
            // 验证文件
            if (file.isEmpty()) {
                logger.error("File is empty");
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "File is empty"
                );
                return ResponseEntity.badRequest().body(response);
            }

            // 获取数据类型建议
            Map<String, Object> suggestions = aiSuggestionService.analyzeDataTypesAndSuggest(file);

            logger.info("Successfully generated data type AI suggestions for file: {}", file.getOriginalFilename());
            return ResponseEntity.ok(suggestions);
        } catch (Exception e) {
            logger.error("Error generating data type AI suggestions for file: {}", file.getOriginalFilename(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error generating data type suggestions: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 获取格式化建议
     */
    @PostMapping("/formatting")
    public ResponseEntity<Map<String, Object>> getFormattingSuggestions(@RequestParam("file") MultipartFile file) {
        logger.info("Received request for formatting AI suggestions for file: {}", file.getOriginalFilename());

        try {
            // 验证文件
            if (file.isEmpty()) {
                logger.error("File is empty");
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "File is empty"
                );
                return ResponseEntity.badRequest().body(response);
            }

            // 获取格式化建议
            Map<String, Object> suggestions = aiSuggestionService.getFormattingSuggestions(file);

            logger.info("Successfully generated formatting AI suggestions for file: {}", file.getOriginalFilename());
            return ResponseEntity.ok(suggestions);
        } catch (Exception e) {
            logger.error("Error generating formatting AI suggestions for file: {}", file.getOriginalFilename(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error generating formatting suggestions: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 获取性能优化建议
     */
    @PostMapping("/performance")
    public ResponseEntity<Map<String, Object>> getPerformanceSuggestions(@RequestParam("file") MultipartFile file) {
        logger.info("Received request for performance AI suggestions for file: {}", file.getOriginalFilename());

        try {
            // 验证文件
            if (file.isEmpty()) {
                logger.error("File is empty");
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "File is empty"
                );
                return ResponseEntity.badRequest().body(response);
            }

            // 获取性能建议
            Map<String, Object> suggestions = aiSuggestionService.getPerformanceSuggestions(file);

            logger.info("Successfully generated performance AI suggestions for file: {}", file.getOriginalFilename());
            return ResponseEntity.ok(suggestions);
        } catch (Exception e) {
            logger.error("Error generating performance AI suggestions for file: {}", file.getOriginalFilename(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error generating performance suggestions: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 综合智能分析
     */
    @PostMapping("/comprehensive")
    public ResponseEntity<Map<String, Object>> getComprehensiveAnalysis(@RequestParam("file") MultipartFile file) {
        logger.info("Received request for comprehensive AI analysis for file: {}", file.getOriginalFilename());

        try {
            // 验证文件
            if (file.isEmpty()) {
                logger.error("File is empty");
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "File is empty"
                );
                return ResponseEntity.badRequest().body(response);
            }

            // 获取所有类型的建议
            Map<String, Object> dataTypeSuggestions = aiSuggestionService.analyzeDataTypesAndSuggest(file);
            Map<String, Object> formattingSuggestions = aiSuggestionService.getFormattingSuggestions(file);
            Map<String, Object> performanceSuggestions = aiSuggestionService.getPerformanceSuggestions(file);

            // 组合结果
            Map<String, Object> result = Map.of(
                "success", true,
                "dataTypeAnalysis", dataTypeSuggestions,
                "formattingSuggestions", formattingSuggestions,
                "performanceSuggestions", performanceSuggestions
            );

            logger.info("Successfully generated comprehensive AI analysis for file: {}", file.getOriginalFilename());
            return ResponseEntity.ok(result);
        } catch (Exception e) {
            logger.error("Error generating comprehensive AI analysis for file: {}", file.getOriginalFilename(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error generating comprehensive analysis: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }
}