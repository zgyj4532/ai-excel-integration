package com.example.aiexcel.controller;

import com.example.aiexcel.model.Template;
import com.example.aiexcel.service.TemplateService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.List;
import java.util.Map;
import java.util.Optional;

/**
 * 模板控制器
 * 提供模板管理功能的API
 */
@RestController
@RequestMapping("/api/templates")
public class TemplateController {

    @Autowired
    private TemplateService templateService;

    private static final Logger logger = LoggerFactory.getLogger(TemplateController.class);

    /**
     * 创建新模板
     */
    @PostMapping("/create")
    public ResponseEntity<Map<String, Object>> createTemplate(@RequestBody Map<String, Object> requestBody) {
        logger.info("Received request to create template");

        try {
            // 验证必需参数
            String name = (String) requestBody.get("name");
            String command = (String) requestBody.get("command");
            String userId = (String) requestBody.get("userId");

            if (name == null || name.trim().isEmpty()) {
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "Template name is required"
                );
                return ResponseEntity.badRequest().body(response);
            }

            if (command == null || command.trim().isEmpty()) {
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "Template command is required"
                );
                return ResponseEntity.badRequest().body(response);
            }

            if (userId == null || userId.trim().isEmpty()) {
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "User ID is required"
                );
                return ResponseEntity.badRequest().body(response);
            }

            // 获取可选参数
            String description = (String) requestBody.get("description");
            String category = (String) requestBody.get("category");
            Boolean isPublic = (Boolean) requestBody.get("isPublic");
            String tags = (String) requestBody.get("tags");

            // 创建模板
            Template template = templateService.createTemplate(name, description, command, category, userId, isPublic, tags);

            Map<String, Object> response = Map.of(
                "success", true,
                "data", template,
                "message", "Template created successfully"
            );

            logger.info("Successfully created template '{}'", name);
            return ResponseEntity.ok(response);
        } catch (IllegalArgumentException e) {
            logger.error("Invalid template creation request: {}", e.getMessage());
            Map<String, Object> response = Map.of(
                "success", false,
                "error", e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        } catch (Exception e) {
            logger.error("Error creating template: {}", e.getMessage(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error creating template: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 获取用户的所有模板
     */
    @GetMapping("/user/{userId}")
    public ResponseEntity<Map<String, Object>> getUserTemplates(@PathVariable String userId) {
        logger.info("Received request to get templates for user: {}", userId);

        try {
            List<Template> templates = templateService.getUserTemplates(userId);

            Map<String, Object> response = Map.of(
                "success", true,
                "data", templates,
                "count", templates.size()
            );

            logger.info("Successfully returned {} templates for user: {}", templates.size(), userId);
            return ResponseEntity.ok(response);
        } catch (Exception e) {
            logger.error("Error retrieving templates for user '{}': {}", userId, e.getMessage(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error retrieving templates: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 获取公开模板
     */
    @GetMapping("/public")
    public ResponseEntity<Map<String, Object>> getPublicTemplates() {
        logger.info("Received request to get public templates");

        try {
            List<Template> templates = templateService.getPublicTemplates();

            Map<String, Object> response = Map.of(
                "success", true,
                "data", templates,
                "count", templates.size()
            );

            logger.info("Successfully returned {} public templates", templates.size());
            return ResponseEntity.ok(response);
        } catch (Exception e) {
            logger.error("Error retrieving public templates: {}", e.getMessage(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error retrieving public templates: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 获取热门模板
     */
    @GetMapping("/popular")
    public ResponseEntity<Map<String, Object>> getPopularTemplates() {
        logger.info("Received request to get popular templates");

        try {
            List<Template> templates = templateService.getPopularTemplates();

            Map<String, Object> response = Map.of(
                "success", true,
                "data", templates,
                "count", templates.size()
            );

            logger.info("Successfully returned {} popular templates", templates.size());
            return ResponseEntity.ok(response);
        } catch (Exception e) {
            logger.error("Error retrieving popular templates: {}", e.getMessage(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error retrieving popular templates: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 获取单个模板详情
     */
    @GetMapping("/{id}")
    public ResponseEntity<Map<String, Object>> getTemplate(@PathVariable Long id) {
        logger.info("Received request to get template with ID: {}", id);

        try {
            Optional<Template> templateOpt = templateService.getTemplateById(id);

            if (templateOpt.isPresent()) {
                Map<String, Object> response = Map.of(
                    "success", true,
                    "data", templateOpt.get()
                );
                logger.info("Successfully returned template with ID: {}", id);
                return ResponseEntity.ok(response);
            } else {
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "Template not found"
                );
                logger.warn("Template with ID {} not found", id);
                return ResponseEntity.badRequest().body(response);
            }
        } catch (Exception e) {
            logger.error("Error retrieving template with ID {}: {}", id, e.getMessage(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error retrieving template: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 更新模板
     */
    @PutMapping("/{id}")
    public ResponseEntity<Map<String, Object>> updateTemplate(
            @PathVariable Long id,
            @RequestBody Map<String, Object> requestBody) {
        logger.info("Received request to update template with ID: {}", id);

        try {
            // 验证必需参数
            String name = (String) requestBody.get("name");
            String command = (String) requestBody.get("command");
            String userId = (String) requestBody.get("userId");

            if (name == null || name.trim().isEmpty()) {
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "Template name is required"
                );
                return ResponseEntity.badRequest().body(response);
            }

            if (command == null || command.trim().isEmpty()) {
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "Template command is required"
                );
                return ResponseEntity.badRequest().body(response);
            }

            if (userId == null || userId.trim().isEmpty()) {
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "User ID is required"
                );
                return ResponseEntity.badRequest().body(response);
            }

            // 获取可选参数
            String description = (String) requestBody.get("description");
            String category = (String) requestBody.get("category");
            Boolean isPublic = (Boolean) requestBody.get("isPublic");
            String tags = (String) requestBody.get("tags");

            // 更新模板
            Template template = templateService.updateTemplate(id, name, description, command, category, userId, isPublic, tags);

            if (template != null) {
                Map<String, Object> response = Map.of(
                    "success", true,
                    "data", template,
                    "message", "Template updated successfully"
                );
                logger.info("Successfully updated template with ID: {}", id);
                return ResponseEntity.ok(response);
            } else {
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "Template not found or unauthorized"
                );
                logger.warn("Failed to update template with ID: {} (not found or unauthorized)", id);
                return ResponseEntity.badRequest().body(response);
            }
        } catch (Exception e) {
            logger.error("Error updating template with ID {}: {}", id, e.getMessage(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error updating template: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 删除模板
     */
    @DeleteMapping("/{id}")
    public ResponseEntity<Map<String, Object>> deleteTemplate(@PathVariable Long id, @RequestParam String userId) {
        logger.info("Received request to delete template with ID: {} for user: {}", id, userId);

        try {
            boolean success = templateService.deleteTemplate(id, userId);

            if (success) {
                Map<String, Object> response = Map.of(
                    "success", true,
                    "message", "Template deleted successfully"
                );
                logger.info("Successfully deleted template with ID: {}", id);
                return ResponseEntity.ok(response);
            } else {
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "Template not found or unauthorized"
                );
                logger.warn("Failed to delete template with ID: {} (not found or unauthorized)", id);
                return ResponseEntity.badRequest().body(response);
            }
        } catch (Exception e) {
            logger.error("Error deleting template with ID {}: {}", id, e.getMessage(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error deleting template: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 根据分类获取模板
     */
    @GetMapping("/category/{category}")
    public ResponseEntity<Map<String, Object>> getTemplatesByCategory(@PathVariable String category) {
        logger.info("Received request to get templates in category: {}", category);

        try {
            List<Template> templates = templateService.getTemplatesByCategory(category);

            Map<String, Object> response = Map.of(
                "success", true,
                "data", templates,
                "count", templates.size()
            );

            logger.info("Successfully returned {} templates in category: {}", templates.size(), category);
            return ResponseEntity.ok(response);
        } catch (Exception e) {
            logger.error("Error retrieving templates in category '{}': {}", category, e.getMessage(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error retrieving templates: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 搜索模板
     */
    @GetMapping("/search")
    public ResponseEntity<Map<String, Object>> searchTemplates(@RequestParam String query) {
        logger.info("Received request to search templates with query: {}", query);

        try {
            List<Template> templates = templateService.searchTemplates(query);

            Map<String, Object> response = Map.of(
                "success", true,
                "data", templates,
                "count", templates.size()
            );

            logger.info("Successfully found {} templates matching query: {}", templates.size(), query);
            return ResponseEntity.ok(response);
        } catch (Exception e) {
            logger.error("Error searching templates with query '{}': {}", query, e.getMessage(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error searching templates: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 搜索用户模板
     */
    @GetMapping("/search/user/{userId}")
    public ResponseEntity<Map<String, Object>> searchUserTemplates(
            @PathVariable String userId,
            @RequestParam String query) {
        logger.info("Received request to search templates for user '{}' with query: {}", userId, query);

        try {
            List<Template> templates = templateService.searchUserTemplates(userId, query);

            Map<String, Object> response = Map.of(
                "success", true,
                "data", templates,
                "count", templates.size()
            );

            logger.info("Successfully found {} templates for user '{}' matching query: {}", templates.size(), userId, query);
            return ResponseEntity.ok(response);
        } catch (Exception e) {
            logger.error("Error searching templates for user '{}' with query '{}': {}", userId, query, e.getMessage(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error searching templates: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 应用模板
     */
    @PostMapping("/apply/{id}")
    public ResponseEntity<Map<String, Object>> applyTemplate(@PathVariable Long id) {
        logger.info("Received request to apply template with ID: {}", id);

        try {
            String command = templateService.applyTemplate(id);

            if (command != null) {
                Map<String, Object> response = Map.of(
                    "success", true,
                    "command", command,
                    "message", "Template applied successfully"
                );
                logger.info("Successfully applied template with ID: {}", id);
                return ResponseEntity.ok(response);
            } else {
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "Template not found"
                );
                logger.warn("Template with ID {} not found for application", id);
                return ResponseEntity.badRequest().body(response);
            }
        } catch (Exception e) {
            logger.error("Error applying template with ID {}: {}", id, e.getMessage(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error applying template: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }
}