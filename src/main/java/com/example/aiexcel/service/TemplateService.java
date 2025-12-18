package com.example.aiexcel.service;

import com.example.aiexcel.model.Template;
import com.example.aiexcel.repository.TemplateRepository;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.time.LocalDateTime;
import java.util.List;
import java.util.Optional;

/**
 * 模板服务
 * 提供模板的创建、管理、应用等功能
 */
@Service
public class TemplateService {

    @Autowired
    private TemplateRepository templateRepository;

    private static final Logger logger = LoggerFactory.getLogger(TemplateService.class);

    /**
     * 创建新模板
     */
    public Template createTemplate(String name, String description, String command, String category, String userId, Boolean isPublic, String tags) {
        try {
            // 检查是否已存在同名模板
            if (templateRepository.existsByUserIdAndName(userId, name)) {
                logger.warn("Template with name '{}' already exists for user '{}'", name, userId);
                throw new IllegalArgumentException("Template with name '" + name + "' already exists");
            }

            Template template = new Template(name, description, command, category, userId, isPublic);
            template.setTags(tags);
            template = templateRepository.save(template);

            logger.info("Created template '{}' for user '{}'", name, userId);
            return template;
        } catch (Exception e) {
            logger.error("Error creating template for user '{}': {}", userId, e.getMessage(), e);
            throw e;
        }
    }

    /**
     * 获取用户的所有模板
     */
    public List<Template> getUserTemplates(String userId) {
        try {
            List<Template> templates = templateRepository.findByUserId(userId);
            logger.debug("Retrieved {} templates for user '{}'", templates.size(), userId);
            return templates;
        } catch (Exception e) {
            logger.error("Error retrieving templates for user '{}': {}", userId, e.getMessage(), e);
            return List.of();
        }
    }

    /**
     * 获取公开模板
     */
    public List<Template> getPublicTemplates() {
        try {
            List<Template> templates = templateRepository.findByIsPublicTrue();
            logger.debug("Retrieved {} public templates", templates.size());
            return templates;
        } catch (Exception e) {
            logger.error("Error retrieving public templates: {}", e.getMessage(), e);
            return List.of();
        }
    }

    /**
     * 获取热门模板
     */
    public List<Template> getPopularTemplates() {
        try {
            List<Template> templates = templateRepository.findTop10ByIsPublicTrueOrderByCreatedAtDesc();
            logger.debug("Retrieved {} popular templates", templates.size());
            return templates;
        } catch (Exception e) {
            logger.error("Error retrieving popular templates: {}", e.getMessage(), e);
            return List.of();
        }
    }

    /**
     * 根据ID获取模板
     */
    public Optional<Template> getTemplateById(Long id) {
        try {
            Optional<Template> template = templateRepository.findById(id);
            if (template.isPresent()) {
                logger.debug("Retrieved template with ID: {}", id);
            } else {
                logger.debug("Template with ID {} not found", id);
            }
            return template;
        } catch (Exception e) {
            logger.error("Error retrieving template with ID {}: {}", id, e.getMessage(), e);
            return Optional.empty();
        }
    }

    /**
     * 更新模板
     */
    public Template updateTemplate(Long id, String name, String description, String command, String category, String userId, Boolean isPublic, String tags) {
        try {
            Optional<Template> existingTemplateOpt = templateRepository.findById(id);
            if (existingTemplateOpt.isEmpty()) {
                logger.error("Template with ID {} not found for update", id);
                return null;
            }

            Template existingTemplate = existingTemplateOpt.get();
            
            // 检查是否是模板所有者
            if (!existingTemplate.getUserId().equals(userId)) {
                logger.error("User '{}' does not have permission to update template with ID {}", userId, id);
                return null;
            }

            // 检查是否已存在同名模板（排除当前模板）
            if (!existingTemplate.getName().equals(name) && 
                templateRepository.existsByUserIdAndName(userId, name)) {
                logger.warn("Template with name '{}' already exists for user '{}'", name, userId);
                return null;
            }

            existingTemplate.setName(name);
            existingTemplate.setDescription(description);
            existingTemplate.setCommand(command);
            existingTemplate.setCategory(category);
            existingTemplate.setIsPublic(isPublic != null ? isPublic : false);
            existingTemplate.setTags(tags);
            existingTemplate.setUpdatedAt(LocalDateTime.now());

            Template updatedTemplate = templateRepository.save(existingTemplate);
            logger.info("Updated template with ID: {}", id);
            return updatedTemplate;
        } catch (Exception e) {
            logger.error("Error updating template with ID {}: {}", id, e.getMessage(), e);
            return null;
        }
    }

    /**
     * 删除模板
     */
    public boolean deleteTemplate(Long id, String userId) {
        try {
            Optional<Template> templateOpt = templateRepository.findById(id);
            if (templateOpt.isEmpty()) {
                logger.error("Template with ID {} not found for deletion", id);
                return false;
            }

            Template template = templateOpt.get();
            
            // 检查是否是模板所有者
            if (!template.getUserId().equals(userId)) {
                logger.error("User '{}' does not have permission to delete template with ID {}", userId, id);
                return false;
            }

            templateRepository.delete(template);
            logger.info("Deleted template with ID: {}", id);
            return true;
        } catch (Exception e) {
            logger.error("Error deleting template with ID {}: {}", id, e.getMessage(), e);
            return false;
        }
    }

    /**
     * 根据分类获取模板
     */
    public List<Template> getTemplatesByCategory(String category) {
        try {
            List<Template> templates = templateRepository.findByCategory(category);
            logger.debug("Retrieved {} templates in category '{}'", templates.size(), category);
            return templates;
        } catch (Exception e) {
            logger.error("Error retrieving templates in category '{}': {}", category, e.getMessage(), e);
            return List.of();
        }
    }

    /**
     * 根据用户ID和分类获取模板
     */
    public List<Template> getUserTemplatesByCategory(String userId, String category) {
        try {
            List<Template> templates = templateRepository.findByUserIdAndCategory(userId, category);
            logger.debug("Retrieved {} templates for user '{}' in category '{}'", templates.size(), userId, category);
            return templates;
        } catch (Exception e) {
            logger.error("Error retrieving templates for user '{}' in category '{}': {}", userId, category, e.getMessage(), e);
            return List.of();
        }
    }

    /**
     * 搜索模板（按名称）
     */
    public List<Template> searchTemplates(String searchTerm) {
        try {
            List<Template> templates = templateRepository.findByNameContainingIgnoreCase(searchTerm);
            logger.debug("Found {} templates matching search term '{}'", templates.size(), searchTerm);
            return templates;
        } catch (Exception e) {
            logger.error("Error searching templates with term '{}': {}", searchTerm, e.getMessage(), e);
            return List.of();
        }
    }

    /**
     * 搜索用户模板（按名称）
     */
    public List<Template> searchUserTemplates(String userId, String searchTerm) {
        try {
            List<Template> templates = templateRepository.findByUserIdAndNameContainingIgnoreCase(userId, searchTerm);
            logger.debug("Found {} templates for user '{}' matching search term '{}'", templates.size(), userId, searchTerm);
            return templates;
        } catch (Exception e) {
            logger.error("Error searching templates for user '{}' with term '{}': {}", userId, searchTerm, e.getMessage(), e);
            return List.of();
        }
    }

    /**
     * 应用模板到命令
     * 这个方法将返回模板的命令内容，供其他服务使用
     */
    public String applyTemplate(Long templateId) {
        try {
            Optional<Template> templateOpt = templateRepository.findById(templateId);
            if (templateOpt.isEmpty()) {
                logger.error("Template with ID {} not found for application", templateId);
                return null;
            }

            String command = templateOpt.get().getCommand();
            logger.debug("Applied template with ID {}, returning command", templateId);
            return command;
        } catch (Exception e) {
            logger.error("Error applying template with ID {}: {}", templateId, e.getMessage(), e);
            return null;
        }
    }
}