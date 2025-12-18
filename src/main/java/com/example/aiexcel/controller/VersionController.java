package com.example.aiexcel.controller;

import com.example.aiexcel.model.FileVersion;
import com.example.aiexcel.service.VersionService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;
import java.util.Map;

/**
 * 文件版本控制器
 * 提供文件版本管理功能的API
 */
@RestController
@RequestMapping("/api/versions")
public class VersionController {

    @Autowired
    private VersionService versionService;

    private static final Logger logger = LoggerFactory.getLogger(VersionController.class);

    /**
     * 创建新版本
     */
    @PostMapping("/create")
    public ResponseEntity<Map<String, Object>> createVersion(
            @RequestParam("file") MultipartFile file,
            @RequestParam("fileId") String fileId,
            @RequestParam(value = "changeDescription", required = false) String changeDescription,
            @RequestParam(value = "userId", required = false) String userId) {
        
        logger.info("Received request to create version for file: {} (ID: {})", file.getOriginalFilename(), fileId);

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

            // 创建版本
            FileVersion version = versionService.createVersion(fileId, file, changeDescription, userId);

            if (version != null) {
                Map<String, Object> response = Map.of(
                    "success", true,
                    "data", version,
                    "message", "Version created successfully"
                );
                logger.info("Successfully created version {} for file: {}", version.getVersionNumber(), fileId);
                return ResponseEntity.ok(response);
            } else {
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "Failed to create version"
                );
                logger.error("Failed to create version for file: {}", fileId);
                return ResponseEntity.badRequest().body(response);
            }
        } catch (Exception e) {
            logger.error("Error creating version for file: {}", fileId, e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error creating version: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 获取文件的所有版本
     */
    @GetMapping("/file/{fileId}")
    public ResponseEntity<Map<String, Object>> getFileVersions(@PathVariable String fileId) {
        logger.info("Received request to get versions for file: {}", fileId);

        try {
            List<FileVersion> versions = versionService.getFileVersions(fileId);

            Map<String, Object> response = Map.of(
                "success", true,
                "data", versions,
                "count", versions.size()
            );

            logger.info("Successfully returned {} versions for file: {}", versions.size(), fileId);
            return ResponseEntity.ok(response);
        } catch (Exception e) {
            logger.error("Error retrieving versions for file: {}", fileId, e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error retrieving versions: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 获取特定版本
     */
    @GetMapping("/{fileId}/{versionNumber}")
    public ResponseEntity<Map<String, Object>> getVersion(
            @PathVariable String fileId,
            @PathVariable Integer versionNumber) {
        logger.info("Received request to get version {} for file: {}", versionNumber, fileId);

        try {
            FileVersion version = versionService.getVersion(fileId, versionNumber);

            if (version != null) {
                Map<String, Object> response = Map.of(
                    "success", true,
                    "data", version
                );
                logger.info("Successfully returned version {} for file: {}", versionNumber, fileId);
                return ResponseEntity.ok(response);
            } else {
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "Version not found"
                );
                logger.warn("Version {} not found for file: {}", versionNumber, fileId);
                return ResponseEntity.badRequest().body(response);
            }
        } catch (Exception e) {
            logger.error("Error retrieving version {} for file: {}", versionNumber, fileId, e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error retrieving version: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 获取最新版本
     */
    @GetMapping("/latest/{fileId}")
    public ResponseEntity<Map<String, Object>> getLatestVersion(@PathVariable String fileId) {
        logger.info("Received request to get latest version for file: {}", fileId);

        try {
            FileVersion version = versionService.getLatestVersion(fileId);

            if (version != null) {
                Map<String, Object> response = Map.of(
                    "success", true,
                    "data", version
                );
                logger.info("Successfully returned latest version for file: {}", fileId);
                return ResponseEntity.ok(response);
            } else {
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "No versions found"
                );
                logger.warn("No versions found for file: {}", fileId);
                return ResponseEntity.badRequest().body(response);
            }
        } catch (Exception e) {
            logger.error("Error retrieving latest version for file: {}", fileId, e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error retrieving latest version: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 恢复到特定版本
     */
    @PostMapping("/restore/{fileId}/{versionNumber}")
    public ResponseEntity<Map<String, Object>> restoreToVersion(
            @PathVariable String fileId,
            @PathVariable Integer versionNumber,
            @RequestParam(value = "userId", required = false) String userId) {
        logger.info("Received request to restore file: {} to version {}", fileId, versionNumber);

        try {
            boolean success = versionService.restoreToVersion(fileId, versionNumber, userId);

            if (success) {
                Map<String, Object> response = Map.of(
                    "success", true,
                    "message", "Successfully restored to version " + versionNumber
                );
                logger.info("Successfully restored file: {} to version {}", fileId, versionNumber);
                return ResponseEntity.ok(response);
            } else {
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "Failed to restore to version " + versionNumber
                );
                logger.warn("Failed to restore file: {} to version {}", fileId, versionNumber);
                return ResponseEntity.badRequest().body(response);
            }
        } catch (Exception e) {
            logger.error("Error restoring file: {} to version {}: {}", fileId, versionNumber, e.getMessage(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error restoring to version: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 获取当前活跃版本
     */
    @GetMapping("/current/{fileId}")
    public ResponseEntity<Map<String, Object>> getCurrentVersion(@PathVariable String fileId) {
        logger.info("Received request to get current version for file: {}", fileId);

        try {
            FileVersion version = versionService.getCurrentVersion(fileId);

            if (version != null) {
                Map<String, Object> response = Map.of(
                    "success", true,
                    "data", version
                );
                logger.info("Successfully returned current version for file: {}", fileId);
                return ResponseEntity.ok(response);
            } else {
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "No current version found"
                );
                logger.warn("No current version found for file: {}", fileId);
                return ResponseEntity.badRequest().body(response);
            }
        } catch (Exception e) {
            logger.error("Error retrieving current version for file: {}", fileId, e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error retrieving current version: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 下载特定版本的文件
     */
    @GetMapping("/download/{fileId}/{versionNumber}")
    public ResponseEntity<ByteArrayResource> downloadVersion(
            @PathVariable String fileId,
            @PathVariable Integer versionNumber) {
        logger.info("Received request to download version {} for file: {}", versionNumber, fileId);

        try {
            FileVersion version = versionService.getVersion(fileId, versionNumber);
            
            if (version != null && version.getFileContent() != null) {
                // 解码文件内容
                byte[] fileContent = java.util.Base64.getDecoder().decode(version.getFileContent());
                ByteArrayResource resource = new ByteArrayResource(fileContent);

                return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + version.getFileName() + "\"")
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .contentLength(fileContent.length)
                    .body(resource);
            } else {
                logger.warn("Version {} not found for file: {} or content is null", versionNumber, fileId);
                return ResponseEntity.notFound().build();
            }
        } catch (Exception e) {
            logger.error("Error downloading version {} for file: {}: {}", versionNumber, fileId, e.getMessage(), e);
            return ResponseEntity.badRequest().build();
        }
    }

    /**
     * 比较两个版本
     */
    @GetMapping("/compare/{fileId}")
    public ResponseEntity<Map<String, Object>> compareVersions(
            @PathVariable String fileId,
            @RequestParam Integer version1,
            @RequestParam Integer version2) {
        logger.info("Received request to compare versions {} and {} for file: {}", version1, version2, fileId);

        try {
            String comparison = versionService.compareVersions(fileId, version1, version2);

            if (comparison != null) {
                Map<String, Object> response = Map.of(
                    "success", true,
                    "data", comparison,
                    "version1", version1,
                    "version2", version2
                );
                logger.info("Successfully compared versions {} and {} for file: {}", version1, version2, fileId);
                return ResponseEntity.ok(response);
            } else {
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "Failed to compare versions"
                );
                logger.warn("Failed to compare versions {} and {} for file: {}", version1, version2, fileId);
                return ResponseEntity.badRequest().body(response);
            }
        } catch (Exception e) {
            logger.error("Error comparing versions {} and {} for file: {}: {}", version1, version2, fileId, e.getMessage(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error comparing versions: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 清理旧版本（保留最新的N个版本）
     */
    @DeleteMapping("/cleanup/{fileId}")
    public ResponseEntity<Map<String, Object>> cleanupOldVersions(
            @PathVariable String fileId,
            @RequestParam(defaultValue = "5") int keepLatestCount) {
        logger.info("Received request to cleanup old versions for file: {}, keeping latest {}", fileId, keepLatestCount);

        try {
            int deletedCount = versionService.cleanupOldVersions(fileId, keepLatestCount);

            Map<String, Object> response = Map.of(
                "success", true,
                "deletedCount", deletedCount,
                "message", "Cleaned up " + deletedCount + " old versions"
            );

            logger.info("Successfully cleaned up {} old versions for file: {}", deletedCount, fileId);
            return ResponseEntity.ok(response);
        } catch (Exception e) {
            logger.error("Error cleaning up old versions for file: {}: {}", fileId, e.getMessage(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error cleaning up old versions: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }
}