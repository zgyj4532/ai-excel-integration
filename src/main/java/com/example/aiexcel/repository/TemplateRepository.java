package com.example.aiexcel.repository;

import com.example.aiexcel.model.Template;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import java.util.List;

/**
 * 模板仓库
 */
@Repository
public interface TemplateRepository extends JpaRepository<Template, Long> {

    /**
     * 根据用户ID获取模板列表
     */
    List<Template> findByUserId(String userId);

    /**
     * 根据用户ID和是否公开获取模板列表
     */
    List<Template> findByUserIdAndIsPublic(String userId, Boolean isPublic);

    /**
     * 获取公开模板列表
     */
    List<Template> findByIsPublicTrue();

    /**
     * 根据分类获取模板列表
     */
    List<Template> findByCategory(String category);

    /**
     * 根据用户ID和分类获取模板列表
     */
    List<Template> findByUserIdAndCategory(String userId, String category);

    /**
     * 根据名称模糊搜索模板
     */
    List<Template> findByNameContainingIgnoreCase(String name);

    /**
     * 根据用户ID和名称搜索模板
     */
    List<Template> findByUserIdAndNameContainingIgnoreCase(String userId, String name);

    /**
     * 根据标签搜索模板
     */
    @Query("SELECT t FROM Template t WHERE CONCAT(',', t.tags, ',') LIKE CONCAT('%,', :tag, ',%')")
    List<Template> findByTag(@Param("tag") String tag);

    /**
     * 根据用户ID和标签搜索模板
     */
    @Query("SELECT t FROM Template t WHERE t.userId = :userId AND CONCAT(',', t.tags, ',') LIKE CONCAT('%,', :tag, ',%')")
    List<Template> findByUserIdAndTag(@Param("userId") String userId, @Param("tag") String tag);

    /**
     * 获取热门模板（按使用次数排序，这里简化为按ID降序）
     */
    List<Template> findTop10ByIsPublicTrueOrderByCreatedAtDesc();

    /**
     * 检查用户是否已存在同名模板
     */
    boolean existsByUserIdAndName(String userId, String name);
}