package com.example.aiexcel.config;

import org.springframework.cache.CacheManager;
import org.springframework.cache.annotation.EnableCaching;
import org.springframework.cache.concurrent.ConcurrentMapCacheManager;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import java.util.List;

/**
 * 缓存配置
 * 使用ConcurrentMapCacheManager实现内存缓存，适用于2c2g单实例部署
 */
@Configuration
@EnableCaching
public class CacheConfig {

    @Bean
    public CacheManager cacheManager() {
        ConcurrentMapCacheManager manager = new ConcurrentMapCacheManager();
        manager.setCacheNames(List.of("analysis", "templates", "ai-status"));
        return manager;
    }
}
