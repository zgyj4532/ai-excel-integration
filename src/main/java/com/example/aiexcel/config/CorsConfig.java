package com.example.aiexcel.config;

import org.springframework.context.annotation.Configuration;
import org.springframework.web.servlet.config.annotation.CorsRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;

@Configuration
public class CorsConfig implements WebMvcConfigurer {

    @Override
    public void addCorsMappings(CorsRegistry registry) {
        String allowed = System.getenv("API_ALLOW_ORIGIN");
        if (allowed == null || allowed.isEmpty()) {
            // 默认：包含前端 dev server 和后端当前 server.port
            String serverPort = System.getProperty("server.port");
            if (serverPort == null || serverPort.isEmpty()) serverPort = "8080";
            // 常见本地前端端口 5173（Vite），以及后端运行端口
            allowed = "http://localhost:5173,http://localhost:" + serverPort;
        }

        String[] origins = allowed.split("\\s*,\\s*");

        registry.addMapping("/api/**")
                .allowedOrigins(origins)
                .allowedMethods("GET", "POST", "PUT", "DELETE", "OPTIONS")
                .allowedHeaders("*")
                .allowCredentials(false)
                .maxAge(3600);
    }
}
