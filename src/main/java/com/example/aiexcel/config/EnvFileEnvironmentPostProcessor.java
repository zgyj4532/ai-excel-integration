package com.example.aiexcel.config;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.env.EnvironmentPostProcessor;
import org.springframework.core.env.ConfigurableEnvironment;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.support.PropertiesLoaderUtils;

import java.io.*;
import java.util.Properties;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class EnvFileEnvironmentPostProcessor implements EnvironmentPostProcessor {

    @Override
    public void postProcessEnvironment(ConfigurableEnvironment environment, SpringApplication application) {
        // 获取当前工作目录
        String currentDir = System.getProperty("user.dir");
        String envFilePath = currentDir + File.separator + ".env";

        // 手动解析 .env 文件
        Properties properties = loadEnvFile(envFilePath);
        if (properties != null) {
            // 将属性添加到环境中
            environment.getPropertySources().addFirst(new org.springframework.core.env.PropertiesPropertySource("env", properties));

            // 输出调试信息
            System.out.println("Successfully loaded .env file from: " + envFilePath);

            // 输出关键配置值（不包含敏感信息）
            String model = properties.getProperty("QWEN_MODEL_NAME");
            if (model != null) {
                System.out.println("QWEN_MODEL_NAME loaded: " + model);
            }
        } else {
            System.err.println("Could not find or load .env file at: " + envFilePath);
        }
    }

    private Properties loadEnvFile(String envFilePath) {
        Properties properties = new Properties();
        FileSystemResource resource = new FileSystemResource(envFilePath);

        if (!resource.exists()) {
            return null;
        }

        Pattern pattern = Pattern.compile("^([A-Za-z_][A-Za-z0-9_]*)=(.*)$");

        try (BufferedReader reader = new BufferedReader(new InputStreamReader(resource.getInputStream()))) {
            String line;
            while ((line = reader.readLine()) != null) {
                line = line.trim();

                // 忽略空行和注释行
                if (line.isEmpty() || line.startsWith("#") || line.startsWith("!")) {
                    continue;
                }

                // 解析键值对
                Matcher matcher = pattern.matcher(line);
                if (matcher.matches()) {
                    String key = matcher.group(1);
                    String value = matcher.group(2);

                    // 处理带引号的值
                    if ((value.startsWith("\"") && value.endsWith("\"")) ||
                        (value.startsWith("'") && value.endsWith("'"))) {
                        value = value.substring(1, value.length() - 1);
                    }

                    properties.setProperty(key, value);
                }
            }
        } catch (IOException e) {
            System.err.println("Could not load .env file: " + e.getMessage());
            return null;
        }

        return properties;
    }
}