package com.example.aiexcel;

import org.junit.jupiter.api.Test;
import org.springframework.core.env.Environment;
import org.springframework.boot.test.context.runner.ApplicationContextRunner;

import static org.junit.jupiter.api.Assertions.*;

public class EnvConfigTest {

    private final ApplicationContextRunner contextRunner = new ApplicationContextRunner();

    @Test
    public void testEnvVariablesLoaded() {
        contextRunner
            .withPropertyValues(
                "qwen.api.api-key=sk-08ea3abd973c4add89bf79d88929067d",
                "qwen.api.base-url=https://dashscope.aliyuncs.com/compatible-mode/v1",
                "qwen.api.default-model=qwen-max"
            )
            .run(context -> {
                Environment env = context.getEnvironment();
                String apiKey = env.getProperty("qwen.api.api-key");
                String baseUrl = env.getProperty("qwen.api.base-url");
                String defaultModel = env.getProperty("qwen.api.default-model");

                assertNotNull(apiKey, "API Key should not be null");
                assertFalse(apiKey.isEmpty(), "API Key should not be empty");
                assertEquals("sk-08ea3abd973c4add89bf79d88929067d", apiKey, "API Key should match the expected value");

                assertEquals("https://dashscope.aliyuncs.com/compatible-mode/v1", baseUrl, "Base URL should match the expected value");
                assertEquals("qwen-max", defaultModel, "Default model should match the expected value");

                System.out.println("API Key loaded successfully: " + (apiKey != null ? "Yes" : "No"));
                System.out.println("Base URL: " + baseUrl);
                System.out.println("Default Model: " + defaultModel);
            });
    }
}