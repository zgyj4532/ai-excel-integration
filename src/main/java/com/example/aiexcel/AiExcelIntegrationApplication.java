package com.example.aiexcel;

import com.example.aiexcel.config.EnvConfig;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Import;
import org.springframework.data.jpa.repository.config.EnableJpaAuditing;

@SpringBootApplication
@Import(EnvConfig.class)  // 导入环境配置
@EnableJpaAuditing
public class AiExcelIntegrationApplication {

    public static void main(String[] args) {
        SpringApplication.run(AiExcelIntegrationApplication.class, args);
    }

}