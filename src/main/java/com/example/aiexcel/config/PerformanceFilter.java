package com.example.aiexcel.config;

import jakarta.servlet.FilterChain;
import jakarta.servlet.ServletException;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.Ordered;
import org.springframework.core.annotation.Order;
import org.springframework.stereotype.Component;
import org.springframework.web.filter.OncePerRequestFilter;

import java.io.IOException;

/**
 * API性能日志过滤器
 * 记录每个请求的耗时，超过3秒标记为SLOW
 */
@Component
@Order(Ordered.HIGHEST_PRECEDENCE)
public class PerformanceFilter extends OncePerRequestFilter {

    private static final Logger logger = LoggerFactory.getLogger(PerformanceFilter.class);
    private static final long SLOW_THRESHOLD_MS = 3000;

    @Override
    protected void doFilterInternal(HttpServletRequest request, HttpServletResponse response,
                                    FilterChain filterChain) throws ServletException, IOException {
        long start = System.currentTimeMillis();
        filterChain.doFilter(request, response);
        long duration = System.currentTimeMillis() - start;

        String method = request.getMethod();
        String uri = request.getRequestURI();
        int status = response.getStatus();

        if (duration > SLOW_THRESHOLD_MS) {
            logger.warn("SLOW API: {} {} -> {} ({}ms)", method, uri, status, duration);
        } else {
            logger.info("API: {} {} -> {} ({}ms)", method, uri, status, duration);
        }
    }
}
