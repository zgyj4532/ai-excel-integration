package com.example.aiexcel.service.analysis.impl;

import com.example.aiexcel.dto.AiRequest;
import com.example.aiexcel.service.ai.AiService;
import com.example.aiexcel.service.excel.ExcelService;
import com.example.aiexcel.service.analysis.CustomerAnalysisService;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.*;

@Service
public class CustomerAnalysisServiceImpl implements CustomerAnalysisService {

    @Autowired
    private AiService aiService;

    @Autowired
    private ExcelService excelService;

    @Override
    public Map<String, Object> performRFMAnalysis(MultipartFile file) throws IOException {
        Map<String, Object> result = new HashMap<>();

        // 1. 加载Excel文件
        Workbook workbook = excelService.loadWorkbook(file);

        // 2. 获取Excel数据
        String excelData = excelService.getExcelDataAsString(workbook);

        // 3. 构建AI请求进行RFM分析
        String prompt = String.format(
            "This is customer transaction data:\n\n%s\n\n" +
            "Perform RFM (Recency, Frequency, Monetary) analysis on this data. " +
            "Segment customers based on RFM scores (1-5 scale for each factor). " +
            "Identify high-value customers, at-risk customers, and sleeping beauties. " +
            "Provide a detailed analysis of customer segments and recommendations for each segment.",
            excelData
        );

        AiRequest aiRequest = new AiRequest();
        aiRequest.setMessages(Arrays.asList(
            new AiRequest.Message("system",
                "You are an expert in customer analytics and RFM analysis. " +
                "RFM stands for Recency (how recently a customer has purchased), " +
                "Frequency (how often a customer purchases), and " +
                "Monetary (how much a customer spends). " +
                "Use the data provided to calculate RFM scores and segment customers. " +
                "Typically, higher scores indicate better customers. " +
                "For Recency: higher score for more recent purchases. " +
                "For Frequency: higher score for more frequent purchases. " +
                "For Monetary: higher score for higher spending. " +
                "Provide actionable insights for each segment."),
            new AiRequest.Message("user", prompt)
        ));

        // 4. 调用AI服务
        var aiResponse = aiService.generateResponse(aiRequest);

        // 5. 构建结果
        result.put("rfmAnalysis", aiResponse.getChoices()[0].getMessage().getContent());
        result.put("excelDataPreview", excelData.substring(0, Math.min(excelData.length(), 500)) + "...");
        result.put("analysisType", "RFM Analysis");
        result.put("success", true);

        return result;
    }

    @Override
    public Map<String, Object> calculateCustomerLifetimeValue(MultipartFile file) throws IOException {
        Map<String, Object> result = new HashMap<>();

        // 1. 加载Excel文件
        Workbook workbook = excelService.loadWorkbook(file);

        // 2. 获取Excel数据
        String excelData = excelService.getExcelDataAsString(workbook);

        // 3. 构建AI请求进行CLV计算
        String prompt = String.format(
            "This is customer data:\n\n%s\n\n" +
            "Calculate Customer Lifetime Value (CLV) for the customers. " +
            "Consider factors like average purchase value, purchase frequency, customer lifespan, and profit margin. " +
            "Identify high-value customers based on their CLV scores. " +
            "Provide a detailed breakdown of how CLV was calculated and recommendations for customer retention strategies.",
            excelData
        );

        AiRequest aiRequest = new AiRequest();
        aiRequest.setMessages(Arrays.asList(
            new AiRequest.Message("system",
                "You are an expert in customer analytics and financial modeling. " +
                "Customer Lifetime Value (CLV) is calculated as: " +
                "CLV = (Average Purchase Value × Purchase Frequency) × Customer Lifespan × Profit Margin. " +
                "Alternatively, it can be calculated using: " +
                "CLV = Average Order Value × Number of Repeat Purchases × Average Customer Lifespan. " +
                "Provide a detailed analysis of customer segments based on their CLV, " +
                "and suggest strategies to increase CLV for different segments."),
            new AiRequest.Message("user", prompt)
        ));

        // 4. 调用AI服务
        var aiResponse = aiService.generateResponse(aiRequest);

        // 5. 构建结果
        result.put("clvAnalysis", aiResponse.getChoices()[0].getMessage().getContent());
        result.put("excelDataPreview", excelData.substring(0, Math.min(excelData.length(), 500)) + "...");
        result.put("analysisType", "Customer Lifetime Value");
        result.put("success", true);

        return result;
    }

    @Override
    public Map<String, Object> segmentCustomers(MultipartFile file) throws IOException {
        Map<String, Object> result = new HashMap<>();

        // 1. 加载Excel文件
        Workbook workbook = excelService.loadWorkbook(file);

        // 2. 获取Excel数据
        String excelData = excelService.getExcelDataAsString(workbook);

        // 3. 构建AI请求进行客户细分
        String prompt = String.format(
            "This is customer data:\n\n%s\n\n" +
            "Perform customer segmentation analysis. " +
            "Use various criteria like demographics, behavior, purchase history, and value to segment customers. " +
            "Classify customers into segments such as VIP, Regular, Potential, At-risk, etc. " +
            "Provide characteristics of each segment and specific marketing strategies for each segment.",
            excelData
        );

        AiRequest aiRequest = new AiRequest();
        aiRequest.setMessages(Arrays.asList(
            new AiRequest.Message("system",
                "You are an expert in customer segmentation and behavioral analytics. " +
                "Segment customers using multiple criteria including but not limited to: " +
                "Demographic (age, location, etc.), Behavioral (purchase patterns, engagement), " +
                "Psychographic (preferences, lifestyle), and Value-based (RFM, CLV). " +
                "Common segmentation models include: " +
                "1. ABC Analysis (based on value) " +
                "2. Demographic Segmentation " +
                "3. Behavioral Segmentation " +
                "4. Psychographic Segmentation " +
                "5. Geographic Segmentation " +
                "Provide actionable insights for each segment and suggest tailored strategies."),
            new AiRequest.Message("user", prompt)
        ));

        // 4. 调用AI服务
        var aiResponse = aiService.generateResponse(aiRequest);

        // 5. 构建结果
        result.put("customerSegmentation", aiResponse.getChoices()[0].getMessage().getContent());
        result.put("excelDataPreview", excelData.substring(0, Math.min(excelData.length(), 500)) + "...");
        result.put("analysisType", "Customer Segmentation");
        result.put("success", true);

        return result;
    }

    @Override
    public Map<String, Object> predictChurnRisk(MultipartFile file) throws IOException {
        Map<String, Object> result = new HashMap<>();

        // 1. 加载Excel文件
        Workbook workbook = excelService.loadWorkbook(file);

        // 2. 获取Excel数据
        String excelData = excelService.getExcelDataAsString(workbook);

        // 3. 构建AI请求进行流失风险预测
        String prompt = String.format(
            "This is customer data:\n\n%s\n\n" +
            "Analyze customer churn risk. " +
            "Identify customers who are most likely to stop using the service or product. " +
            "Consider factors such as: decrease in purchase frequency, longer time since last purchase, " +
            "decrease in order value, inactivity periods, customer complaints, etc. " +
            "Provide a risk score for each customer segment and suggest retention strategies.",
            excelData
        );

        AiRequest aiRequest = new AiRequest();
        aiRequest.setMessages(Arrays.asList(
            new AiRequest.Message("system",
                "You are an expert in customer retention and churn prediction. " +
                "Identify churn indicators such as: " +
                "1. Decreased engagement/purchases over time " +
                "2. Longer intervals between purchases " +
                "3. Reduced order values " +
                "4. Lack of response to marketing efforts " +
                "5. Decreased customer service interactions " +
                "6. Price sensitivity " +
                "7. Switch to competitors " +
                "Provide a risk classification (Low/Medium/High), " +
                "list customers at highest risk, and suggest targeted retention strategies."),
            new AiRequest.Message("user", prompt)
        ));

        // 4. 调用AI服务
        var aiResponse = aiService.generateResponse(aiRequest);

        // 5. 构建结果
        result.put("churnAnalysis", aiResponse.getChoices()[0].getMessage().getContent());
        result.put("excelDataPreview", excelData.substring(0, Math.min(excelData.length(), 500)) + "...");
        result.put("analysisType", "Churn Risk Prediction");
        result.put("success", true);

        return result;
    }

    @Override
    public Map<String, Object> calculateCACvsCLV(MultipartFile file) throws IOException {
        Map<String, Object> result = new HashMap<>();

        // 1. 加载Excel文件
        Workbook workbook = excelService.loadWorkbook(file);

        // 2. 获取Excel数据
        String excelData = excelService.getExcelDataAsString(workbook);

        // 3. 构建AI请求进行CAC与CLV分析
        String prompt = String.format(
            "This is customer acquisition and transaction data:\n\n%s\n\n" +
            "Calculate and analyze Customer Acquisition Cost (CAC) versus Customer Lifetime Value (CLV). " +
            "Determine the CAC:CLV ratio and provide insights on acquisition efficiency. " +
            "Identify the most cost-effective acquisition channels and suggest optimization strategies. " +
            "Explain the importance of maintaining a healthy CAC:CLV ratio (typically 1:3 as a benchmark).",
            excelData
        );

        AiRequest aiRequest = new AiRequest();
        aiRequest.setMessages(Arrays.asList(
            new AiRequest.Message("system",
                "You are an expert in growth metrics and customer acquisition analytics. " +
                "Customer Acquisition Cost (CAC) is calculated as: Total Marketing & Sales Expenses / Number of New Customers Acquired. " +
                "Customer Lifetime Value (CLV) is calculated as: Average Order Value × Purchase Frequency × Customer Lifespan. " +
                "The CAC:CLV ratio is a critical metric that indicates acquisition efficiency. " +
                "A healthy ratio is typically 1:3 (CLV should be 3x CAC). " +
                "If the ratio is too low (e.g., 1:1), it means the company is spending too much to acquire customers. " +
                "If the ratio is too high (e.g., 1:10), it might mean the company is under-spending on acquisition. " +
                "Provide analysis of the ratio, identify most effective channels, and suggest optimization strategies."),
            new AiRequest.Message("user", prompt)
        ));

        // 4. 调用AI服务
        var aiResponse = aiService.generateResponse(aiRequest);

        // 5. 构建结果
        result.put("cacClvAnalysis", aiResponse.getChoices()[0].getMessage().getContent());
        result.put("excelDataPreview", excelData.substring(0, Math.min(excelData.length(), 500)) + "...");
        result.put("analysisType", "CAC vs CLV Analysis");
        result.put("success", true);

        return result;
    }

    @Override
    public Map<String, Object> analyzeCustomerCohorts(MultipartFile file) throws IOException {
        Map<String, Object> result = new HashMap<>();

        // 1. 加载Excel文件
        Workbook workbook = excelService.loadWorkbook(file);

        // 2. 获取Excel数据
        String excelData = excelService.getExcelDataAsString(workbook);

        // 3. 构建AI请求进行队列分析
        String prompt = String.format(
            "This is customer transaction data with dates:\n\n%s\n\n" +
            "Perform customer cohort analysis. " +
            "Group customers by acquisition period (e.g., month or quarter of first purchase) " +
            "and track their retention rates over time. " +
            "Calculate retention rates for each cohort and identify patterns. " +
            "Provide insights on which cohorts have the best retention and suggest reasons.",
            excelData
        );

        AiRequest aiRequest = new AiRequest();
        aiRequest.setMessages(Arrays.asList(
            new AiRequest.Message("system",
                "You are an expert in cohort analysis and retention metrics. " +
                "Cohort analysis tracks groups of users who share a common characteristic over time. " +
                "Typical cohorts are based on acquisition date (signup month/quarter). " +
                "For each cohort, calculate and track:\n" +
                "1. Retention rate over time (1-day, 7-day, 30-day, etc.)\n" +
                "2. Revenue per user over time\n" +
                "3. Engagement metrics over time\n" +
                "4. Identify trends in customer loyalty\n" +
                "5. Compare performance between different cohorts\n" +
                "Explain the insights and suggest actions based on cohort performance."),
            new AiRequest.Message("user", prompt)
        ));

        // 4. 调用AI服务
        var aiResponse = aiService.generateResponse(aiRequest);

        // 5. 构建结果
        result.put("cohortAnalysis", aiResponse.getChoices()[0].getMessage().getContent());
        result.put("excelDataPreview", excelData.substring(0, Math.min(excelData.length(), 500)) + "...");
        result.put("analysisType", "Cohort Analysis");
        result.put("success", true);

        return result;
    }
}