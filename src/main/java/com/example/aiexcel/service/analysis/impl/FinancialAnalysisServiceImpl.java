package com.example.aiexcel.service.analysis.impl;

import com.example.aiexcel.dto.AiRequest;
import com.example.aiexcel.service.ai.AiService;
import com.example.aiexcel.service.excel.ExcelService;
import com.example.aiexcel.service.analysis.FinancialAnalysisService;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.*;

@Service
public class FinancialAnalysisServiceImpl implements FinancialAnalysisService {

    @Autowired
    private AiService aiService;

    @Autowired
    private ExcelService excelService;

    @Override
    public Map<String, Object> analyzeFinancialStatements(MultipartFile file, String analysisType) throws IOException {
        Map<String, Object> result = new HashMap<>();

        // 1. 加载Excel文件
        Workbook workbook = excelService.loadWorkbook(file);

        // 2. 获取Excel数据
        String excelData = excelService.getExcelDataAsString(workbook);

        // 3. 构建AI请求进行财务报表分析
        String prompt = String.format(
            "This is financial data:\n\n%s\n\n" +
            "Perform %s financial statement analysis. " +
            "Analyze key financial metrics, trends, and performance indicators. " +
            "Provide insights on the financial health of the business and recommendations for improvement.",
            excelData, analysisType
        );

        AiRequest aiRequest = new AiRequest();
        aiRequest.setMessages(Arrays.asList(
            new AiRequest.Message("system",
                "You are an expert financial analyst. Analyze financial statements including " +
                "Income Statement, Balance Sheet, and Cash Flow Statement. " +
                "Look for key metrics, trends, and potential issues. " +
                "For Income Statement: focus on revenue growth, cost control, and profitability. " +
                "For Balance Sheet: focus on liquidity, leverage, and asset utilization. " +
                "For Cash Flow Statement: focus on operating cash flow, investment needs, and financing activities. " +
                "Provide actionable insights and recommendations."),
            new AiRequest.Message("user", prompt)
        ));

        // 4. 调用AI服务
        var aiResponse = aiService.generateResponse(aiRequest);

        // 5. 构建结果
        result.put("financialAnalysis", aiResponse.getChoices()[0].getMessage().getContent());
        result.put("excelDataPreview", excelData.substring(0, Math.min(excelData.length(), 500)) + "...");
        result.put("analysisType", analysisType);
        result.put("success", true);

        return result;
    }

    @Override
    public Map<String, Object> calculateFinancialRatios(MultipartFile file) throws IOException {
        Map<String, Object> result = new HashMap<>();

        // 1. 加载Excel文件
        Workbook workbook = excelService.loadWorkbook(file);

        // 2. 获取Excel数据
        String excelData = excelService.getExcelDataAsString(workbook);

        // 3. 构建AI请求进行财务比率计算
        String prompt = String.format(
            "This is financial data:\n\n%s\n\n" +
            "Calculate key financial ratios including: " +
            "1. Liquidity ratios (Current Ratio, Quick Ratio) " +
            "2. Profitability ratios (Gross Profit Margin, Net Profit Margin, ROE, ROA) " +
            "3. Leverage ratios (Debt-to-Equity, Debt Ratio) " +
            "4. Efficiency ratios (Inventory Turnover, Asset Turnover) " +
            "5. Market ratios (if applicable) " +
            "Provide calculations, interpretations, and benchmark comparisons.",
            excelData
        );

        AiRequest aiRequest = new AiRequest();
        aiRequest.setMessages(Arrays.asList(
            new AiRequest.Message("system",
                "You are an expert in financial ratio analysis. Calculate and interpret key ratios:\n" +
                "Liquidity Ratios:\n" +
                "- Current Ratio = Current Assets / Current Liabilities\n" +
                "- Quick Ratio = (Current Assets - Inventory) / Current Liabilities\n\n" +
                "Profitability Ratios:\n" +
                "- Gross Profit Margin = (Revenue - Cost of Goods Sold) / Revenue\n" +
                "- Net Profit Margin = Net Income / Revenue\n" +
                "- Return on Assets (ROA) = Net Income / Total Assets\n" +
                "- Return on Equity (ROE) = Net Income / Shareholder's Equity\n\n" +
                "Leverage Ratios:\n" +
                "- Debt-to-Equity = Total Debt / Total Equity\n" +
                "- Debt Ratio = Total Debt / Total Assets\n\n" +
                "Efficiency Ratios:\n" +
                "- Asset Turnover = Revenue / Average Total Assets\n" +
                "- Inventory Turnover = Cost of Goods Sold / Average Inventory\n\n" +
                "Provide interpretations of what each ratio indicates about the company's performance " +
                "and how they compare to industry benchmarks."),
            new AiRequest.Message("user", prompt)
        ));

        // 4. 调用AI服务
        var aiResponse = aiService.generateResponse(aiRequest);

        // 5. 构建结果
        result.put("financialRatios", aiResponse.getChoices()[0].getMessage().getContent());
        result.put("excelDataPreview", excelData.substring(0, Math.min(excelData.length(), 500)) + "...");
        result.put("analysisType", "Financial Ratios");
        result.put("success", true);

        return result;
    }

    @Override
    public Map<String, Object> analyzeProfitability(MultipartFile file) throws IOException {
        Map<String, Object> result = new HashMap<>();

        // 1. 加载Excel文件
        Workbook workbook = excelService.loadWorkbook(file);

        // 2. 获取Excel数据
        String excelData = excelService.getExcelDataAsString(workbook);

        // 3. 构建AI请求进行盈利能力分析
        String prompt = String.format(
            "This is financial data:\n\n%s\n\n" +
            "Perform comprehensive profitability analysis. " +
            "Analyze gross profit, operating profit, and net profit margins. " +
            "Identify key drivers of profitability and areas for improvement. " +
            "Compare profitability across time periods or business segments.",
            excelData
        );

        AiRequest aiRequest = new AiRequest();
        aiRequest.setMessages(Arrays.asList(
            new AiRequest.Message("system",
                "You are an expert in profitability analysis. Focus on:\n" +
                "1. Gross Profit Margin = (Revenue - Cost of Goods Sold) / Revenue\n" +
                "2. Operating Profit Margin = Operating Income / Revenue\n" +
                "3. Net Profit Margin = Net Income / Revenue\n\n" +
                "Analyze the profitability drivers including:\n" +
                "- Revenue trends and growth\n" +
                "- Cost structure and control\n" +
                "- Operational efficiency\n" +
                "- Product/service mix\n\n" +
                "Identify areas of improvement such as cost reduction opportunities, " +
                "pricing strategies, or operational optimizations. " +
                "Compare current profitability to historical performance and industry benchmarks."),
            new AiRequest.Message("user", prompt)
        ));

        // 4. 调用AI服务
        var aiResponse = aiService.generateResponse(aiRequest);

        // 5. 构建结果
        result.put("profitabilityAnalysis", aiResponse.getChoices()[0].getMessage().getContent());
        result.put("excelDataPreview", excelData.substring(0, Math.min(excelData.length(), 500)) + "...");
        result.put("analysisType", "Profitability Analysis");
        result.put("success", true);

        return result;
    }

    @Override
    public Map<String, Object> analyzeCashFlow(MultipartFile file) throws IOException {
        Map<String, Object> result = new HashMap<>();

        // 1. 加载Excel文件
        Workbook workbook = excelService.loadWorkbook(file);

        // 2. 获取Excel数据
        String excelData = excelService.getExcelDataAsString(workbook);

        // 3. 构建AI请求进行现金流分析
        String prompt = String.format(
            "This is financial data:\n\n%s\n\n" +
            "Perform comprehensive cash flow analysis. " +
            "Analyze operating, investing, and financing cash flows. " +
            "Evaluate cash flow trends, sustainability, and adequacy for business operations. " +
            "Identify potential cash flow issues and provide recommendations.",
            excelData
        );

        AiRequest aiRequest = new AiRequest();
        aiRequest.setMessages(Arrays.asList(
            new AiRequest.Message("system",
                "You are an expert in cash flow analysis. Focus on:\n" +
                "1. Operating Cash Flow (OCF) - cash generated from core business activities\n" +
                "2. Investing Cash Flow - cash used in or generated from investments\n" +
                "3. Financing Cash Flow - cash from debt, equity, and dividend activities\n\n" +
                "Analyze:\n" +
                "- Operating cash flow trends and sustainability\n" +
                "- Free Cash Flow = Operating Cash Flow - Capital Expenditures\n" +
                "- Cash flow to debt ratios\n" +
                "- Seasonal patterns in cash flow\n" +
                "- Days Sales Outstanding, Days Payable Outstanding, Days Inventory Outstanding\n\n" +
                "Identify potential liquidity issues and recommend cash management strategies."),
            new AiRequest.Message("user", prompt)
        ));

        // 4. 调用AI服务
        var aiResponse = aiService.generateResponse(aiRequest);

        // 5. 构建结果
        result.put("cashFlowAnalysis", aiResponse.getChoices()[0].getMessage().getContent());
        result.put("excelDataPreview", excelData.substring(0, Math.min(excelData.length(), 500)) + "...");
        result.put("analysisType", "Cash Flow Analysis");
        result.put("success", true);

        return result;
    }

    @Override
    public Map<String, Object> compareBudgetVsActual(MultipartFile file) throws IOException {
        Map<String, Object> result = new HashMap<>();

        // 1. 加载Excel文件
        Workbook workbook = excelService.loadWorkbook(file);

        // 2. 获取Excel数据
        String excelData = excelService.getExcelDataAsString(workbook);

        // 3. 构建AI请求进行预算与实际对比分析
        String prompt = String.format(
            "This is budget vs actual financial data:\n\n%s\n\n" +
            "Perform budget vs actual variance analysis. " +
            "Calculate variances for key line items (revenue, costs, expenses). " +
            "Determine whether variances are favorable or unfavorable. " +
            "Provide explanations for significant variances and recommendations for future budgeting.",
            excelData
        );

        AiRequest aiRequest = new AiRequest();
        aiRequest.setMessages(Arrays.asList(
            new AiRequest.Message("system",
                "You are an expert in variance analysis and budget planning. For each line item, calculate:\n" +
                "1. Variance = Actual - Budget\n" +
                "2. Percentage Variance = (Actual - Budget) / Budget * 100\n" +
                "3. Identify whether variances are favorable (F) or unfavorable (U)\n\n" +
                "For revenue variances:\n" +
                "- Positive variance = favorable (more revenue than budgeted)\n" +
                "- Negative variance = unfavorable (less revenue than budgeted)\n\n" +
                "For cost/expense variances:\n" +
                "- Positive variance = unfavorable (more costs than budgeted)\n" +
                "- Negative variance = favorable (less costs than budgeted)\n\n" +
                "Analyze significant variances (typically >5% or >absolute threshold), " +
                "provide potential causes, and recommend corrective actions. " +
                "Also suggest improvements to budgeting process based on variance patterns."),
            new AiRequest.Message("user", prompt)
        ));

        // 4. 调用AI服务
        var aiResponse = aiService.generateResponse(aiRequest);

        // 5. 构建结果
        result.put("budgetActualAnalysis", aiResponse.getChoices()[0].getMessage().getContent());
        result.put("excelDataPreview", excelData.substring(0, Math.min(excelData.length(), 500)) + "...");
        result.put("analysisType", "Budget vs Actual");
        result.put("success", true);

        return result;
    }
}