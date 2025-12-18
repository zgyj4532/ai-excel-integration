import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class FormulaTest {
    public static void main(String[] args) throws IOException {
        // 创建工作簿和工作表
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Test Sheet");

        // 创建一些测试数据
        Row row1 = sheet.createRow(0);
        row1.createCell(0).setCellValue("Name");
        row1.createCell(1).setCellValue("Value1");
        row1.createCell(2).setCellValue("Value2");
        row1.createCell(3).setCellValue("Total");

        Row row2 = sheet.createRow(1);
        row2.createCell(0).setCellValue("Item1");
        row2.createCell(1).setCellValue(10);
        row2.createCell(2).setCellValue(20);

        Row row3 = sheet.createRow(2);
        row3.createCell(0).setCellValue("Item2");
        row3.createCell(1).setCellValue(30);
        row3.createCell(2).setCellValue(40);

        // 设置公式
        Cell formulaCell1 = row2.createCell(3);
        formulaCell1.setCellFormula("B2+C2");  // 不以=开头，这是正确的POI方式

        Cell formulaCell2 = row3.createCell(3);
        formulaCell2.setCellFormula("B3+C3");

        // 创建公式计算器并计算
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        evaluator.evaluateInCell(formulaCell1);
        evaluator.evaluateInCell(formulaCell2);

        // 保存文件
        FileOutputStream fileOut = new FileOutputStream("formula_test.xlsx");
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();

        System.out.println("Excel文件已创建，公式已计算");
        
        // 重新打开并验证结果
        Workbook wb = new XSSFWorkbook();
        Sheet sh = wb.createSheet("Test Sheet");

        Row r1 = sh.createRow(0);
        r1.createCell(0).setCellValue("Name");
        r1.createCell(1).setCellValue("Value1");
        r1.createCell(2).setCellValue("Value2");
        r1.createCell(3).setCellValue("Total");

        Row r2 = sh.createRow(1);
        r2.createCell(0).setCellValue("Item1");
        r2.createCell(1).setCellValue(10);
        r2.createCell(2).setCellValue(20);

        Row r3 = sh.createRow(2);
        r3.createCell(0).setCellValue("Item2");
        r3.createCell(1).setCellValue(30);
        r3.createCell(2).setCellValue(40);

        Cell fc1 = r2.createCell(3);
        fc1.setCellFormula("B2+C2");

        Cell fc2 = r3.createCell(3);
        fc2.setCellFormula("B3+C3");

        FormulaEvaluator fe = wb.getCreationHelper().createFormulaEvaluator();
        fe.evaluateInCell(fc1);
        fe.evaluateInCell(fc2);

        System.out.println("公式单元格1的值: " + fc1.getNumericCellValue()); // 应该是30
        System.out.println("公式单元格2的值: " + fc2.getNumericCellValue()); // 应该是70
    }
}