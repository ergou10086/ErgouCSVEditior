package hbnu.project.ergoucsveditior.service;

import hbnu.project.ergoucsveditior.model.CSVData;
import hbnu.project.ergoucsveditior.model.ExportSettings;
import hbnu.project.ergoucsveditior.model.HighlightManager;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.*;


import java.io.File;
import java.io.FileOutputStream;

/**
 * Excel导出服务
 * 使用Apache POI实现
 */
public class ExcelExportService {
    
    /**
     * 导出CSV数据到Excel文件
     * 
     * @param csvData CSV数据
     * @param file 目标文件
     * @param settings 导出设置
     * @param highlightManager 高亮管理器
     * @throws Exception 导出异常
     */
    public void exportToExcel(CSVData csvData, File file, ExportSettings settings, 
                              HighlightManager highlightManager) throws Exception {
        
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet(settings.getExcelSheetName());
        
        int startRow = 0;
        
        // 创建标题样式
        XSSFCellStyle headerStyle = workbook.createCellStyle();
        XSSFFont headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short)11);
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        
        // 创建通用单元格样式（确保中文正确显示）
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        XSSFFont cellFont = workbook.createFont();
        cellFont.setFontName("宋体"); // 使用宋体支持中文
        cellFont.setFontHeightInPoints((short)10);
        cellStyle.setFont(cellFont);
        cellStyle.setWrapText(true); // 自动换行
        
        // 如果包含标题行
        if (settings.isExcelIncludeHeader()) {
            XSSFRow headerRow = sheet.createRow(startRow++);
            for (int col = 0; col < csvData.getColumns(); col++) {
                XSSFCell cell = headerRow.createCell(col);
                cell.setCellValue("列 " + (col + 1));
                cell.setCellStyle(headerStyle);
            }
        }
        
        // 写入数据
        for (int row = 0; row < csvData.getRows(); row++) {
            XSSFRow excelRow = sheet.createRow(startRow + row);
            for (int col = 0; col < csvData.getColumns(); col++) {
                XSSFCell cell = excelRow.createCell(col);
                String value = csvData.getCellValue(row, col);
                cell.setCellValue(value != null ? value : "");
                
                // 应用样式和高亮
                if (settings.isExcelApplyHighlight() && highlightManager != null) {
                    javafx.scene.paint.Color highlightColor = 
                        highlightManager.getFinalHighlightColor(row, col);
                    if (highlightColor != null) {
                        XSSFCellStyle highlightStyle = workbook.createCellStyle();
                        highlightStyle.cloneStyleFrom(cellStyle);
                        
                        byte[] rgb = new byte[]{
                            (byte)(highlightColor.getRed() * 255),
                            (byte)(highlightColor.getGreen() * 255),
                            (byte)(highlightColor.getBlue() * 255)
                        };
                        XSSFColor xssfColor = new XSSFColor(rgb, null);
                        highlightStyle.setFillForegroundColor(xssfColor);
                        highlightStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                        cell.setCellStyle(highlightStyle);
                    } else {
                        cell.setCellStyle(cellStyle);
                    }
                } else {
                    cell.setCellStyle(cellStyle);
                }
            }
        }
        
        // 自动调整列宽
        if (settings.isExcelAutoSizeColumns()) {
            for (int col = 0; col < csvData.getColumns(); col++) {
                sheet.autoSizeColumn(col);
                // 为中文字符调整宽度（POI的autoSizeColumn对中文支持不好）
                int currentWidth = sheet.getColumnWidth(col);
                sheet.setColumnWidth(col, (int)(currentWidth * 1.2)); // 增加20%宽度
            }
        } else {
            // 设置默认列宽
            for (int col = 0; col < csvData.getColumns(); col++) {
                sheet.setColumnWidth(col, 4000); // 默认宽度
            }
        }
        
        // 写入文件
        try (FileOutputStream outputStream = new FileOutputStream(file)) {
            workbook.write(outputStream);
        }
        
        workbook.close();
    }
}

