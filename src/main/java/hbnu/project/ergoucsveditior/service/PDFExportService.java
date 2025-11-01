package hbnu.project.ergoucsveditior.service;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import hbnu.project.ergoucsveditior.model.CSVData;
import hbnu.project.ergoucsveditior.settings.ExportSettings;
import hbnu.project.ergoucsveditior.manager.HighlightManager;


import java.io.File;
import java.io.FileOutputStream;

/**
 * PDF导出服务
 * 使用iText实现
 */
public class PDFExportService {
    
    /**
     * 导出CSV数据到PDF文件
     * 
     * @param csvData CSV数据
     * @param file 目标文件
     * @param settings 导出设置
     * @param highlightManager 高亮管理器
     * @throws Exception 导出异常
     */
    public void exportToPDF(CSVData csvData, File file, ExportSettings settings,
                           HighlightManager highlightManager) throws Exception {
        
        Document document = new Document();
        
        // 设置页面大小和方向
        Rectangle pageSize = getPageSize(settings.getPdfPageSize());
        if ("Landscape".equals(settings.getPdfOrientation())) {
            pageSize = pageSize.rotate();
        }
        document.setPageSize(pageSize);
        
        PdfWriter.getInstance(document, new FileOutputStream(file));
        document.open();
        
        // 设置中文字体（使用iTextAsian提供的中文字体）
        BaseFont bfChinese = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
        
        // 添加标题（使用中文字体）
        Font titleFont = new Font(bfChinese, 16, Font.BOLD);
        Paragraph title = new Paragraph(settings.getPdfTitle(), titleFont);
        title.setAlignment(Element.ALIGN_CENTER);
        document.add(title);
        document.add(new Paragraph("\n"));
        
        // 创建表格
        PdfPTable table = new PdfPTable(csvData.getColumns());
        table.setWidthPercentage(100);
        table.setSpacingBefore(10f);
        table.setSpacingAfter(10f);
        
        // 设置字体（使用中文字体）
        Font cellFont = new Font(bfChinese, settings.getPdfFontSize(), Font.NORMAL);
        Font headerFont = new Font(bfChinese, settings.getPdfFontSize(), Font.BOLD);
        
        // 添加标题行
        if (settings.isPdfIncludeHeader()) {
            for (int col = 0; col < csvData.getColumns(); col++) {
                PdfPCell headerCell = new PdfPCell(new Phrase("列 " + (col + 1), headerFont));
                headerCell.setBackgroundColor(BaseColor.LIGHT_GRAY);
                headerCell.setHorizontalAlignment(Element.ALIGN_CENTER);
                headerCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
                headerCell.setPadding(5);
                table.addCell(headerCell);
            }
        }
        
        // 添加数据
        for (int row = 0; row < csvData.getRows(); row++) {
            for (int col = 0; col < csvData.getColumns(); col++) {
                String value = csvData.getCellValue(row, col);
                PdfPCell cell = new PdfPCell(new Phrase(value != null ? value : "", cellFont));
                cell.setPadding(3);
                cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
                
                // 应用高亮
                if (settings.isPdfApplyHighlight() && highlightManager != null) {
                    javafx.scene.paint.Color highlightColor = 
                        highlightManager.getFinalHighlightColor(row, col);
                    if (highlightColor != null) {
                        BaseColor pdfColor = new BaseColor(
                            (int)(highlightColor.getRed() * 255),
                            (int)(highlightColor.getGreen() * 255),
                            (int)(highlightColor.getBlue() * 255)
                        );
                        cell.setBackgroundColor(pdfColor);
                    }
                }
                
                table.addCell(cell);
            }
        }
        
        document.add(table);
        document.close();
    }
    
    /**
     * 获取页面大小
     */
    private Rectangle getPageSize(String pageSize) {
        return switch (pageSize) {
            case "A3" -> PageSize.A3;
            case "Letter" -> PageSize.LETTER;
            case "Legal" -> PageSize.LEGAL;
            default -> PageSize.A4;
        };
    }
}

