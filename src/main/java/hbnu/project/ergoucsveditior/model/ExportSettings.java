package hbnu.project.ergoucsveditior.model;

import java.io.*;
import java.util.Properties;

/**
 * 导出设置管理
 */
public class ExportSettings {
    private static final String SETTINGS_FILE = "export_settings.properties";
    private Properties properties;
    
    // Excel导出设置
    private boolean excelIncludeHeader;
    private boolean excelAutoSizeColumns;
    private String excelSheetName;
    private boolean excelApplyHighlight;
    
    // PDF导出设置
    private String pdfPageSize; // A4, A3, Letter等
    private String pdfOrientation; // Portrait（纵向）或 Landscape（横向）
    private int pdfFontSize;
    private boolean pdfIncludeHeader;
    private boolean pdfApplyHighlight;
    private String pdfTitle;
    
    public ExportSettings() {
        properties = new Properties();
        loadDefaults();
        load();
    }
    
    /**
     * 加载默认设置
     */
    private void loadDefaults() {
        // Excel默认设置
        excelIncludeHeader = true;
        excelAutoSizeColumns = true;
        excelSheetName = "Sheet1";
        excelApplyHighlight = true;
        
        // PDF默认设置
        pdfPageSize = "A4";
        pdfOrientation = "Portrait";
        pdfFontSize = 10;
        pdfIncludeHeader = true;
        pdfApplyHighlight = true;
        pdfTitle = "CSV数据导出";
    }
    
    /**
     * 从文件加载设置
     */
    public void load() {
        File file = new File(SETTINGS_FILE);
        if (file.exists()) {
            try (FileInputStream fis = new FileInputStream(file)) {
                properties.load(fis);
                
                // 加载Excel设置
                excelIncludeHeader = Boolean.parseBoolean(properties.getProperty("excelIncludeHeader", String.valueOf(excelIncludeHeader)));
                excelAutoSizeColumns = Boolean.parseBoolean(properties.getProperty("excelAutoSizeColumns", String.valueOf(excelAutoSizeColumns)));
                excelSheetName = properties.getProperty("excelSheetName", excelSheetName);
                excelApplyHighlight = Boolean.parseBoolean(properties.getProperty("excelApplyHighlight", String.valueOf(excelApplyHighlight)));
                
                // 加载PDF设置
                pdfPageSize = properties.getProperty("pdfPageSize", pdfPageSize);
                pdfOrientation = properties.getProperty("pdfOrientation", pdfOrientation);
                pdfFontSize = Integer.parseInt(properties.getProperty("pdfFontSize", String.valueOf(pdfFontSize)));
                pdfIncludeHeader = Boolean.parseBoolean(properties.getProperty("pdfIncludeHeader", String.valueOf(pdfIncludeHeader)));
                pdfApplyHighlight = Boolean.parseBoolean(properties.getProperty("pdfApplyHighlight", String.valueOf(pdfApplyHighlight)));
                pdfTitle = properties.getProperty("pdfTitle", pdfTitle);
            } catch (IOException | NumberFormatException e) {
                loadDefaults();
            }
        }
    }
    
    /**
     * 保存设置到文件
     */
    public void save() {
        // 保存Excel设置
        properties.setProperty("excelIncludeHeader", String.valueOf(excelIncludeHeader));
        properties.setProperty("excelAutoSizeColumns", String.valueOf(excelAutoSizeColumns));
        properties.setProperty("excelSheetName", excelSheetName);
        properties.setProperty("excelApplyHighlight", String.valueOf(excelApplyHighlight));
        
        // 保存PDF设置
        properties.setProperty("pdfPageSize", pdfPageSize);
        properties.setProperty("pdfOrientation", pdfOrientation);
        properties.setProperty("pdfFontSize", String.valueOf(pdfFontSize));
        properties.setProperty("pdfIncludeHeader", String.valueOf(pdfIncludeHeader));
        properties.setProperty("pdfApplyHighlight", String.valueOf(pdfApplyHighlight));
        properties.setProperty("pdfTitle", pdfTitle);
        
        try (FileOutputStream fos = new FileOutputStream(SETTINGS_FILE)) {
            properties.store(fos, "Export Settings");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    // Getters and Setters
    public boolean isExcelIncludeHeader() {
        return excelIncludeHeader;
    }
    
    public void setExcelIncludeHeader(boolean excelIncludeHeader) {
        this.excelIncludeHeader = excelIncludeHeader;
    }
    
    public boolean isExcelAutoSizeColumns() {
        return excelAutoSizeColumns;
    }
    
    public void setExcelAutoSizeColumns(boolean excelAutoSizeColumns) {
        this.excelAutoSizeColumns = excelAutoSizeColumns;
    }
    
    public String getExcelSheetName() {
        return excelSheetName;
    }
    
    public void setExcelSheetName(String excelSheetName) {
        this.excelSheetName = excelSheetName;
    }
    
    public boolean isExcelApplyHighlight() {
        return excelApplyHighlight;
    }
    
    public void setExcelApplyHighlight(boolean excelApplyHighlight) {
        this.excelApplyHighlight = excelApplyHighlight;
    }
    
    public String getPdfPageSize() {
        return pdfPageSize;
    }
    
    public void setPdfPageSize(String pdfPageSize) {
        this.pdfPageSize = pdfPageSize;
    }
    
    public String getPdfOrientation() {
        return pdfOrientation;
    }
    
    public void setPdfOrientation(String pdfOrientation) {
        this.pdfOrientation = pdfOrientation;
    }
    
    public int getPdfFontSize() {
        return pdfFontSize;
    }
    
    public void setPdfFontSize(int pdfFontSize) {
        this.pdfFontSize = pdfFontSize;
    }
    
    public boolean isPdfIncludeHeader() {
        return pdfIncludeHeader;
    }
    
    public void setPdfIncludeHeader(boolean pdfIncludeHeader) {
        this.pdfIncludeHeader = pdfIncludeHeader;
    }
    
    public boolean isPdfApplyHighlight() {
        return pdfApplyHighlight;
    }
    
    public void setPdfApplyHighlight(boolean pdfApplyHighlight) {
        this.pdfApplyHighlight = pdfApplyHighlight;
    }
    
    public String getPdfTitle() {
        return pdfTitle;
    }
    
    public void setPdfTitle(String pdfTitle) {
        this.pdfTitle = pdfTitle;
    }
}

