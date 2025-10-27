package hbnu.project.ergoucsveditior.model;

import java.io.*;
import java.util.Properties;

/**
 * 应用设置管理
 */
public class Settings {
    private static final String SETTINGS_FILE = "csv_editor_settings.properties";
    private Properties properties;
    
    // 设置项
    private String defaultEncoding;
    private int maxHistorySize;
    private boolean autoSaveEnabled;
    private int autoSaveInterval; // 分钟
    private String lineEnding; // LF 或 CRLF
    private String highlightConflictStrategy; // 高亮冲突解决策略
    
    // 高亮偏好设置
    private String defaultCellHighlightColor;      // 默认单元格高亮背景色
    private String defaultCellTextColor;           // 默认单元格文本颜色
    private String defaultRowHighlightColor;       // 默认行高亮背景色
    private String defaultColumnHighlightColor;    // 默认列高亮背景色
    private String searchHighlightColor;           // 搜索结果高亮背景色
    private String selectedRowColor;               // 选中行颜色
    
    // 新增设置项
    private boolean autoDetectDelimiter;           // 自动检测分隔符
    private String escapeMode;                     // 转义字符处理方式（"重复引号" 或 "反斜杠转义"）
    private boolean firstRowAsHeader;              // 首行是否为列标题
    private boolean showLineNumbers;               // 是否显示行号
    private String theme;                          // 主题（"浅色" 或 "深色"）
    private String tableBorderColor;               // 表格边框颜色
    private String tableGridColor;                 // 网格线颜色
    private String backgroundImagePath;            // 背景图片路径
    private double backgroundImageOpacity;         // 背景图片透明度（0.0-1.0）
    private String backgroundImageFitMode;         // 背景图片适应模式
    
    // 列宽和行高设置
    private String columnWidthMode;                // 列宽模式（"自动适配内容" 或 "固定宽度"）
    private String rowHeightMode;                  // 行高模式（"自动适配内容" 或 "固定高度"）
    private double defaultColumnWidth;             // 默认列宽
    private double defaultRowHeight;               // 默认行高
    private double minColumnWidth;                 // 最小列宽
    private double maxColumnWidth;                 // 最大列宽
    private double minRowHeight;                   // 最小行高
    private double maxRowHeight;                   // 最大行高
    private double tableZoomLevel;                 // 表格缩放级别（0.0-1.0，1.0表示100%）
    
    public Settings() {
        properties = new Properties();
        loadDefaults();
        load();
    }
    
    /**
     * 加载默认设置
     */
    private void loadDefaults() {
        defaultEncoding = "UTF-8";
        maxHistorySize = 50;
        autoSaveEnabled = false;
        autoSaveInterval = 5;
        lineEnding = System.lineSeparator().equals("\r\n") ? "CRLF" : "LF";
        highlightConflictStrategy = "覆盖策略";
        
        // 高亮偏好默认值
        defaultCellHighlightColor = "#FFFF99";    // 浅黄色
        defaultCellTextColor = "";                  // 空表示使用默认文本颜色
        defaultRowHighlightColor = "#ADD8E6";      // 浅蓝色
        defaultColumnHighlightColor = "#90EE90";   // 浅绿色
        searchHighlightColor = "#FFD700";          // 金色
        selectedRowColor = "#3498DB";              // 蓝色
        
        // 新增设置项默认值
        autoDetectDelimiter = true;                // 默认开启自动检测分隔符
        escapeMode = "重复引号";                   // 默认使用重复引号转义
        firstRowAsHeader = true;                   // 默认首行为标题
        showLineNumbers = true;                    // 默认显示行号
        theme = "浅色";                            // 默认浅色主题
        tableBorderColor = "#CCCCCC";              // 默认边框颜色
        tableGridColor = "#E0E0E0";                // 默认网格线颜色
        backgroundImagePath = "";                  // 默认无背景图片
        backgroundImageOpacity = 0.5;              // 默认透明度50%
        backgroundImageFitMode = "保持比例";         // 默认保持比例
        
        // 列宽和行高设置默认值
        columnWidthMode = "自动适配内容";          // 默认自动适配列宽
        rowHeightMode = "自动适配内容";            // 默认自动适配行高
        defaultColumnWidth = 100.0;                // 默认列宽100像素
        defaultRowHeight = 25.0;                   // 默认行高25像素
        minColumnWidth = 30.0;                     // 最小列宽30像素
        maxColumnWidth = 500.0;                    // 最大列宽500像素
        minRowHeight = 20.0;                       // 最小行高20像素
        maxRowHeight = 200.0;                      // 最大行高200像素
        tableZoomLevel = 1.0;                      // 默认缩放级别100%
    }
    
    /**
     * 从文件加载设置
     */
    public void load() {
        File file = new File(SETTINGS_FILE);
        if (file.exists()) {
            try (FileInputStream fis = new FileInputStream(file)) {
                properties.load(fis);
                
                defaultEncoding = properties.getProperty("defaultEncoding", defaultEncoding);
                maxHistorySize = Integer.parseInt(properties.getProperty("maxHistorySize", String.valueOf(maxHistorySize)));
                autoSaveEnabled = Boolean.parseBoolean(properties.getProperty("autoSaveEnabled", String.valueOf(autoSaveEnabled)));
                autoSaveInterval = Integer.parseInt(properties.getProperty("autoSaveInterval", String.valueOf(autoSaveInterval)));
                lineEnding = properties.getProperty("lineEnding", lineEnding);
                highlightConflictStrategy = properties.getProperty("highlightConflictStrategy", highlightConflictStrategy);
                
                // 加载高亮偏好设置
                defaultCellHighlightColor = properties.getProperty("defaultCellHighlightColor", defaultCellHighlightColor);
                defaultCellTextColor = properties.getProperty("defaultCellTextColor", defaultCellTextColor);
                defaultRowHighlightColor = properties.getProperty("defaultRowHighlightColor", defaultRowHighlightColor);
                defaultColumnHighlightColor = properties.getProperty("defaultColumnHighlightColor", defaultColumnHighlightColor);
                searchHighlightColor = properties.getProperty("searchHighlightColor", searchHighlightColor);
                selectedRowColor = properties.getProperty("selectedRowColor", selectedRowColor);
                
                // 加载新增设置项
                autoDetectDelimiter = Boolean.parseBoolean(properties.getProperty("autoDetectDelimiter", String.valueOf(autoDetectDelimiter)));
                escapeMode = properties.getProperty("escapeMode", escapeMode);
                firstRowAsHeader = Boolean.parseBoolean(properties.getProperty("firstRowAsHeader", String.valueOf(firstRowAsHeader)));
                showLineNumbers = Boolean.parseBoolean(properties.getProperty("showLineNumbers", String.valueOf(showLineNumbers)));
                theme = properties.getProperty("theme", theme);
                tableBorderColor = properties.getProperty("tableBorderColor", tableBorderColor);
                tableGridColor = properties.getProperty("tableGridColor", tableGridColor);
                backgroundImagePath = properties.getProperty("backgroundImagePath", backgroundImagePath);
                backgroundImageOpacity = Double.parseDouble(properties.getProperty("backgroundImageOpacity", String.valueOf(backgroundImageOpacity)));
                backgroundImageFitMode = properties.getProperty("backgroundImageFitMode", backgroundImageFitMode);
                
                // 加载列宽和行高设置
                columnWidthMode = properties.getProperty("columnWidthMode", columnWidthMode);
                rowHeightMode = properties.getProperty("rowHeightMode", rowHeightMode);
                defaultColumnWidth = Double.parseDouble(properties.getProperty("defaultColumnWidth", String.valueOf(defaultColumnWidth)));
                defaultRowHeight = Double.parseDouble(properties.getProperty("defaultRowHeight", String.valueOf(defaultRowHeight)));
                minColumnWidth = Double.parseDouble(properties.getProperty("minColumnWidth", String.valueOf(minColumnWidth)));
                maxColumnWidth = Double.parseDouble(properties.getProperty("maxColumnWidth", String.valueOf(maxColumnWidth)));
                minRowHeight = Double.parseDouble(properties.getProperty("minRowHeight", String.valueOf(minRowHeight)));
                maxRowHeight = Double.parseDouble(properties.getProperty("maxRowHeight", String.valueOf(maxRowHeight)));
                tableZoomLevel = Double.parseDouble(properties.getProperty("tableZoomLevel", String.valueOf(tableZoomLevel)));
            } catch (IOException | NumberFormatException e) {
                // 加载失败，使用默认设置
                loadDefaults();
            }
        }
    }
    
    /**
     * 保存设置到文件
     */
    public void save() {
        properties.setProperty("defaultEncoding", defaultEncoding);
        properties.setProperty("maxHistorySize", String.valueOf(maxHistorySize));
        properties.setProperty("autoSaveEnabled", String.valueOf(autoSaveEnabled));
        properties.setProperty("autoSaveInterval", String.valueOf(autoSaveInterval));
        properties.setProperty("lineEnding", lineEnding);
        properties.setProperty("highlightConflictStrategy", highlightConflictStrategy);
        
        // 保存高亮偏好设置
        properties.setProperty("defaultCellHighlightColor", defaultCellHighlightColor);
        properties.setProperty("defaultCellTextColor", defaultCellTextColor);
        properties.setProperty("defaultRowHighlightColor", defaultRowHighlightColor);
        properties.setProperty("defaultColumnHighlightColor", defaultColumnHighlightColor);
        properties.setProperty("searchHighlightColor", searchHighlightColor);
        properties.setProperty("selectedRowColor", selectedRowColor);
        
        // 保存新增设置项
        properties.setProperty("autoDetectDelimiter", String.valueOf(autoDetectDelimiter));
        properties.setProperty("escapeMode", escapeMode);
        properties.setProperty("firstRowAsHeader", String.valueOf(firstRowAsHeader));
        properties.setProperty("showLineNumbers", String.valueOf(showLineNumbers));
        properties.setProperty("theme", theme);
        properties.setProperty("tableBorderColor", tableBorderColor);
        properties.setProperty("tableGridColor", tableGridColor);
        properties.setProperty("backgroundImagePath", backgroundImagePath);
        properties.setProperty("backgroundImageOpacity", String.valueOf(backgroundImageOpacity));
        properties.setProperty("backgroundImageFitMode", backgroundImageFitMode);
        
        // 保存列宽和行高设置
        properties.setProperty("columnWidthMode", columnWidthMode);
        properties.setProperty("rowHeightMode", rowHeightMode);
        properties.setProperty("defaultColumnWidth", String.valueOf(defaultColumnWidth));
        properties.setProperty("defaultRowHeight", String.valueOf(defaultRowHeight));
        properties.setProperty("minColumnWidth", String.valueOf(minColumnWidth));
        properties.setProperty("maxColumnWidth", String.valueOf(maxColumnWidth));
        properties.setProperty("minRowHeight", String.valueOf(minRowHeight));
        properties.setProperty("maxRowHeight", String.valueOf(maxRowHeight));
        properties.setProperty("tableZoomLevel", String.valueOf(tableZoomLevel));
        
        try (FileOutputStream fos = new FileOutputStream(SETTINGS_FILE)) {
            properties.store(fos, "CSV Editor Settings");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    // Getters and Setters
    public String getDefaultEncoding() {
        return defaultEncoding;
    }
    
    public void setDefaultEncoding(String defaultEncoding) {
        this.defaultEncoding = defaultEncoding;
    }
    
    public int getMaxHistorySize() {
        return maxHistorySize;
    }
    
    public void setMaxHistorySize(int maxHistorySize) {
        this.maxHistorySize = maxHistorySize;
    }
    
    public boolean isAutoSaveEnabled() {
        return autoSaveEnabled;
    }
    
    public void setAutoSaveEnabled(boolean autoSaveEnabled) {
        this.autoSaveEnabled = autoSaveEnabled;
    }
    
    public int getAutoSaveInterval() {
        return autoSaveInterval;
    }
    
    public void setAutoSaveInterval(int autoSaveInterval) {
        this.autoSaveInterval = autoSaveInterval;
    }
    
    public String getLineEnding() {
        return lineEnding;
    }
    
    public void setLineEnding(String lineEnding) {
        this.lineEnding = lineEnding;
    }
    
    /**
     * 获取换行符字符串
     */
    public String getLineEndingString() {
        return "CRLF".equals(lineEnding) ? "\r\n" : "\n";
    }
    
    public String getHighlightConflictStrategy() {
        return highlightConflictStrategy;
    }
    
    public void setHighlightConflictStrategy(String strategy) {
        this.highlightConflictStrategy = strategy;
    }
    
    // 高亮偏好设置的Getters and Setters
    public String getDefaultCellHighlightColor() {
        return defaultCellHighlightColor;
    }
    
    public void setDefaultCellHighlightColor(String defaultCellHighlightColor) {
        this.defaultCellHighlightColor = defaultCellHighlightColor;
    }
    
    public String getDefaultCellTextColor() {
        return defaultCellTextColor;
    }
    
    public void setDefaultCellTextColor(String defaultCellTextColor) {
        this.defaultCellTextColor = defaultCellTextColor;
    }
    
    public String getDefaultRowHighlightColor() {
        return defaultRowHighlightColor;
    }
    
    public void setDefaultRowHighlightColor(String defaultRowHighlightColor) {
        this.defaultRowHighlightColor = defaultRowHighlightColor;
    }
    
    public String getDefaultColumnHighlightColor() {
        return defaultColumnHighlightColor;
    }
    
    public void setDefaultColumnHighlightColor(String defaultColumnHighlightColor) {
        this.defaultColumnHighlightColor = defaultColumnHighlightColor;
    }
    
    public String getSearchHighlightColor() {
        return searchHighlightColor;
    }
    
    public void setSearchHighlightColor(String searchHighlightColor) {
        this.searchHighlightColor = searchHighlightColor;
    }
    
    public String getSelectedRowColor() {
        return selectedRowColor;
    }
    
    public void setSelectedRowColor(String selectedRowColor) {
        this.selectedRowColor = selectedRowColor;
    }
    
    // 新增设置项的Getters and Setters
    public boolean isAutoDetectDelimiter() {
        return autoDetectDelimiter;
    }
    
    public void setAutoDetectDelimiter(boolean autoDetectDelimiter) {
        this.autoDetectDelimiter = autoDetectDelimiter;
    }
    
    public String getEscapeMode() {
        return escapeMode;
    }
    
    public void setEscapeMode(String escapeMode) {
        this.escapeMode = escapeMode;
    }
    
    public boolean isFirstRowAsHeader() {
        return firstRowAsHeader;
    }
    
    public void setFirstRowAsHeader(boolean firstRowAsHeader) {
        this.firstRowAsHeader = firstRowAsHeader;
    }
    
    public boolean isShowLineNumbers() {
        return showLineNumbers;
    }
    
    public void setShowLineNumbers(boolean showLineNumbers) {
        this.showLineNumbers = showLineNumbers;
    }
    
    public String getTheme() {
        return theme;
    }
    
    public void setTheme(String theme) {
        this.theme = theme;
    }
    
    public String getTableBorderColor() {
        return tableBorderColor;
    }
    
    public void setTableBorderColor(String tableBorderColor) {
        this.tableBorderColor = tableBorderColor;
    }
    
    public String getTableGridColor() {
        return tableGridColor;
    }
    
    public void setTableGridColor(String tableGridColor) {
        this.tableGridColor = tableGridColor;
    }
    
    public String getBackgroundImagePath() {
        return backgroundImagePath;
    }
    
    public void setBackgroundImagePath(String backgroundImagePath) {
        this.backgroundImagePath = backgroundImagePath;
    }
    
    public double getBackgroundImageOpacity() {
        return backgroundImageOpacity;
    }
    
    public void setBackgroundImageOpacity(double backgroundImageOpacity) {
        this.backgroundImageOpacity = backgroundImageOpacity;
    }
    
    public String getBackgroundImageFitMode() {
        return backgroundImageFitMode;
    }
    
    public void setBackgroundImageFitMode(String backgroundImageFitMode) {
        this.backgroundImageFitMode = backgroundImageFitMode;
    }
    
    // 列宽和行高设置的Getters and Setters
    public String getColumnWidthMode() {
        return columnWidthMode;
    }
    
    public void setColumnWidthMode(String columnWidthMode) {
        this.columnWidthMode = columnWidthMode;
    }
    
    public String getRowHeightMode() {
        return rowHeightMode;
    }
    
    public void setRowHeightMode(String rowHeightMode) {
        this.rowHeightMode = rowHeightMode;
    }
    
    public double getDefaultColumnWidth() {
        return defaultColumnWidth;
    }
    
    public void setDefaultColumnWidth(double defaultColumnWidth) {
        this.defaultColumnWidth = defaultColumnWidth;
    }
    
    public double getDefaultRowHeight() {
        return defaultRowHeight;
    }
    
    public void setDefaultRowHeight(double defaultRowHeight) {
        this.defaultRowHeight = defaultRowHeight;
    }
    
    public double getMinColumnWidth() {
        return minColumnWidth;
    }
    
    public void setMinColumnWidth(double minColumnWidth) {
        this.minColumnWidth = minColumnWidth;
    }
    
    public double getMaxColumnWidth() {
        return maxColumnWidth;
    }
    
    public void setMaxColumnWidth(double maxColumnWidth) {
        this.maxColumnWidth = maxColumnWidth;
    }
    
    public double getMinRowHeight() {
        return minRowHeight;
    }
    
    public void setMinRowHeight(double minRowHeight) {
        this.minRowHeight = minRowHeight;
    }
    
    public double getMaxRowHeight() {
        return maxRowHeight;
    }
    
    public void setMaxRowHeight(double maxRowHeight) {
        this.maxRowHeight = maxRowHeight;
    }
    
    public double getTableZoomLevel() {
        return tableZoomLevel;
    }
    
    public void setTableZoomLevel(double tableZoomLevel) {
        this.tableZoomLevel = tableZoomLevel;
    }
}

