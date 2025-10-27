package hbnu.project.ergoucsveditior.model;

import javafx.scene.paint.Color;

/**
 * 高亮信息类
 * 存储单元格、行、列的高亮颜色和类型
 */
public class HighlightInfo {
    private HighlightType type;
    private Color backgroundColor; // 背景颜色
    private Color textColor;       // 文本颜色
    private long timestamp; // 标记时间戳，用于冲突解决
    
    public enum HighlightType {
        CELL,       // 单元格高亮
        ROW,        // 行高亮
        COLUMN,     // 列高亮
        SEARCH      // 搜索匹配高亮
    }
    
    public HighlightInfo(HighlightType type, Color backgroundColor) {
        this(type, backgroundColor, null);
    }
    
    public HighlightInfo(HighlightType type, Color backgroundColor, Color textColor) {
        this.type = type;
        this.backgroundColor = backgroundColor;
        this.textColor = textColor;
        this.timestamp = System.currentTimeMillis();
    }
    
    public HighlightType getType() {
        return type;
    }
    
    public void setType(HighlightType type) {
        this.type = type;
    }
    
    public Color getBackgroundColor() {
        return backgroundColor;
    }
    
    public void setBackgroundColor(Color backgroundColor) {
        this.backgroundColor = backgroundColor;
    }
    
    public Color getTextColor() {
        return textColor;
    }
    
    public void setTextColor(Color textColor) {
        this.textColor = textColor;
    }
    
    // 保持向后兼容
    public Color getColor() {
        return backgroundColor;
    }
    
    public void setColor(Color color) {
        this.backgroundColor = color;
    }
    
    public long getTimestamp() {
        return timestamp;
    }
    
    public void setTimestamp(long timestamp) {
        this.timestamp = timestamp;
    }
    
    /**
     * 转换背景颜色为CSS格式
     */
    public String getBackgroundColorCss() {
        if (backgroundColor == null) return "";
        return String.format("rgba(%d, %d, %d, %.2f)",
            (int)(backgroundColor.getRed() * 255),
            (int)(backgroundColor.getGreen() * 255),
            (int)(backgroundColor.getBlue() * 255),
            backgroundColor.getOpacity());
    }
    
    /**
     * 转换文本颜色为CSS格式
     */
    public String getTextColorCss() {
        if (textColor == null) return "";
        return String.format("rgba(%d, %d, %d, %.2f)",
            (int)(textColor.getRed() * 255),
            (int)(textColor.getGreen() * 255),
            (int)(textColor.getBlue() * 255),
            textColor.getOpacity());
    }
    
    // 保持向后兼容
    public String getColorCss() {
        return getBackgroundColorCss();
    }
}

