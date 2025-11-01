package hbnu.project.ergoucsveditior.settings;

import hbnu.project.ergoucsveditior.manager.ConfigManager;
import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Properties;

/**
 * 工具栏配置模型
 * 管理工具栏按钮的显示和样式
 */
public class ToolbarConfig {
    private Properties properties;
    
    // 按钮ID常量
    public static final String BTN_NEW = "new";
    public static final String BTN_OPEN = "open";
    public static final String BTN_SAVE = "save";
    public static final String BTN_SAVE_AS = "saveAs";
    public static final String BTN_UNDO = "undo";
    public static final String BTN_ADD_ROW = "addRow";
    public static final String BTN_ADD_COLUMN = "addColumn";
    public static final String BTN_DELETE_ROW = "deleteRow";
    public static final String BTN_DELETE_COLUMN = "deleteColumn";
    public static final String BTN_SEARCH = "search";
    public static final String BTN_HIGHLIGHT = "highlight";
    public static final String BTN_CLEAR_HIGHLIGHT = "clearHighlight";
    public static final String BTN_AUTO_MARK = "autoMark";
    public static final String BTN_SETTINGS = "settings";
    
    // 默认显示的按钮（按顺序）
    private List<String> visibleButtons;
    
    // 按钮样式配置
    private String buttonColor;
    private String buttonHoverColor;
    private String buttonTextColor;
    
    public ToolbarConfig() {
        properties = new Properties();
        loadDefaults();
        load();
    }
    
    /**
     * 加载默认配置
     */
    private void loadDefaults() {
        // 默认显示的按钮
        visibleButtons = new ArrayList<>(Arrays.asList(
            BTN_NEW,
            BTN_OPEN,
            BTN_SAVE,
            "separator",
            BTN_UNDO,
            "separator",
            BTN_ADD_ROW,
            BTN_ADD_COLUMN,
            BTN_DELETE_ROW,
            BTN_DELETE_COLUMN,
            "separator",
            BTN_SEARCH,
            BTN_HIGHLIGHT,
            BTN_CLEAR_HIGHLIGHT,
            BTN_AUTO_MARK
        ));
        
        // 默认按钮样式
        buttonColor = "#f0f0f0";
        buttonHoverColor = "#e0e0e0";
        buttonTextColor = "#333333";
    }
    
    /**
     * 从文件加载配置
     */
    public void load() {
        try (InputStream inputStream = ConfigManager.getConfigInputStream(ConfigManager.TOOLBAR_FILE)) {
            if (inputStream != null) {
                properties.load(inputStream);
                
                // 加载可见按钮列表
                String buttonsStr = properties.getProperty("visibleButtons");
                if (buttonsStr != null && !buttonsStr.isEmpty()) {
                    visibleButtons = new ArrayList<>(Arrays.asList(buttonsStr.split(",")));
                    
                    // 检查是否包含新添加的按钮，如果没有则自动添加
                    if (!visibleButtons.contains(BTN_AUTO_MARK)) {
                        // 在清除高亮按钮后添加自动标记按钮
                        int clearHighlightIndex = visibleButtons.indexOf(BTN_CLEAR_HIGHLIGHT);
                        if (clearHighlightIndex >= 0 && clearHighlightIndex < visibleButtons.size() - 1) {
                            visibleButtons.add(clearHighlightIndex + 1, BTN_AUTO_MARK);
                        } else {
                            // 如果找不到清除高亮按钮，就添加到末尾
                            visibleButtons.add(BTN_AUTO_MARK);
                        }
                        // 保存更新后的配置
                        save();
                    }
                }
                
                // 加载样式配置
                buttonColor = properties.getProperty("buttonColor", buttonColor);
                buttonHoverColor = properties.getProperty("buttonHoverColor", buttonHoverColor);
                buttonTextColor = properties.getProperty("buttonTextColor", buttonTextColor);
            } else {
                System.out.println("工具栏配置文件不存在，使用默认配置");
            }
        } catch (IOException e) {
            System.out.println("加载工具栏配置文件时出错，使用默认配置: " + e.getMessage());
        }
    }
    
    /**
     * 保存配置到文件
     */
    public void save() {
        // 保存可见按钮列表
        properties.setProperty("visibleButtons", String.join(",", visibleButtons));
        
        // 保存样式配置
        properties.setProperty("buttonColor", buttonColor);
        properties.setProperty("buttonHoverColor", buttonHoverColor);
        properties.setProperty("buttonTextColor", buttonTextColor);
        
        try (OutputStream os = ConfigManager.getConfigOutputStream(ConfigManager.TOOLBAR_FILE)) {
            properties.store(os, "CSV Editor Toolbar Configuration");
        } catch (IOException e) {
            System.err.println("无法保存工具栏配置文件: " + e.getMessage());
        }
    }
    
    /**
     * 获取所有可用的按钮ID
     */
    public static String[] getAllButtons() {
        return new String[]{
            BTN_NEW,
            BTN_OPEN,
            BTN_SAVE,
            BTN_SAVE_AS,
            BTN_UNDO,
            BTN_ADD_ROW,
            BTN_ADD_COLUMN,
            BTN_DELETE_ROW,
            BTN_DELETE_COLUMN,
            BTN_SEARCH,
            BTN_HIGHLIGHT,
            BTN_CLEAR_HIGHLIGHT,
            BTN_AUTO_MARK,
            BTN_SETTINGS
        };
    }
    
    /**
     * 获取按钮的显示名称
     */
    public static String getButtonDisplayName(String buttonId) {
        switch (buttonId) {
            case BTN_NEW: return "新建";
            case BTN_OPEN: return "打开";
            case BTN_SAVE: return "保存";
            case BTN_SAVE_AS: return "另存为";
            case BTN_UNDO: return "撤销";
            case BTN_ADD_ROW: return "添加行";
            case BTN_ADD_COLUMN: return "添加列";
            case BTN_DELETE_ROW: return "删除行";
            case BTN_DELETE_COLUMN: return "删除列";
            case BTN_SEARCH: return "搜索";
            case BTN_HIGHLIGHT: return "高亮";
            case BTN_CLEAR_HIGHLIGHT: return "清除高亮";
            case BTN_AUTO_MARK: return "自动标记";
            case BTN_SETTINGS: return "设置";
            default: return buttonId;
        }
    }
    
    // Getters and Setters
    public List<String> getVisibleButtons() {
        return new ArrayList<>(visibleButtons);
    }
    
    public void setVisibleButtons(List<String> buttons) {
        this.visibleButtons = new ArrayList<>(buttons);
    }
    
    public String getButtonColor() {
        return buttonColor;
    }
    
    public void setButtonColor(String buttonColor) {
        this.buttonColor = buttonColor;
    }
    
    public String getButtonHoverColor() {
        return buttonHoverColor;
    }
    
    public void setButtonHoverColor(String buttonHoverColor) {
        this.buttonHoverColor = buttonHoverColor;
    }
    
    public String getButtonTextColor() {
        return buttonTextColor;
    }
    
    public void setButtonTextColor(String buttonTextColor) {
        this.buttonTextColor = buttonTextColor;
    }
}

