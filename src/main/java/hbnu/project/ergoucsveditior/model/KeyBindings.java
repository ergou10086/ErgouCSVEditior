package hbnu.project.ergoucsveditior.model;

import hbnu.project.ergoucsveditior.manager.ConfigManager;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyCodeCombination;
import javafx.scene.input.KeyCombination;

import java.io.*;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

/**
 * 快捷键绑定管理
 */
public class KeyBindings {
    private Properties properties;
    private Map<String, KeyCombination> bindings;
    
    // 操作名称常量
    public static final String ACTION_NEW = "new";
    public static final String ACTION_OPEN = "open";
    public static final String ACTION_SAVE = "save";
    public static final String ACTION_SAVE_AS = "saveAs";
    public static final String ACTION_UNDO = "undo";
    public static final String ACTION_ADD_ROW = "addRow";
    public static final String ACTION_ADD_COLUMN = "addColumn";
    public static final String ACTION_DELETE_ROW = "deleteRow";
    public static final String ACTION_DELETE_COLUMN = "deleteColumn";
    public static final String ACTION_EXIT = "exit";
    public static final String ACTION_CELL_NEWLINE = "cellNewline";
    public static final String ACTION_SEARCH = "search";
    public static final String ACTION_HIGHLIGHT = "highlight";
    public static final String ACTION_CLEAR_HIGHLIGHT = "clearHighlight";
    public static final String ACTION_COPY = "copy";
    public static final String ACTION_PASTE = "paste";
    public static final String ACTION_CLEAR_CELL = "clearCell";
    public static final String ACTION_EXPORT_CSV = "exportCSV";
    public static final String ACTION_ZOOM_TABLE = "zoomTable";  // 表格缩放（固定，不可修改）
    
    public KeyBindings() {
        properties = new Properties();
        bindings = new HashMap<>();
        loadDefaults();
        load();
    }
    
    /**
     * 加载默认快捷键
     */
    private void loadDefaults() {
        // 默认快捷键设置
        bindings.put(ACTION_NEW, new KeyCodeCombination(KeyCode.N, KeyCombination.CONTROL_DOWN));
        bindings.put(ACTION_OPEN, new KeyCodeCombination(KeyCode.O, KeyCombination.CONTROL_DOWN));
        bindings.put(ACTION_SAVE, new KeyCodeCombination(KeyCode.S, KeyCombination.CONTROL_DOWN));
        bindings.put(ACTION_SAVE_AS, new KeyCodeCombination(KeyCode.S, KeyCombination.CONTROL_DOWN, KeyCombination.SHIFT_DOWN));
        bindings.put(ACTION_UNDO, new KeyCodeCombination(KeyCode.Z, KeyCombination.CONTROL_DOWN));
        bindings.put(ACTION_CELL_NEWLINE, new KeyCodeCombination(KeyCode.ENTER, KeyCombination.CONTROL_DOWN));
        bindings.put(ACTION_SEARCH, new KeyCodeCombination(KeyCode.F, KeyCombination.CONTROL_DOWN));
        bindings.put(ACTION_HIGHLIGHT, new KeyCodeCombination(KeyCode.M, KeyCombination.CONTROL_DOWN));
        bindings.put(ACTION_CLEAR_HIGHLIGHT, new KeyCodeCombination(KeyCode.M, KeyCombination.CONTROL_DOWN, KeyCombination.SHIFT_DOWN));
        bindings.put(ACTION_COPY, new KeyCodeCombination(KeyCode.C, KeyCombination.CONTROL_DOWN));
        bindings.put(ACTION_PASTE, new KeyCodeCombination(KeyCode.V, KeyCombination.CONTROL_DOWN));
        bindings.put(ACTION_CLEAR_CELL, new KeyCodeCombination(KeyCode.DELETE));
        bindings.put(ACTION_EXPORT_CSV, new KeyCodeCombination(KeyCode.E, KeyCombination.CONTROL_DOWN));
        // 注意：表格缩放使用 Ctrl+滚轮，这是固定的，不可修改
        bindings.put(ACTION_ZOOM_TABLE, null); // 用null表示这是一个特殊的鼠标操作，不是键盘快捷键
        // 其他操作默认不设置快捷键
    }
    
    /**
     * 从文件加载快捷键配置
     */
    public void load() {
        try (InputStream is = ConfigManager.getConfigInputStream(ConfigManager.KEYBINDINGS_FILE)) {
            if (is != null) {
                properties.load(is);
                
                // 从配置文件读取并解析快捷键
                for (String action : properties.stringPropertyNames()) {
                    // 跳过固定的快捷键（如缩放）
                    if (!isModifiable(action)) {
                        continue;
                    }
                    
                    String keyString = properties.getProperty(action);
                    if (keyString != null && !keyString.trim().isEmpty()) {
                        try {
                            KeyCombination keyCombination = KeyCombination.valueOf(keyString);
                            bindings.put(action, keyCombination);
                        } catch (IllegalArgumentException e) {
                            // 忽略无效的快捷键配置
                        }
                    }
                }
            }
        } catch (IOException e) {
            // 加载失败，使用默认设置
        }
    }
    
    /**
     * 保存快捷键配置到文件
     */
    public void save() {
        for (Map.Entry<String, KeyCombination> entry : bindings.entrySet()) {
            // 跳过固定的快捷键（如缩放）
            if (!isModifiable(entry.getKey())) {
                continue;
            }
            
            if (entry.getValue() != null) {
                properties.setProperty(entry.getKey(), entry.getValue().toString());
            } else {
                properties.remove(entry.getKey());
            }
        }
        
        try (OutputStream os = ConfigManager.getConfigOutputStream(ConfigManager.KEYBINDINGS_FILE)) {
            properties.store(os, "CSV Editor Key Bindings");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    /**
     * 获取指定操作的快捷键
     */
    public KeyCombination getBinding(String action) {
        return bindings.get(action);
    }
    
    /**
     * 设置指定操作的快捷键
     */
    public void setBinding(String action, KeyCombination keyCombination) {
        bindings.put(action, keyCombination);
    }
    
    /**
     * 移除指定操作的快捷键
     */
    public void removeBinding(String action) {
        bindings.remove(action);
    }
    
    /**
     * 获取所有快捷键绑定
     */
    public Map<String, KeyCombination> getAllBindings() {
        return new HashMap<>(bindings);
    }
    
    /**
     * 获取操作的显示名称
     */
    public static String getActionDisplayName(String action) {
        switch (action) {
            case ACTION_NEW: return "新建";
            case ACTION_OPEN: return "打开";
            case ACTION_SAVE: return "保存";
            case ACTION_SAVE_AS: return "另存为";
            case ACTION_UNDO: return "撤销";
            case ACTION_ADD_ROW: return "添加行";
            case ACTION_ADD_COLUMN: return "添加列";
            case ACTION_DELETE_ROW: return "删除行";
            case ACTION_DELETE_COLUMN: return "删除列";
            case ACTION_EXIT: return "退出";
            case ACTION_CELL_NEWLINE: return "单元格内换行";
            case ACTION_SEARCH: return "搜索";
            case ACTION_HIGHLIGHT: return "标记高亮";
            case ACTION_CLEAR_HIGHLIGHT: return "清除高亮";
            case ACTION_COPY: return "复制";
            case ACTION_PASTE: return "粘贴";
            case ACTION_CLEAR_CELL: return "清除单元格";
            case ACTION_EXPORT_CSV: return "导出为CSV";
            case ACTION_ZOOM_TABLE: return "表格缩放（Ctrl+滚轮）";
            default: return action;
        }
    }
    
    /**
     * 获取所有可绑定的操作
     */
    public static String[] getAllActions() {
        return new String[] {
            ACTION_ZOOM_TABLE,  // 放在第一位，因为只允许调整第一个快捷键
            ACTION_NEW,
            ACTION_OPEN,
            ACTION_SAVE,
            ACTION_SAVE_AS,
            ACTION_UNDO,
            ACTION_ADD_ROW,
            ACTION_ADD_COLUMN,
            ACTION_DELETE_ROW,
            ACTION_DELETE_COLUMN,
            ACTION_EXIT,
            ACTION_CELL_NEWLINE,
            ACTION_SEARCH,
            ACTION_HIGHLIGHT,
            ACTION_CLEAR_HIGHLIGHT,
            ACTION_COPY,
            ACTION_PASTE,
            ACTION_CLEAR_CELL,
            ACTION_EXPORT_CSV
        };
    }
    
    /**
     * 判断某个快捷键是否可以修改
     */
    public static boolean isModifiable(String action) {
        // 表格缩放快捷键固定为Ctrl+滚轮，不允许修改
        return !ACTION_ZOOM_TABLE.equals(action);
    }
}

