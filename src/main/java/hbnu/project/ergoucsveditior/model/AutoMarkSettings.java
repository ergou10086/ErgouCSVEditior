package hbnu.project.ergoucsveditior.model;

import java.io.*;
import java.util.*;

/**
 * 自动标记设置管理
 */
public class AutoMarkSettings {
    private static final String SETTINGS_FILE = "csv_editor_automark_settings.properties";
    private static final String RULES_FILE = "csv_editor_automark_rules.dat";
    
    private Properties properties;
    
    // 默认颜色设置
    private String numberMarkColor;
    private String stringMarkColor;
    private String formatMarkColor;
    private String emptyMarkColor;
    
    // 规则模板
    private List<AutoMarkRule> savedRules;
    
    public AutoMarkSettings() {
        properties = new Properties();
        savedRules = new ArrayList<>();
        loadDefaults();
        load();
    }
    
    /**
     * 加载默认设置
     */
    private void loadDefaults() {
        numberMarkColor = "#FFEB3B";  // 黄色
        stringMarkColor = "#4CAF50";  // 绿色
        formatMarkColor = "#F44336";  // 红色
        emptyMarkColor = "#9E9E9E";   // 灰色
    }
    
    /**
     * 从文件加载设置
     */
    public void load() {
        // 加载颜色设置
        File file = new File(SETTINGS_FILE);
        if (file.exists()) {
            try (FileInputStream fis = new FileInputStream(file)) {
                properties.load(fis);
                
                numberMarkColor = properties.getProperty("numberMarkColor", numberMarkColor);
                stringMarkColor = properties.getProperty("stringMarkColor", stringMarkColor);
                formatMarkColor = properties.getProperty("formatMarkColor", formatMarkColor);
                emptyMarkColor = properties.getProperty("emptyMarkColor", emptyMarkColor);
                
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        
        // 加载规则模板
        loadRules();
    }
    
    /**
     * 保存设置到文件
     */
    public void save() {
        // 保存颜色设置
        properties.setProperty("numberMarkColor", numberMarkColor);
        properties.setProperty("stringMarkColor", stringMarkColor);
        properties.setProperty("formatMarkColor", formatMarkColor);
        properties.setProperty("emptyMarkColor", emptyMarkColor);
        
        try (FileOutputStream fos = new FileOutputStream(SETTINGS_FILE)) {
            properties.store(fos, "CSV Editor Auto Mark Settings");
        } catch (IOException e) {
            e.printStackTrace();
        }
        
        // 保存规则模板
        saveRules();
    }
    
    /**
     * 加载规则模板
     */
    @SuppressWarnings("unchecked")
    private void loadRules() {
        File file = new File(RULES_FILE);
        if (file.exists()) {
            try (ObjectInputStream ois = new ObjectInputStream(new FileInputStream(file))) {
                savedRules = (List<AutoMarkRule>) ois.readObject();
            } catch (IOException | ClassNotFoundException e) {
                savedRules = new ArrayList<>();
            }
        }
    }
    
    /**
     * 保存规则模板
     */
    private void saveRules() {
        try (ObjectOutputStream oos = new ObjectOutputStream(new FileOutputStream(RULES_FILE))) {
            oos.writeObject(savedRules);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    /**
     * 添加规则模板
     */
    public void addRuleTemplate(AutoMarkRule rule) {
        savedRules.add(rule);
        saveRules();
    }
    
    /**
     * 移除规则模板
     */
    public void removeRuleTemplate(AutoMarkRule rule) {
        savedRules.remove(rule);
        saveRules();
    }
    
    /**
     * 获取所有规则模板
     */
    public List<AutoMarkRule> getRuleTemplates() {
        return new ArrayList<>(savedRules);
    }
    
    /**
     * 获取规则类型的默认颜色
     */
    public String getDefaultColorForRuleType(AutoMarkRule.RuleType type) {
        if (type == null) {
            return numberMarkColor;
        }
        
        switch (type) {
            case NUMBER_GREATER:
            case NUMBER_LESS:
            case NUMBER_EQUAL:
            case NUMBER_PRIME:
                return numberMarkColor;
                
            case STRING_CONTAINS:
            case STRING_REGEX:
                return stringMarkColor;
                
            case FORMAT_EMAIL:
            case FORMAT_PHONE:
            case FORMAT_URL:
            case FORMAT_ID_CARD:
                return formatMarkColor;
                
            case EMPTY_NULL:
            case EMPTY_WHITESPACE:
            case EMPTY_ZERO_LENGTH:
                return emptyMarkColor;
                
            default:
                return numberMarkColor;
        }
    }
    
    // Getters and Setters
    public String getNumberMarkColor() {
        return numberMarkColor;
    }
    
    public void setNumberMarkColor(String numberMarkColor) {
        this.numberMarkColor = numberMarkColor;
    }
    
    public String getStringMarkColor() {
        return stringMarkColor;
    }
    
    public void setStringMarkColor(String stringMarkColor) {
        this.stringMarkColor = stringMarkColor;
    }
    
    public String getFormatMarkColor() {
        return formatMarkColor;
    }
    
    public void setFormatMarkColor(String formatMarkColor) {
        this.formatMarkColor = formatMarkColor;
    }
    
    public String getEmptyMarkColor() {
        return emptyMarkColor;
    }
    
    public void setEmptyMarkColor(String emptyMarkColor) {
        this.emptyMarkColor = emptyMarkColor;
    }
}

