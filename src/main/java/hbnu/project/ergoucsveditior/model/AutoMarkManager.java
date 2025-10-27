package hbnu.project.ergoucsveditior.model;

import java.util.*;
import java.util.regex.Pattern;
import java.util.regex.PatternSyntaxException;

/**
 * 自动标记管理器
 * 负责应用和验证自动标记规则
 */
public class AutoMarkManager {
    private List<AutoMarkRule> rules;
    private Map<String, String> autoMarkColors; // key: "row,col", value: color
    
    public AutoMarkManager() {
        this.rules = new ArrayList<>();
        this.autoMarkColors = new HashMap<>();
    }
    
    /**
     * 添加规则
     */
    public void addRule(AutoMarkRule rule) {
        rules.add(rule);
    }
    
    /**
     * 移除规则
     */
    public void removeRule(AutoMarkRule rule) {
        rules.remove(rule);
    }
    
    /**
     * 获取所有规则
     */
    public List<AutoMarkRule> getRules() {
        return new ArrayList<>(rules);
    }
    
    /**
     * 清除所有规则
     */
    public void clearRules() {
        rules.clear();
        autoMarkColors.clear();
    }
    
    /**
     * 应用所有规则到CSV数据
     */
    public void applyRules(CSVData csvData) {
        autoMarkColors.clear();
        
        for (AutoMarkRule rule : rules) {
            if (!rule.isEnabled()) {
                continue;
            }
            
            applyRule(rule, csvData);
        }
    }
    
    /**
     * 应用单个规则
     */
    private void applyRule(AutoMarkRule rule, CSVData csvData) {
        for (int row = 0; row < csvData.getRows(); row++) {
            for (int col = 0; col < csvData.getColumns(); col++) {
                // 检查应用范围
                if (!isInScope(rule, col)) {
                    continue;
                }
                
                String value = csvData.getCellValue(row, col);
                if (matchesRule(rule, value)) {
                    String key = row + "," + col;
                    autoMarkColors.put(key, rule.getColor());
                }
            }
        }
    }
    
    /**
     * 检查列是否在规则应用范围内
     */
    private boolean isInScope(AutoMarkRule rule, int col) {
        switch (rule.getScope()) {
            case ALL_COLUMNS:
                return true;
            case SPECIFIED_COLUMNS:
                if (rule.getSpecifiedColumns() == null) {
                    return false;
                }
                for (int specCol : rule.getSpecifiedColumns()) {
                    if (specCol == col) {
                        return true;
                    }
                }
                return false;
            case SELECTED_COLUMN:
                // 这个需要从外部传入选中列
                return true;
            default:
                return false;
        }
    }
    
    /**
     * 检查值是否匹配规则
     */
    public boolean matchesRule(AutoMarkRule rule, String value) {
        if (rule == null) {
            return false;
        }
        
        switch (rule.getType()) {
            case NUMBER_GREATER:
                return checkNumberGreater(value, rule.getParameter());
            case NUMBER_LESS:
                return checkNumberLess(value, rule.getParameter());
            case NUMBER_EQUAL:
                return checkNumberEqual(value, rule.getParameter());
            case NUMBER_PRIME:
                return isPrime(value);
            case STRING_CONTAINS:
                return checkStringContains(value, rule.getParameter());
            case STRING_REGEX:
                return checkRegex(value, rule.getParameter());
            case FORMAT_EMAIL:
                return !isEmail(value);
            case FORMAT_PHONE:
                return !isPhone(value);
            case FORMAT_URL:
                return !isURL(value);
            case FORMAT_ID_CARD:
                return !isIDCard(value);
            case EMPTY_NULL:
                return isEmpty(value);
            case EMPTY_WHITESPACE:
                return isWhitespaceOnly(value);
            case EMPTY_ZERO_LENGTH:
                return isZeroLength(value);
            default:
                return false;
        }
    }
    
    /**
     * 获取单元格的自动标记颜色
     */
    public String getAutoMarkColor(int row, int col) {
        return autoMarkColors.get(row + "," + col);
    }
    
    /**
     * 清除所有自动标记
     */
    public void clearAutoMarks() {
        autoMarkColors.clear();
    }
    
    /**
     * 清除单个单元格的自动标记
     */
    public void clearCellAutoMark(int row, int col) {
        String key = row + "," + col;
        autoMarkColors.remove(key);
    }
    
    /**
     * 清除整行的自动标记
     */
    public void clearRowAutoMark(int row, int columns) {
        for (int col = 0; col < columns; col++) {
            String key = row + "," + col;
            autoMarkColors.remove(key);
        }
    }
    
    /**
     * 清除整列的自动标记
     */
    public void clearColumnAutoMark(int col, int rows) {
        for (int row = 0; row < rows; row++) {
            String key = row + "," + col;
            autoMarkColors.remove(key);
        }
    }
    
    // ==================== 规则校验方法 ====================
    
    /**
     * 检查数字是否大于指定值
     */
    private boolean checkNumberGreater(String value, String parameter) {
        try {
            double num = Double.parseDouble(value);
            double threshold = Double.parseDouble(parameter);
            return num > threshold;
        } catch (NumberFormatException e) {
            return false;
        }
    }
    
    /**
     * 检查数字是否小于指定值
     */
    private boolean checkNumberLess(String value, String parameter) {
        try {
            double num = Double.parseDouble(value);
            double threshold = Double.parseDouble(parameter);
            return num < threshold;
        } catch (NumberFormatException e) {
            return false;
        }
    }
    
    /**
     * 检查数字是否等于指定值
     */
    private boolean checkNumberEqual(String value, String parameter) {
        try {
            double num = Double.parseDouble(value);
            double threshold = Double.parseDouble(parameter);
            return Math.abs(num - threshold) < 0.0001;
        } catch (NumberFormatException e) {
            return false;
        }
    }
    
    /**
     * 检查是否为质数
     */
    public boolean isPrime(String value) {
        try {
            long num = Long.parseLong(value);
            if (num <= 1) {
                return false;
            }
            if (num <= 3) {
                return true;
            }
            if (num % 2 == 0 || num % 3 == 0) {
                return false;
            }
            for (long i = 5; i * i <= num; i += 6) {
                if (num % i == 0 || num % (i + 2) == 0) {
                    return false;
                }
            }
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }
    
    /**
     * 检查字符串是否包含指定内容
     */
    private boolean checkStringContains(String value, String parameter) {
        if (value == null || parameter == null) {
            return false;
        }
        return value.contains(parameter);
    }
    
    /**
     * 检查字符串是否匹配正则表达式
     */
    private boolean checkRegex(String value, String parameter) {
        if (value == null || parameter == null) {
            return false;
        }
        try {
            return Pattern.matches(parameter, value);
        } catch (PatternSyntaxException e) {
            return false;
        }
    }
    
    /**
     * 检查是否为有效邮箱
     */
    public boolean isEmail(String value) {
        if (value == null || value.isEmpty()) {
            return false;
        }
        String emailRegex = "^[A-Za-z0-9+_.-]+@[A-Za-z0-9.-]+\\.[A-Za-z]{2,}$";
        return Pattern.matches(emailRegex, value);
    }
    
    /**
     * 检查是否为有效手机号
     */
    public boolean isPhone(String value) {
        if (value == null || value.isEmpty()) {
            return false;
        }
        // 国内11位手机号
        String chinaPhone = "^1[3-9]\\d{9}$";
        // 国际格式 +国家码+号码
        String internationalPhone = "^\\+\\d{1,3}\\d{10,11}$";
        return Pattern.matches(chinaPhone, value) || Pattern.matches(internationalPhone, value);
    }
    
    /**
     * 检查是否为有效URL
     */
    public boolean isURL(String value) {
        if (value == null || value.isEmpty()) {
            return false;
        }
        String urlRegex = "^(http|https)://[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}(/.*)?$";
        return Pattern.matches(urlRegex, value);
    }
    
    /**
     * 检查是否为有效身份证号
     */
    public boolean isIDCard(String value) {
        if (value == null || value.length() != 18) {
            return false;
        }
        // 简单验证：前17位是数字，最后一位是数字或X
        String idCardRegex = "^\\d{17}[0-9Xx]$";
        if (!Pattern.matches(idCardRegex, value)) {
            return false;
        }
        
        // 验证校验码
        try {
            int[] weights = {7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2};
            char[] checkCodes = {'1', '0', 'X', '9', '8', '7', '6', '5', '4', '3', '2'};
            
            int sum = 0;
            for (int i = 0; i < 17; i++) {
                sum += (value.charAt(i) - '0') * weights[i];
            }
            int mod = sum % 11;
            char checkCode = checkCodes[mod];
            
            return Character.toUpperCase(value.charAt(17)) == checkCode;
        } catch (Exception e) {
            return false;
        }
    }
    
    /**
     * 检查是否为空值
     */
    private boolean isEmpty(String value) {
        return value == null || value.isEmpty();
    }
    
    /**
     * 检查是否仅含空白字符
     */
    private boolean isWhitespaceOnly(String value) {
        if (value == null || value.isEmpty()) {
            return false;
        }
        return value.trim().isEmpty();
    }
    
    /**
     * 检查长度是否为0
     */
    private boolean isZeroLength(String value) {
        return value != null && value.length() == 0;
    }
}

