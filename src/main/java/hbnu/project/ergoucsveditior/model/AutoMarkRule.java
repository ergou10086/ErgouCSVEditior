package hbnu.project.ergoucsveditior.model;

import java.io.Serializable;

/**
 * 自动标记规则
 */
public class AutoMarkRule implements Serializable {
    private static final long serialVersionUID = 1L;
    
    public enum RuleType {
        NUMBER_GREATER,      // 数字大于
        NUMBER_LESS,         // 数字小于
        NUMBER_EQUAL,        // 数字等于
        NUMBER_PRIME,        // 质数
        STRING_CONTAINS,     // 字符串包含
        STRING_REGEX,        // 正则表达式
        FORMAT_EMAIL,        // 邮箱格式
        FORMAT_PHONE,        // 手机号格式
        FORMAT_URL,          // URL格式
        FORMAT_ID_CARD,      // 身份证号格式
        EMPTY_NULL,          // 完全空值
        EMPTY_WHITESPACE,    // 仅含空白字符
        EMPTY_ZERO_LENGTH    // 长度为0
    }
    
    public enum ApplyScope {
        SELECTED_COLUMN,     // 当前选中列
        ALL_COLUMNS,         // 所有列
        SPECIFIED_COLUMNS    // 指定列
    }
    
    private String id;
    private String name;
    private RuleType type;
    private String parameter;        // 规则参数（如数值、字符串、正则表达式）
    private String color;            // 标记颜色（十六进制）
    private ApplyScope scope;
    private int[] specifiedColumns;  // 指定的列索引
    private boolean enabled;
    
    public AutoMarkRule() {
        this.id = java.util.UUID.randomUUID().toString();
        this.enabled = true;
        this.scope = ApplyScope.ALL_COLUMNS;
    }
    
    public AutoMarkRule(String name, RuleType type, String parameter, String color) {
        this();
        this.name = name;
        this.type = type;
        this.parameter = parameter;
        this.color = color;
    }
    
    // Getters and Setters
    public String getId() {
        return id;
    }
    
    public void setId(String id) {
        this.id = id;
    }
    
    public String getName() {
        return name;
    }
    
    public void setName(String name) {
        this.name = name;
    }
    
    public RuleType getType() {
        return type;
    }
    
    public void setType(RuleType type) {
        this.type = type;
    }
    
    public String getParameter() {
        return parameter;
    }
    
    public void setParameter(String parameter) {
        this.parameter = parameter;
    }
    
    public String getColor() {
        return color;
    }
    
    public void setColor(String color) {
        this.color = color;
    }
    
    public ApplyScope getScope() {
        return scope;
    }
    
    public void setScope(ApplyScope scope) {
        this.scope = scope;
    }
    
    public int[] getSpecifiedColumns() {
        return specifiedColumns;
    }
    
    public void setSpecifiedColumns(int[] specifiedColumns) {
        this.specifiedColumns = specifiedColumns;
    }
    
    public boolean isEnabled() {
        return enabled;
    }
    
    public void setEnabled(boolean enabled) {
        this.enabled = enabled;
    }
    
    @Override
    public String toString() {
        return name != null ? name : type.toString();
    }
}

