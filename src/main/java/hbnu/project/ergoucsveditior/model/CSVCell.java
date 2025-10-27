package hbnu.project.ergoucsveditior.model;

import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;

/**
 * CSV单元格模型
 * 使用JavaFX的Property机制实现数据绑定
 */
public class CSVCell {
    private final StringProperty value;
    
    public CSVCell() {
        this("");
    }
    
    public CSVCell(String value) {
        this.value = new SimpleStringProperty(value);
    }
    
    public String getValue() {
        return value.get();
    }
    
    public void setValue(String value) {
        this.value.set(value);
    }
    
    public StringProperty valueProperty() {
        return value;
    }
}

