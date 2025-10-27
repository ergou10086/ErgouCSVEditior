package hbnu.project.ergoucsveditior.model;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

/**
 * CSV数据模型
 * 管理整个表格的数据结构
 */
public class CSVData {
    private ObservableList<ObservableList<CSVCell>> data;
    private int rows;
    private int columns;
    
    /**
     * 创建空的CSV数据
     */
    public CSVData() {
        this.data = FXCollections.observableArrayList();
        this.rows = 0;
        this.columns = 0;
    }
    
    /**
     * 创建指定大小的CSV数据
     * @param rows 行数
     * @param columns 列数
     */
    public CSVData(int rows, int columns) {
        this.data = FXCollections.observableArrayList();
        this.rows = rows;
        this.columns = columns;
        initializeData(rows, columns);
    }
    
    /**
     * 初始化数据表格
     */
    private void initializeData(int rows, int columns) {
        data.clear();
        for (int i = 0; i < rows; i++) {
            ObservableList<CSVCell> row = FXCollections.observableArrayList();
            for (int j = 0; j < columns; j++) {
                row.add(new CSVCell());
            }
            data.add(row);
        }
    }
    
    /**
     * 获取指定位置的单元格
     */
    public CSVCell getCell(int row, int column) {
        if (row >= 0 && row < rows && column >= 0 && column < columns) {
            return data.get(row).get(column);
        }
        return null;
    }
    
    /**
     * 设置指定位置的单元格值
     */
    public void setCellValue(int row, int column, String value) {
        CSVCell cell = getCell(row, column);
        if (cell != null) {
            cell.setValue(value);
        }
    }
    
    /**
     * 获取指定位置的单元格值
     */
    public String getCellValue(int row, int column) {
        CSVCell cell = getCell(row, column);
        return cell != null ? cell.getValue() : "";
    }
    
    /**
     * 添加新行
     */
    public void addRow() {
        ObservableList<CSVCell> row = FXCollections.observableArrayList();
        for (int i = 0; i < columns; i++) {
            row.add(new CSVCell());
        }
        data.add(row);
        rows++;
    }
    
    /**
     * 在指定位置插入新行
     */
    public void insertRow(int index) {
        if (index < 0 || index > rows) {
            addRow();
            return;
        }
        ObservableList<CSVCell> row = FXCollections.observableArrayList();
        for (int i = 0; i < columns; i++) {
            row.add(new CSVCell());
        }
        data.add(index, row);
        rows++;
    }
    
    /**
     * 添加新列
     */
    public void addColumn() {
        for (ObservableList<CSVCell> row : data) {
            row.add(new CSVCell());
        }
        columns++;
    }
    
    /**
     * 在指定位置插入新列
     */
    public void insertColumn(int index) {
        if (index < 0 || index > columns) {
            addColumn();
            return;
        }
        for (ObservableList<CSVCell> row : data) {
            row.add(index, new CSVCell());
        }
        columns++;
    }
    
    /**
     * 删除指定行
     */
    public void removeRow(int index) {
        if (index >= 0 && index < rows) {
            data.remove(index);
            rows--;
        }
    }
    
    /**
     * 删除指定列
     */
    public void removeColumn(int index) {
        if (index >= 0 && index < columns) {
            for (ObservableList<CSVCell> row : data) {
                row.remove(index);
            }
            columns--;
        }
    }
    
    /**
     * 重置数据为指定大小
     */
    public void resize(int newRows, int newColumns) {
        this.rows = newRows;
        this.columns = newColumns;
        initializeData(newRows, newColumns);
    }
    
    /**
     * 清空所有数据
     */
    public void clear() {
        data.clear();
        rows = 0;
        columns = 0;
    }

    /**
     * 清空所有数据,直接调用clear
     */
    public void clearData() {
        data.clear();
    }

    /**
     * 复制另一个CSVData对象的数据
     */
    public void copyFrom(CSVData other) {
        this.data.clear();
        for (ObservableList<CSVCell> row : other.getData()) {
            ObservableList<CSVCell> newRow = FXCollections.observableArrayList();
            for (CSVCell cell : row) {
                newRow.add(new CSVCell(cell.getValue()));
            }
            this.data.add(newRow);
        }
    }
    
    // Getters
    public ObservableList<ObservableList<CSVCell>> getData() {
        return data;
    }
    
    public int getRows() {
        return rows;
    }
    
    public int getColumns() {
        return columns;
    }
    
    public void setData(ObservableList<ObservableList<CSVCell>> data) {
        this.data = data;
        this.rows = data.size();
        this.columns = data.isEmpty() ? 0 : data.get(0).size();
    }
}

