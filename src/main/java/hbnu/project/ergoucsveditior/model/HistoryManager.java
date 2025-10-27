package hbnu.project.ergoucsveditior.model;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

import java.util.Stack;

/**
 * 历史记录管理器
 * 用于实现撤销功能
 */
public class HistoryManager {
    private Stack<CSVSnapshot> undoStack;
    private int maxHistorySize;
    
    public HistoryManager() {
        this(50); // 默认保存50个历史状态
    }
    
    public HistoryManager(int maxHistorySize) {
        this.maxHistorySize = maxHistorySize;
        this.undoStack = new Stack<>();
    }
    
    /**
     * 保存当前状态
     */
    public void saveState(CSVData csvData) {
        // 如果栈满了，移除最旧的记录
        if (undoStack.size() >= maxHistorySize) {
            undoStack.remove(0);
        }
        
        CSVSnapshot snapshot = new CSVSnapshot(csvData);
        undoStack.push(snapshot);
    }
    
    /**
     * 撤销到上一个状态
     */
    public CSVData undo() {
        if (canUndo()) {
            // 移除当前状态
            undoStack.pop();
            
            // 返回上一个状态
            if (!undoStack.isEmpty()) {
                CSVSnapshot snapshot = undoStack.peek();
                return snapshot.restore();
            }
        }
        return null;
    }
    
    /**
     * 是否可以撤销
     */
    public boolean canUndo() {
        return undoStack.size() > 1; // 至少要有两个状态才能撤销
    }
    
    /**
     * 清空历史记录
     */
    public void clear() {
        undoStack.clear();
    }
    
    /**
     * 获取历史记录数量
     */
    public int getHistorySize() {
        return undoStack.size();
    }
    
    /**
     * CSV数据快照
     */
    private static class CSVSnapshot {
        private final ObservableList<ObservableList<CSVCell>> data;
        
        public CSVSnapshot(CSVData csvData) {
            
            // 深拷贝数据
            this.data = FXCollections.observableArrayList();
            for (ObservableList<CSVCell> row : csvData.getData()) {
                ObservableList<CSVCell> newRow = FXCollections.observableArrayList();
                for (CSVCell cell : row) {
                    newRow.add(new CSVCell(cell.getValue()));
                }
                this.data.add(newRow);
            }
        }
        
        public CSVData restore() {
            CSVData csvData = new CSVData();
            
            // 深拷贝恢复数据
            ObservableList<ObservableList<CSVCell>> restoredData = FXCollections.observableArrayList();
            for (ObservableList<CSVCell> row : this.data) {
                ObservableList<CSVCell> newRow = FXCollections.observableArrayList();
                for (CSVCell cell : row) {
                    newRow.add(new CSVCell(cell.getValue()));
                }
                restoredData.add(newRow);
            }
            
            csvData.setData(restoredData);
            return csvData;
        }
    }
}

