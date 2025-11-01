package hbnu.project.ergoucsveditior.manager;

import hbnu.project.ergoucsveditior.model.HighlightInfo;
import javafx.scene.paint.Color;
import java.util.HashMap;
import java.util.Map;

/**
 * 高亮管理器
 * 管理所有单元格、行、列的高亮状态
 */
public class HighlightManager {
    // 单元格高亮 - key: "row,col"
    private Map<String, HighlightInfo> cellHighlights;
    
    // 行高亮 - key: row
    private Map<Integer, HighlightInfo> rowHighlights;
    
    // 列高亮 - key: col
    private Map<Integer, HighlightInfo> columnHighlights;
    
    // 默认颜色配置
    private Color defaultCellHighlightColor;
    private Color defaultRowHighlightColor;
    private Color defaultColumnHighlightColor;
    private Color defaultTextHighlightColor;
    private Color defaultSearchHighlightColor;
    
    // 冲突解决策略
    public enum ConflictStrategy {
        覆盖策略,    // 后标记覆盖先标记
        随机策略     // 随机选择一种颜色
    }
    
    private ConflictStrategy conflictStrategy;
    
    // 存储冲突单元格的随机选择结果
    private Map<String, Color> conflictResolutions;
    
    public HighlightManager() {
        cellHighlights = new HashMap<>();
        rowHighlights = new HashMap<>();
        columnHighlights = new HashMap<>();
        conflictResolutions = new HashMap<>();
        
        // 设置默认颜色
        defaultCellHighlightColor = Color.rgb(255, 255, 153, 0.7);      // 浅黄色
        defaultRowHighlightColor = Color.rgb(173, 216, 230, 0.5);       // 浅蓝色
        defaultColumnHighlightColor = Color.rgb(144, 238, 144, 0.5);    // 浅绿色
        defaultTextHighlightColor = Color.rgb(255, 215, 0, 0.8);        // 金黄色
        defaultSearchHighlightColor = Color.rgb(255, 165, 0, 0.6);      // 橙色
        
        // 默认使用覆盖策略
        conflictStrategy = ConflictStrategy.覆盖策略;
    }
    
    /**
     * 设置单元格高亮
     */
    public void setCellHighlight(int row, int col, Color color) {
        setCellHighlight(row, col, color, null);
    }
    
    /**
     * 设置单元格高亮（带文本颜色）
     */
    public void setCellHighlight(int row, int col, Color backgroundColor, Color textColor) {
        String key = row + "," + col;
        cellHighlights.put(key, new HighlightInfo(HighlightInfo.HighlightType.CELL, backgroundColor, textColor));
    }
    
    /**
     * 设置行高亮
     */
    public void setRowHighlight(int row, Color color) {
        setRowHighlight(row, color, null);
    }
    
    /**
     * 设置行高亮（带文本颜色）
     */
    public void setRowHighlight(int row, Color backgroundColor, Color textColor) {
        rowHighlights.put(row, new HighlightInfo(HighlightInfo.HighlightType.ROW, backgroundColor, textColor));
    }
    
    /**
     * 设置列高亮
     */
    public void setColumnHighlight(int col, Color color) {
        setColumnHighlight(col, color, null);
    }
    
    /**
     * 设置列高亮（带文本颜色）
     */
    public void setColumnHighlight(int col, Color backgroundColor, Color textColor) {
        columnHighlights.put(col, new HighlightInfo(HighlightInfo.HighlightType.COLUMN, backgroundColor, textColor));
    }
    
    /**
     * 获取单元格高亮信息
     */
    public HighlightInfo getCellHighlight(int row, int col) {
        String key = row + "," + col;
        return cellHighlights.get(key);
    }
    
    /**
     * 获取行高亮信息
     */
    public HighlightInfo getRowHighlight(int row) {
        return rowHighlights.get(row);
    }
    
    /**
     * 获取列高亮信息
     */
    public HighlightInfo getColumnHighlight(int col) {
        return columnHighlights.get(col);
    }
    
    /**
     * 清除单元格高亮
     */
    public void clearCellHighlight(int row, int col) {
        String key = row + "," + col;
        cellHighlights.remove(key);
    }
    
    /**
     * 清除行高亮
     */
    public void clearRowHighlight(int row) {
        rowHighlights.remove(row);
    }
    
    /**
     * 清除列高亮
     */
    public void clearColumnHighlight(int col) {
        columnHighlights.remove(col);
    }
    
    /**
     * 清除所有高亮
     */
    public void clearAllHighlights() {
        cellHighlights.clear();
        rowHighlights.clear();
        columnHighlights.clear();
        conflictResolutions.clear();
    }
    
    /**
     * 获取单元格的最终高亮背景颜色（优先级：单元格 > 行列冲突处理）
     */
    public Color getFinalHighlightColor(int row, int col) {
        // 优先级1：单元格高亮
        HighlightInfo cellHighlight = getCellHighlight(row, col);
        if (cellHighlight != null) {
            return cellHighlight.getBackgroundColor();
        }
        
        HighlightInfo rowHighlight = getRowHighlight(row);
        HighlightInfo colHighlight = getColumnHighlight(col);
        
        // 优先级2：处理行列高亮冲突
        if (rowHighlight != null && colHighlight != null) {
            return resolveConflict(row, col, rowHighlight, colHighlight);
        }
        
        // 优先级3：单独的行高亮或列高亮
        if (rowHighlight != null) {
            return rowHighlight.getBackgroundColor();
        }
        
        if (colHighlight != null) {
            return colHighlight.getBackgroundColor();
        }
        
        return null;
    }
    
    /**
     * 获取单元格的最终文本颜色（优先级：单元格 > 行 > 列）
     */
    public Color getFinalTextColor(int row, int col) {
        // 优先级1：单元格文本高亮
        HighlightInfo cellHighlight = getCellHighlight(row, col);
        if (cellHighlight != null && cellHighlight.getTextColor() != null) {
            return cellHighlight.getTextColor();
        }
        
        // 优先级2：行文本高亮
        HighlightInfo rowHighlight = getRowHighlight(row);
        if (rowHighlight != null && rowHighlight.getTextColor() != null) {
            return rowHighlight.getTextColor();
        }
        
        // 优先级3：列文本高亮
        HighlightInfo colHighlight = getColumnHighlight(col);
        if (colHighlight != null && colHighlight.getTextColor() != null) {
            return colHighlight.getTextColor();
        }
        
        return null;
    }
    
    /**
     * 解决行列高亮冲突
     */
    private Color resolveConflict(int row, int col, HighlightInfo rowHighlight, HighlightInfo colHighlight) {
        String key = row + "," + col;
        
        if (conflictStrategy == ConflictStrategy.覆盖策略) {
            // 覆盖策略：比较时间戳，返回较晚的颜色
            if (rowHighlight.getTimestamp() > colHighlight.getTimestamp()) {
                return rowHighlight.getColor();
            } else {
                return colHighlight.getColor();
            }
        } else {
            // 随机策略：首次冲突时随机选择，之后使用缓存结果
            if (!conflictResolutions.containsKey(key)) {
                // 随机选择行色或列色
                Color selectedColor = Math.random() < 0.5 ? 
                    rowHighlight.getColor() : colHighlight.getColor();
                conflictResolutions.put(key, selectedColor);
            }
            return conflictResolutions.get(key);
        }
    }
    
    /**
     * 清除冲突解决缓存
     */
    public void clearConflictResolutions() {
        conflictResolutions.clear();
    }
    
    // Getters and Setters for default colors
    public Color getDefaultCellHighlightColor() {
        return defaultCellHighlightColor;
    }
    
    public void setDefaultCellHighlightColor(Color color) {
        this.defaultCellHighlightColor = color;
    }
    
    public Color getDefaultRowHighlightColor() {
        return defaultRowHighlightColor;
    }
    
    public void setDefaultRowHighlightColor(Color color) {
        this.defaultRowHighlightColor = color;
    }
    
    public Color getDefaultColumnHighlightColor() {
        return defaultColumnHighlightColor;
    }
    
    public void setDefaultColumnHighlightColor(Color color) {
        this.defaultColumnHighlightColor = color;
    }
    
    public Color getDefaultTextHighlightColor() {
        return defaultTextHighlightColor;
    }
    
    public void setDefaultTextHighlightColor(Color color) {
        this.defaultTextHighlightColor = color;
    }
    
    public Color getDefaultSearchHighlightColor() {
        return defaultSearchHighlightColor;
    }
    
    public void setDefaultSearchHighlightColor(Color color) {
        this.defaultSearchHighlightColor = color;
    }
    
    public ConflictStrategy getConflictStrategy() {
        return conflictStrategy;
    }
    
    public void setConflictStrategy(ConflictStrategy strategy) {
        this.conflictStrategy = strategy;
        // 切换策略时清除冲突解决缓存
        clearConflictResolutions();
    }
    
    /**
     * 移动行时更新高亮信息
     * @param fromRow 源行索引
     * @param toRow 目标行索引
     */
    public void moveRow(int fromRow, int toRow) {
        // 1. 移动行高亮
        HighlightInfo rowHighlight = rowHighlights.remove(fromRow);
        if (rowHighlight != null) {
            rowHighlights.put(toRow, rowHighlight);
        }
        
        // 2. 移动单元格高亮
        Map<String, HighlightInfo> newCellHighlights = new HashMap<>();
        for (Map.Entry<String, HighlightInfo> entry : cellHighlights.entrySet()) {
            String key = entry.getKey();
            String[] parts = key.split(",");
            int row = Integer.parseInt(parts[0]);
            int col = Integer.parseInt(parts[1]);
            
            // 更新行索引
            if (row == fromRow) {
                // 被移动的行
                newCellHighlights.put(toRow + "," + col, entry.getValue());
            } else if (fromRow < toRow) {
                // 向下移动：中间的行向上移
                if (row > fromRow && row <= toRow) {
                    newCellHighlights.put((row - 1) + "," + col, entry.getValue());
                } else {
                    newCellHighlights.put(key, entry.getValue());
                }
            } else {
                // 向上移动：中间的行向下移
                if (row >= toRow && row < fromRow) {
                    newCellHighlights.put((row + 1) + "," + col, entry.getValue());
                } else {
                    newCellHighlights.put(key, entry.getValue());
                }
            }
        }
        cellHighlights = newCellHighlights;
        
        // 3. 更新中间行的行高亮
        if (fromRow < toRow) {
            // 向下移动
            for (int i = fromRow + 1; i <= toRow; i++) {
                HighlightInfo highlight = rowHighlights.remove(i);
                if (highlight != null) {
                    rowHighlights.put(i - 1, highlight);
                }
            }
        } else if (fromRow > toRow) {
            // 向上移动
            for (int i = toRow; i < fromRow; i++) {
                HighlightInfo highlight = rowHighlights.remove(i);
                if (highlight != null) {
                    rowHighlights.put(i + 1, highlight);
                }
            }
        }
        
        // 4. 清除冲突解决缓存（因为行号改变了）
        clearConflictResolutions();
    }
}

