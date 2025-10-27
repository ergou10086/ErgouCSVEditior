package hbnu.project.ergoucsveditior.service;

import hbnu.project.ergoucsveditior.model.CSVCell;
import hbnu.project.ergoucsveditior.model.CSVData;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.csv.CSVRecord;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

/**
 * CSV文件服务
 * 使用 Apache Commons CSV 库处理 CSV 格式，避免手动解析的漏洞
 * 支持 LF 和 CRLF 换行符，保证跨平台兼容性
 */
public class CSVService {
    
    private String lineEnding = "\n"; // 默认使用LF
    private boolean autoDetectDelimiter = true; // 自动检测分隔符
    private String escapeMode = "重复引号"; // 转义模式：重复引号 或 反斜杠转义
    
    /**
     * 设置换行符类型
     */
    public void setLineEnding(String lineEnding) {
        this.lineEnding = lineEnding;
    }
    
    /**
     * 设置是否自动检测分隔符
     */
    public void setAutoDetectDelimiter(boolean autoDetectDelimiter) {
        this.autoDetectDelimiter = autoDetectDelimiter;
    }
    
    /**
     * 设置转义模式
     */
    public void setEscapeMode(String escapeMode) {
        this.escapeMode = escapeMode;
    }
    
    /**
     * 检测文件的分隔符
     * 根据文件扩展名或分析文件内容来确定分隔符
     */
    private char detectDelimiter(File file) throws IOException {
        String fileName = file.getName().toLowerCase();
        
        // 根据文件扩展名判断
        if (fileName.endsWith(".tsv")) {
            return '\t'; // Tab分隔
        } else if (fileName.endsWith(".psv")) {
            return '|'; // 管道符分隔
        } else if (fileName.endsWith(".csv")) {
            // 对于CSV文件，读取前几行分析
            return analyzeDelimiter(file);
        }
        
        // 默认返回逗号
        return ',';
    }
    
    /**
     * 分析文件内容确定分隔符
     * 读取前几行，统计不同分隔符出现的频率
     */
    private char analyzeDelimiter(File file) throws IOException {
        char[] possibleDelimiters = {',', '\t', ';', '|'};
        int[] counts = new int[possibleDelimiters.length];
        
        try (BufferedReader reader = new BufferedReader(
                new InputStreamReader(new FileInputStream(file), StandardCharsets.UTF_8))) {
            
            // 读取前5行进行分析
            int linesToAnalyze = 5;
            int lineCount = 0;
            String line;
            
            while ((line = reader.readLine()) != null && lineCount < linesToAnalyze) {
                // 统计每种分隔符出现的次数
                for (int i = 0; i < possibleDelimiters.length; i++) {
                    int count = line.length() - line.replace(String.valueOf(possibleDelimiters[i]), "").length();
                    counts[i] += count;
                }
                lineCount++;
            }
        }
        
        // 找出出现次数最多的分隔符
        int maxIndex = 0;
        for (int i = 1; i < counts.length; i++) {
            if (counts[i] > counts[maxIndex]) {
                maxIndex = i;
            }
        }
        
        return possibleDelimiters[maxIndex];
    }
    
    /**
     * 根据转义模式获取CSV格式
     */
    private CSVFormat getCSVFormat(char delimiter) {
        CSVFormat.Builder builder = CSVFormat.RFC4180.builder()
                .setDelimiter(delimiter);
        
        if ("反斜杠转义".equals(escapeMode)) {
            // 使用反斜杠作为转义字符
            builder.setEscape('\\');
            builder.setQuote('"');
        } else {
            // 默认使用重复引号转义（RFC4180标准）
            builder.setQuote('"');
            builder.setEscape(null); // RFC4180使用重复引号
        }
        
        return builder.build();
    }
    
    /**
     * 从文件加载CSV数据
     * 使用 Apache Commons CSV 进行安全的 CSV 解析
     * 自动检测并处理 LF/CRLF 换行符
     * 支持自动检测分隔符和自定义转义模式
     * 
     * @param file 文件对象
     * @return CSV数据对象
     * @throws IOException 读取文件异常
     */
    public CSVData loadFromFile(File file) throws IOException {
        CSVData csvData = new CSVData();
        ObservableList<ObservableList<CSVCell>> data = FXCollections.observableArrayList();
        
        // 检测分隔符
        char delimiter = autoDetectDelimiter ? detectDelimiter(file) : ',';
        
        // 获取CSV格式（根据转义模式）
        CSVFormat format = getCSVFormat(delimiter);
        
        // 使用 BufferedReader 读取，自动处理各种换行符
        try (BufferedReader reader = new BufferedReader(
                new InputStreamReader(new FileInputStream(file), StandardCharsets.UTF_8));
             CSVParser csvParser = format.parse(reader)) {
            
            int maxColumns = 0;
            
            // 使用 Apache Commons CSV 解析每一行
            for (CSVRecord csvRecord : csvParser) {
                ObservableList<CSVCell> row = FXCollections.observableArrayList();
                
                // 遍历当前行的所有列
                for (int i = 0; i < csvRecord.size(); i++) {
                    String value = csvRecord.get(i);
                    // 保留单元格内的换行符
                    row.add(new CSVCell(value != null ? value : ""));
                }
                
                data.add(row);
                maxColumns = Math.max(maxColumns, row.size());
            }
            
            // 确保所有行的列数相同（填充空单元格）
            for (ObservableList<CSVCell> row : data) {
                while (row.size() < maxColumns) {
                    row.add(new CSVCell(""));
                }
            }
            
            csvData.setData(data);
        }
        
        return csvData;
    }
    
    /**
     * 保存CSV数据到文件
     * 使用 Apache Commons CSV 进行安全的 CSV 生成
     * 使用设置的换行符类型 (LF/CRLF)
     * 
     * @param csvData CSV数据对象
     * @param file 文件对象
     * @throws IOException 写入文件异常
     */
    public void saveToFile(CSVData csvData, File file) throws IOException {
        // 获取CSV格式（使用逗号分隔符和转义模式）
        CSVFormat.Builder builder = CSVFormat.RFC4180.builder()
                .setDelimiter(',');
        
        // 设置转义模式
        if ("反斜杠转义".equals(escapeMode)) {
            builder.setEscape('\\');
            builder.setQuote('"');
        } else {
            builder.setQuote('"');
            builder.setEscape(null);
        }
        
        // 设置换行符
        if ("\r\n".equals(lineEnding)) {
            builder.setRecordSeparator("\r\n");
        } else {
            builder.setRecordSeparator("\n");
        }
        
        CSVFormat format = builder.build();
        
        try (Writer writer = new OutputStreamWriter(new FileOutputStream(file), StandardCharsets.UTF_8);
             CSVPrinter csvPrinter = new CSVPrinter(writer, format)) {
            
            ObservableList<ObservableList<CSVCell>> data = csvData.getData();
            
            // 遍历每一行
            for (ObservableList<CSVCell> row : data) {
                List<String> values = new ArrayList<>();
                
                // 收集当前行的所有值
                for (CSVCell cell : row) {
                    String value = cell.getValue();
                    // 保留单元格内的换行符
                    values.add(value != null ? value : "");
                }
                
                // 使用 CSVPrinter 自动处理转义和引号
                csvPrinter.printRecord(values);
            }
            
            // 刷新缓冲区确保所有数据都写入
            csvPrinter.flush();
        }
    }
}

