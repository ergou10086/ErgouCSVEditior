package hbnu.project.ergoucsveditior.controller;

import hbnu.project.ergoucsveditior.model.CSVCell;
import hbnu.project.ergoucsveditior.model.CSVData;
import hbnu.project.ergoucsveditior.model.HighlightInfo;
import hbnu.project.ergoucsveditior.service.CSVService;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;
import java.io.IOException;
import java.util.Optional;

/**
 * 主控制器
 * 负责处理用户交互和协调Model与View
 */
public class MainController {
    
    @FXML
    private javafx.scene.layout.BorderPane rootPane;
    
    @FXML
    private TableView<ObservableList<CSVCell>> tableView;
    
    @FXML
    private Label statusLabel;
    
    @FXML
    private Label positionLabel;
    
    @FXML
    private Button undoButton;
    
    @FXML
    private ToolBar toolbar;
    
    // 背景图片视图
    private javafx.scene.layout.StackPane backgroundPane;
    private javafx.scene.image.ImageView backgroundImageView;
    
    private CSVData csvData;
    private CSVService csvService;
    private File currentFile;
    private hbnu.project.ergoucsveditior.model.HistoryManager historyManager;
    private hbnu.project.ergoucsveditior.model.Settings settings;
    private hbnu.project.ergoucsveditior.model.KeyBindings keyBindings;
    private hbnu.project.ergoucsveditior.model.ToolbarConfig toolbarConfig;
    private hbnu.project.ergoucsveditior.model.ExportSettings exportSettings;
    private boolean dataModified = false; // 标记数据是否被修改
    
    // 搜索相关
    private java.util.List<SearchResult> searchResults;
    private int currentSearchIndex = -1;
    private String lastSearchText = "";
    private boolean lastSearchCaseSensitive = false;
    
    // 高亮相关
    private hbnu.project.ergoucsveditior.model.HighlightManager highlightManager;
    
    // 自动标记相关
    private hbnu.project.ergoucsveditior.model.AutoMarkManager autoMarkManager;
    private hbnu.project.ergoucsveditior.model.AutoMarkSettings autoMarkSettings;
    
    // 剪贴板
    private String clipboardContent = "";
    
    /**
     * 初始化控制器
     */
    @FXML
    public void initialize() {
        csvData = new CSVData();
        csvService = new CSVService();
        settings = new hbnu.project.ergoucsveditior.model.Settings();
        historyManager = new hbnu.project.ergoucsveditior.model.HistoryManager(settings.getMaxHistorySize());
        keyBindings = new hbnu.project.ergoucsveditior.model.KeyBindings();
        highlightManager = new hbnu.project.ergoucsveditior.model.HighlightManager();
        autoMarkManager = new hbnu.project.ergoucsveditior.model.AutoMarkManager();
        autoMarkSettings = new hbnu.project.ergoucsveditior.model.AutoMarkSettings();
        toolbarConfig = new hbnu.project.ergoucsveditior.model.ToolbarConfig();
        exportSettings = new hbnu.project.ergoucsveditior.model.ExportSettings();
        
        // 从设置中加载高亮冲突策略
        String strategyName = settings.getHighlightConflictStrategy();
        if ("随机策略".equals(strategyName)) {
            highlightManager.setConflictStrategy(hbnu.project.ergoucsveditior.model.HighlightManager.ConflictStrategy.随机策略);
        } else {
            highlightManager.setConflictStrategy(hbnu.project.ergoucsveditior.model.HighlightManager.ConflictStrategy.覆盖策略);
        }
        
        // 初始化工具栏
        initializeToolbar();
        
        // 设置表格为可编辑
        tableView.setEditable(true);
        
        // 设置表格占位符
        tableView.setPlaceholder(new Label("请创建新表格或打开已有文件"));
        
        // 使用UNCONSTRAINED策略，我们将手动控制列宽
        tableView.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);
        
        // 设置键盘快捷键
        setupKeyboardShortcuts();
        
        // 设置窗口关闭事件处理
        javafx.application.Platform.runLater(() -> {
            Stage stage = getStage();
            if (stage != null) {
                stage.setOnCloseRequest(event -> {
                    if (!handleExit()) {
                        event.consume(); // 取消关闭
                    }
                });
            }
        });
        
        // 监听表格宽度变化，动态调整列宽
        tableView.widthProperty().addListener((obs, oldWidth, newWidth) -> {
            // 只在表格有列的情况下调整
            if (!tableView.getColumns().isEmpty()) {
                adjustColumnWidths();
            }
        });
        
        // 监听单元格选中事件，显示行列坐标
        tableView.getSelectionModel().selectedIndexProperty().addListener((obs, oldSelection, newSelection) -> {
            updatePositionLabel();
        });
        
        // 监听表格焦点变化，更新位置信息
        tableView.focusedProperty().addListener((obs, wasFocused, isNowFocused) -> {
            if (isNowFocused) {
                updatePositionLabel();
            }
        });
        
        // 设置Ctrl+鼠标滚轮缩放功能
        setupTableZoom();
        
        updateStatus("就绪");
        updatePositionLabel();
        updateUndoButton();
        
        // 应用初始主题
        applyInitialTheme();
    }
    
    /**
     * 初始化工具栏
     */
    private void initializeToolbar() {
        if (toolbar == null) {
            return;
        }
        
        toolbar.getItems().clear();
        
        // 根据配置动态添加按钮
        for (String buttonId : toolbarConfig.getVisibleButtons()) {
            if ("separator".equals(buttonId)) {
                toolbar.getItems().add(new Separator(javafx.geometry.Orientation.VERTICAL));
            } else {
                Button button = createToolbarButton(buttonId);
                if (button != null) {
                    toolbar.getItems().add(button);
                }
            }
        }
        
        // 应用按钮样式
        applyToolbarStyle();
    }
    
    /**
     * 创建工具栏按钮
     */
    private Button createToolbarButton(String buttonId) {
        Button button = new Button(hbnu.project.ergoucsveditior.model.ToolbarConfig.getButtonDisplayName(buttonId));
        
        // 绑定事件处理器
        switch (buttonId) {
            case hbnu.project.ergoucsveditior.model.ToolbarConfig.BTN_NEW:
                button.setOnAction(e -> handleNew());
                break;
            case hbnu.project.ergoucsveditior.model.ToolbarConfig.BTN_OPEN:
                button.setOnAction(e -> handleOpen());
                break;
            case hbnu.project.ergoucsveditior.model.ToolbarConfig.BTN_SAVE:
                button.setOnAction(e -> handleSave());
                break;
            case hbnu.project.ergoucsveditior.model.ToolbarConfig.BTN_SAVE_AS:
                button.setOnAction(e -> handleSaveAs());
                break;
            case hbnu.project.ergoucsveditior.model.ToolbarConfig.BTN_UNDO:
                button.setOnAction(e -> handleUndo());
                undoButton = button;
                break;
            case hbnu.project.ergoucsveditior.model.ToolbarConfig.BTN_ADD_ROW:
                button.setOnAction(e -> handleAddRow());
                break;
            case hbnu.project.ergoucsveditior.model.ToolbarConfig.BTN_ADD_COLUMN:
                button.setOnAction(e -> handleAddColumn());
                break;
            case hbnu.project.ergoucsveditior.model.ToolbarConfig.BTN_DELETE_ROW:
                button.setOnAction(e -> handleDeleteRow());
                break;
            case hbnu.project.ergoucsveditior.model.ToolbarConfig.BTN_DELETE_COLUMN:
                button.setOnAction(e -> handleDeleteColumn());
                break;
            case hbnu.project.ergoucsveditior.model.ToolbarConfig.BTN_SEARCH:
                button.setOnAction(e -> handleSearch());
                break;
            case hbnu.project.ergoucsveditior.model.ToolbarConfig.BTN_HIGHLIGHT:
                button.setOnAction(e -> handleHighlightCell());
                break;
            case hbnu.project.ergoucsveditior.model.ToolbarConfig.BTN_CLEAR_HIGHLIGHT:
                button.setOnAction(e -> handleClearHighlight());
                break;
            case hbnu.project.ergoucsveditior.model.ToolbarConfig.BTN_AUTO_MARK:
                button.setOnAction(e -> handleAutoMark());
                break;
            case hbnu.project.ergoucsveditior.model.ToolbarConfig.BTN_SETTINGS:
                button.setOnAction(e -> handleSettings());
                break;
            default:
                return null;
        }
        
        return button;
    }
    
    /**
     * 应用工具栏样式
     */
    private void applyToolbarStyle() {
        if (toolbar == null) {
            return;
        }
        
        String buttonColor = toolbarConfig.getButtonColor();
        String buttonHoverColor = toolbarConfig.getButtonHoverColor();
        String buttonTextColor = toolbarConfig.getButtonTextColor();
        
        for (javafx.scene.Node node : toolbar.getItems()) {
            if (node instanceof Button) {
                Button button = (Button) node;
                button.setStyle(String.format(
                    "-fx-background-color: %s; -fx-text-fill: %s; -fx-padding: 5 10 5 10;",
                    buttonColor, buttonTextColor
                ));
                
                // 添加悬停效果
                button.setOnMouseEntered(e -> {
                    button.setStyle(String.format(
                        "-fx-background-color: %s; -fx-text-fill: %s; -fx-padding: 5 10 5 10;",
                        buttonHoverColor, buttonTextColor
                    ));
                });
                
                button.setOnMouseExited(e -> {
                    button.setStyle(String.format(
                        "-fx-background-color: %s; -fx-text-fill: %s; -fx-padding: 5 10 5 10;",
                        buttonColor, buttonTextColor
                    ));
                });
            }
        }
    }
    
    /**
     * 设置键盘快捷键
     */
    private void setupKeyboardShortcuts() {
        javafx.application.Platform.runLater(() -> {
            Stage stage = getStage();
            if (stage != null && stage.getScene() != null) {
                stage.getScene().setOnKeyPressed(event -> {
                    // 检查各个操作的快捷键
                    if (matchesBinding(event, hbnu.project.ergoucsveditior.model.KeyBindings.ACTION_NEW)) {
                        handleNew();
                        event.consume();
                    } else if (matchesBinding(event, hbnu.project.ergoucsveditior.model.KeyBindings.ACTION_OPEN)) {
                        handleOpen();
                        event.consume();
                    } else if (matchesBinding(event, hbnu.project.ergoucsveditior.model.KeyBindings.ACTION_SAVE)) {
                        handleSave();
                        event.consume();
                    } else if (matchesBinding(event, hbnu.project.ergoucsveditior.model.KeyBindings.ACTION_SAVE_AS)) {
                        handleSaveAs();
                        event.consume();
                    } else if (matchesBinding(event, hbnu.project.ergoucsveditior.model.KeyBindings.ACTION_UNDO)) {
                        handleUndo();
                        event.consume();
                    } else if (matchesBinding(event, hbnu.project.ergoucsveditior.model.KeyBindings.ACTION_ADD_ROW)) {
                        handleAddRow();
                        event.consume();
                    } else if (matchesBinding(event, hbnu.project.ergoucsveditior.model.KeyBindings.ACTION_ADD_COLUMN)) {
                        handleAddColumn();
                        event.consume();
                    } else if (matchesBinding(event, hbnu.project.ergoucsveditior.model.KeyBindings.ACTION_DELETE_ROW)) {
                        handleDeleteRow();
                        event.consume();
                    } else if (matchesBinding(event, hbnu.project.ergoucsveditior.model.KeyBindings.ACTION_DELETE_COLUMN)) {
                        handleDeleteColumn();
                        event.consume();
                    } else if (matchesBinding(event, hbnu.project.ergoucsveditior.model.KeyBindings.ACTION_EXIT)) {
                        handleExit();
                        event.consume();
                    } else if (matchesBinding(event, hbnu.project.ergoucsveditior.model.KeyBindings.ACTION_SEARCH)) {
                        handleSearch();
                        event.consume();
                    } else if (matchesBinding(event, hbnu.project.ergoucsveditior.model.KeyBindings.ACTION_HIGHLIGHT)) {
                        handleHighlightCell();
                        event.consume();
                    } else if (matchesBinding(event, hbnu.project.ergoucsveditior.model.KeyBindings.ACTION_CLEAR_HIGHLIGHT)) {
                        handleClearHighlight();
                        event.consume();
                    } else if (matchesBinding(event, hbnu.project.ergoucsveditior.model.KeyBindings.ACTION_COPY)) {
                        handleCopy();
                        event.consume();
                    } else if (matchesBinding(event, hbnu.project.ergoucsveditior.model.KeyBindings.ACTION_PASTE)) {
                        handlePaste();
                        event.consume();
                    } else if (matchesBinding(event, hbnu.project.ergoucsveditior.model.KeyBindings.ACTION_CLEAR_CELL)) {
                        handleClearCell();
                        event.consume();
                    } else if (matchesBinding(event, hbnu.project.ergoucsveditior.model.KeyBindings.ACTION_EXPORT_CSV)) {
                        handleExportToCSV();
                        event.consume();
                    }
                });
            }
        });
    }
    
    /**
     * 检查事件是否匹配快捷键绑定
     */
    private boolean matchesBinding(javafx.scene.input.KeyEvent event, String action) {
        javafx.scene.input.KeyCombination binding = keyBindings.getBinding(action);
        return binding != null && binding.match(event);
    }
    
    /**
     * 创建新表格
     */
    @FXML
    public void handleNew() {
        // 如果当前已有数据，在新窗口打开
        if (csvData.getRows() > 0) {
            openNewWindow();
            return;
        }
        
        // 创建对话框让用户输入行列数
        Dialog<int[]> dialog = new Dialog<>();
        dialog.setTitle("新建表格");
        dialog.setHeaderText("请输入表格大小");
        
        // 设置按钮
        ButtonType createButtonType = new ButtonType("创建", ButtonBar.ButtonData.OK_DONE);
        dialog.getDialogPane().getButtonTypes().addAll(createButtonType, ButtonType.CANCEL);
        
        // 创建输入字段
        javafx.scene.layout.GridPane grid = new javafx.scene.layout.GridPane();
        grid.setHgap(10);
        grid.setVgap(10);
        grid.setPadding(new javafx.geometry.Insets(20, 150, 10, 10));
        
        TextField rowsField = new TextField("10");
        TextField columnsField = new TextField("5");
        
        grid.add(new Label("行数:"), 0, 0);
        grid.add(rowsField, 1, 0);
        grid.add(new Label("列数:"), 0, 1);
        grid.add(columnsField, 1, 1);
        
        dialog.getDialogPane().setContent(grid);
        
        // 转换结果
        dialog.setResultConverter(dialogButton -> {
            if (dialogButton == createButtonType) {
                try {
                    int rows = Integer.parseInt(rowsField.getText());
                    int columns = Integer.parseInt(columnsField.getText());
                    if (rows > 0 && columns > 0 && rows <= 1000 && columns <= 100) {
                        return new int[]{rows, columns};
                    } else {
                        javafx.application.Platform.runLater(() -> 
                            showError("输入错误", "行数和列数必须在合理范围内（行：1-1000，列：1-100）")
                        );
                    }
                } catch (NumberFormatException e) {
                    javafx.application.Platform.runLater(() -> 
                        showError("输入错误", "请输入有效的数字")
                    );
                }
            }
            return null;
        });
        
        Optional<int[]> result = dialog.showAndWait();
        result.ifPresent(size -> {
            csvData = new CSVData(size[0], size[1]);
            currentFile = null;
            historyManager.clear();
            saveHistory();
            dataModified = false;
            refreshTable();
            // 应用背景图片
            javafx.application.Platform.runLater(() -> applyBackgroundImage());
            updateStatus("已创建新表格: " + size[0] + "行 x " + size[1] + "列");
        });
    }
    
    /**
     * 打开CSV文件
     */
    @FXML
    public void handleOpen() {
        // 如果当前有未保存的数据，提示用户
        if (csvData.getRows() > 0 && !confirmDiscardChanges()) {
            return;
        }
        
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("打开CSV文件");
        fileChooser.getExtensionFilters().addAll(
            new FileChooser.ExtensionFilter("CSV文件", "*.csv"),
            new FileChooser.ExtensionFilter("所有文件", "*.*")
        );
        
        File file = fileChooser.showOpenDialog(getStage());
        if (file != null) {
            try {
                csvService.setLineEnding(settings.getLineEndingString());
                csvService.setAutoDetectDelimiter(settings.isAutoDetectDelimiter());
                csvService.setEscapeMode(settings.getEscapeMode());
                csvData = csvService.loadFromFile(file);
                currentFile = file;
                historyManager.clear();
                saveHistory();
                dataModified = false;
                refreshTable();
                // 应用背景图片
                javafx.application.Platform.runLater(() -> applyBackgroundImage());
                updateStatus("已打开文件: " + file.getName() + " (" + 
                           csvData.getRows() + "行 x " + csvData.getColumns() + "列)");
            } catch (IOException e) {
                showError("打开文件失败", 
                         "无法读取文件，请检查文件是否为有效的CSV格式。\n\n错误信息: " + e.getMessage());
            } catch (Exception e) {
                showError("打开文件失败", 
                         "处理CSV文件时发生错误。\n\n错误信息: " + e.getMessage());
            }
        }
    }
    
    /**
     * 保存CSV文件
     */
    @FXML
    public void handleSave() {
        if (currentFile != null) {
            saveToFile(currentFile);
        } else {
            handleSaveAs();
        }
    }
    
    /**
     * 另存为CSV文件
     */
    @FXML
    public void handleSaveAs() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("保存CSV文件");
        fileChooser.getExtensionFilters().add(
            new FileChooser.ExtensionFilter("CSV文件", "*.csv")
        );
        
        if (currentFile != null) {
            fileChooser.setInitialFileName(currentFile.getName());
        }
        
        File file = fileChooser.showSaveDialog(getStage());
        if (file != null) {
            saveToFile(file);
            currentFile = file;
        }
    }
    
    /**
     * 添加行（在选中位置前插入）
     */
    @FXML
    public void handleAddRow() {
        if (csvData.getColumns() == 0) {
            showInfo("提示", "请先创建表格或打开文件");
            return;
        }
        
        saveHistory();
        dataModified = true;
        
        int selectedIndex = tableView.getSelectionModel().getSelectedIndex();
        if (selectedIndex >= 0) {
            // 在选中行前插入
            csvData.insertRow(selectedIndex);
            refreshTable();
            updateStatus("已在第 " + (selectedIndex + 1) + " 行前添加新行");
        } else {
            // 没有选中行，添加到末尾
            csvData.addRow();
            refreshTable();
            updateStatus("已在末尾添加新行");
        }
    }
    
    /**
     * 添加列（在选中位置前插入）
     */
    @FXML
    public void handleAddColumn() {
        if (csvData.getRows() == 0) {
            showInfo("提示", "请先创建表格或打开文件");
            return;
        }
        
        saveHistory();
        dataModified = true;
        
        // 获取选中的列索引
        @SuppressWarnings("unchecked")
        TablePosition<ObservableList<CSVCell>, ?> focusedCell = 
            (TablePosition<ObservableList<CSVCell>, ?>) tableView.getFocusModel().getFocusedCell();
        
        if (focusedCell != null && focusedCell.getColumn() > 0) {
            // 减1是因为第一列是行号列
            int selectedColumn = focusedCell.getColumn() - 1;
            csvData.insertColumn(selectedColumn);
            refreshTable();
            updateStatus("已在第 " + (selectedColumn + 1) + " 列前添加新列");
        } else {
            // 没有选中列，添加到末尾
            csvData.addColumn();
            refreshTable();
            updateStatus("已在末尾添加新列");
        }
    }
    
    /**
     * 删除选中的行
     */
    @FXML
    public void handleDeleteRow() {
        int selectedIndex = tableView.getSelectionModel().getSelectedIndex();
        if (selectedIndex >= 0) {
            saveHistory();
            dataModified = true;
            csvData.removeRow(selectedIndex);
            refreshTable();
            updateStatus("已删除行 " + (selectedIndex + 1));
        } else {
            showInfo("提示", "请先选择要删除的行");
        }
    }
    
    /**
     * 删除选中的列
     */
    @FXML
    public void handleDeleteColumn() {
        @SuppressWarnings("unchecked")
        TablePosition<ObservableList<CSVCell>, ?> focusedCell = 
            (TablePosition<ObservableList<CSVCell>, ?>) tableView.getFocusModel().getFocusedCell();
        
        if (focusedCell != null && focusedCell.getColumn() > 0) {
            saveHistory();
            dataModified = true;
            // 减1是因为第一列是行号列
            int selectedColumn = focusedCell.getColumn() - 1;
            csvData.removeColumn(selectedColumn);
            refreshTable();
            updateStatus("已删除第 " + (selectedColumn + 1) + " 列");
        } else {
            showInfo("提示", "请先选择要删除的列");
        }
    }
    
    /**
     * 撤销操作
     */
    @FXML
    public void handleUndo() {
        CSVData previousState = historyManager.undo();
        if (previousState != null) {
            csvData = previousState;
            dataModified = true;
            refreshTable();
            updateStatus("已撤销");
        } else {
            showInfo("提示", "无法撤销");
        }
    }
    
    /**
     * 打开设置对话框
     */
    @FXML
    public void handleSettings() {
        showSettingsDialog();
    }
    
    /**
     * 打开快捷键设置对话框
     */
    @FXML
    public void handleKeyBindings() {
        showKeyBindingsDialog();
    }
    
    /**
     * 打开工具栏设置对话框
     */
    @FXML
    public void handleToolbarSettings() {
        showToolbarConfigDialog();
    }
    
    /**
     * 打开高亮设置对话框
     */
    @FXML
    public void handleHighlightSettings() {
        showHighlightSettingsDialog();
    }
    
    /**
     * 打开主题设置对话框
     */
    @FXML
    public void handleThemeSettings() {
        showThemeSettingsDialog();
    }
    
    /**
     * 打开自动标记工具对话框
     */
    @FXML
    public void handleAutoMark() {
        showAutoMarkDialog();
    }
    
    /**
     * 打开列统计计算对话框
     */
    @FXML
    public void handleColumnStatistics() {
        if (csvData.getRows() == 0) {
            showInfo("提示", "没有数据可以进行统计计算");
            return;
        }
        
        if (csvData.getColumns() == 0) {
            showInfo("提示", "没有列可以进行统计计算");
            return;
        }
        
        showColumnStatisticsDialog();
    }
    
    /**
     * 打开数据库持久化对话框
     */
    @FXML
    public void handleDatabasePersistence() {
        if (csvData.getRows() == 0) {
            showInfo("提示", "没有数据可以导入数据库");
            return;
        }
        
        if (csvData.getColumns() == 0) {
            showInfo("提示", "没有列可以导入数据库");
            return;
        }
        
        showDatabasePersistenceDialog();
    }
    
    /**
     * 导出为CSV（带进度条和文件检测）
     */
    @FXML
    public void handleExportToCSV() {
        if (csvData.getRows() == 0) {
            showInfo("提示", "没有数据可以导出");
            return;
        }
        
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("导出为CSV");
        fileChooser.getExtensionFilters().add(
            new FileChooser.ExtensionFilter("CSV文件", "*.csv")
        );
        
        if (currentFile != null) {
            fileChooser.setInitialFileName(currentFile.getName());
        }
        
        File file = fileChooser.showSaveDialog(getStage());
        if (file != null) {
            exportWithProgress(file, "CSV");
        }
    }
    
    /**
     * 导出为TXT
     */
    @FXML
    public void handleExportToTXT() throws Exception {
        if (csvData.getRows() == 0) {
            showInfo("提示", "没有数据可以导出");
            return;
        }
        
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("导出为TXT");
        fileChooser.getExtensionFilters().add(
            new FileChooser.ExtensionFilter("文本文件", "*.txt")
        );
        
        if (currentFile != null) {
            String name = currentFile.getName().replaceFirst("[.][^.]+$", "");
            fileChooser.setInitialFileName(name + ".txt");
        }
        
        File file = fileChooser.showSaveDialog(getStage());
        if (file != null) {
            exportToTXT(file);
        }
    }
    
    /**
     * 导出为HTML
     */
    @FXML
    public void handleExportToHTML() throws Exception {
        if (csvData.getRows() == 0) {
            showInfo("提示", "没有数据可以导出");
            return;
        }
        
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("导出为HTML");
        fileChooser.getExtensionFilters().add(
            new FileChooser.ExtensionFilter("HTML文件", "*.html")
        );
        
        if (currentFile != null) {
            String name = currentFile.getName().replaceFirst("[.][^.]+$", "");
            fileChooser.setInitialFileName(name + ".html");
        }
        
        File file = fileChooser.showSaveDialog(getStage());
        if (file != null) {
            exportToHTML(file);
        }
    }
    
    /**
     * 导出为Excel
     */
    @FXML
    public void handleExportToExcel() throws Exception {
        if (csvData.getRows() == 0) {
            showInfo("提示", "没有数据可以导出");
            return;
        }
        
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("导出为Excel");
        fileChooser.getExtensionFilters().add(
            new FileChooser.ExtensionFilter("Excel文件", "*.xlsx")
        );
        
        if (currentFile != null) {
            String name = currentFile.getName().replaceFirst("[.][^.]+$", "");
            fileChooser.setInitialFileName(name + ".xlsx");
        }
        
        File file = fileChooser.showSaveDialog(getStage());
        if (file != null) {
            exportToExcel(file);
        }
    }
    
    /**
     * 导出为PDF
     */
    @FXML
    public void handleExportToPDF() {
        if (csvData.getRows() == 0) {
            showInfo("提示", "没有数据可以导出");
            return;
        }
        
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("导出为PDF");
        fileChooser.getExtensionFilters().add(
            new FileChooser.ExtensionFilter("PDF文件", "*.pdf")
        );
        
        if (currentFile != null) {
            String name = currentFile.getName().replaceFirst("[.][^.]+$", "");
            fileChooser.setInitialFileName(name + ".pdf");
        }
        
        File file = fileChooser.showSaveDialog(getStage());
        if (file != null) {
            exportWithProgress(file, "PDF");
        }
    }
    
    /**
     * 导出为Markdown
     */
    @FXML
    public void handleExportToMarkdown() {
        if (csvData.getRows() == 0) {
            showInfo("提示", "没有数据可以导出");
            return;
        }
        
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("导出为Markdown");
        fileChooser.getExtensionFilters().add(
            new FileChooser.ExtensionFilter("Markdown文件", "*.md")
        );
        
        if (currentFile != null) {
            String name = currentFile.getName().replaceFirst("[.][^.]+$", "");
            fileChooser.setInitialFileName(name + ".md");
        }
        
        File file = fileChooser.showSaveDialog(getStage());
        if (file != null) {
            exportWithProgress(file, "Markdown");
        }
    }
    
    /**
     * 打开导出设置对话框
     */
    @FXML
    public void handleExportSettings() {
        showExportSettingsDialog();
    }
    
    /**
     * 打开搜索对话框
     */
    @FXML
    public void handleSearch() {
        showSearchDialog();
    }
    
    /**
     * 高亮选中的单元格
     */
    @FXML
    public void handleHighlightCell() {
        int selectedRow = tableView.getSelectionModel().getSelectedIndex();
        @SuppressWarnings("unchecked")
        TablePosition<ObservableList<CSVCell>, ?> focusedCell = 
            (TablePosition<ObservableList<CSVCell>, ?>) tableView.getFocusModel().getFocusedCell();
        
        if (selectedRow >= 0 && focusedCell != null && focusedCell.getColumn() > 0) {
            int selectedColumn = focusedCell.getColumn() - 1;
            showColorPickerDialog(selectedRow, selectedColumn, "cell");
        } else {
            showInfo("提示", "请先选择要高亮的单元格");
        }
    }
    
    /**
     * 清除高亮
     */
    @FXML
    public void handleClearHighlight() {
        int selectedRow = tableView.getSelectionModel().getSelectedIndex();
        @SuppressWarnings("unchecked")
        TablePosition<ObservableList<CSVCell>, ?> focusedCell = 
            (TablePosition<ObservableList<CSVCell>, ?>) tableView.getFocusModel().getFocusedCell();
        
        if (selectedRow >= 0 && focusedCell != null && focusedCell.getColumn() > 0) {
            int selectedColumn = focusedCell.getColumn() - 1;
            highlightManager.clearCellHighlight(selectedRow, selectedColumn);
            // 同时清除该单元格的自动标记
            clearAutoMarkForCell(selectedRow, selectedColumn);
            tableView.refresh();
            updateStatus("已清除单元格高亮");
        }
    }
    
    /**
     * 复制单元格内容
     */
    @FXML
    public void handleCopy() {
        int selectedRow = tableView.getSelectionModel().getSelectedIndex();
        @SuppressWarnings("unchecked")
        TablePosition<ObservableList<CSVCell>, ?> focusedCell = 
            (TablePosition<ObservableList<CSVCell>, ?>) tableView.getFocusModel().getFocusedCell();
        
        if (selectedRow >= 0 && focusedCell != null && focusedCell.getColumn() > 0) {
            int selectedColumn = focusedCell.getColumn() - 1;
            String cellValue = csvData.getCellValue(selectedRow, selectedColumn);
            clipboardContent = cellValue != null ? cellValue : "";
            
            // 也复制到系统剪贴板
            javafx.scene.input.ClipboardContent content = new javafx.scene.input.ClipboardContent();
            content.putString(clipboardContent);
            javafx.scene.input.Clipboard.getSystemClipboard().setContent(content);
            
            updateStatus("已复制单元格内容");
        } else {
            showInfo("提示", "请先选择要复制的单元格");
        }
    }
    
    /**
     * 粘贴单元格内容
     */
    @FXML
    public void handlePaste() {
        int selectedRow = tableView.getSelectionModel().getSelectedIndex();
        @SuppressWarnings("unchecked")
        TablePosition<ObservableList<CSVCell>, ?> focusedCell = 
            (TablePosition<ObservableList<CSVCell>, ?>) tableView.getFocusModel().getFocusedCell();
        
        if (selectedRow >= 0 && focusedCell != null && focusedCell.getColumn() > 0) {
            int selectedColumn = focusedCell.getColumn() - 1;
            
            // 优先使用系统剪贴板
            String pasteContent = clipboardContent;
            javafx.scene.input.Clipboard clipboard = javafx.scene.input.Clipboard.getSystemClipboard();
            if (clipboard.hasString()) {
                pasteContent = clipboard.getString();
            }
            
            saveHistory();
            dataModified = true;
            csvData.setCellValue(selectedRow, selectedColumn, pasteContent);
            tableView.refresh();
            updateStatus("已粘贴内容");
        } else {
            showInfo("提示", "请先选择要粘贴到的单元格");
        }
    }
    
    /**
     * 清除单元格内容
     */
    @FXML
    public void handleClearCell() {
        int selectedRow = tableView.getSelectionModel().getSelectedIndex();
        @SuppressWarnings("unchecked")
        TablePosition<ObservableList<CSVCell>, ?> focusedCell = 
            (TablePosition<ObservableList<CSVCell>, ?>) tableView.getFocusModel().getFocusedCell();
        
        if (selectedRow >= 0 && focusedCell != null && focusedCell.getColumn() > 0) {
            int selectedColumn = focusedCell.getColumn() - 1;
            saveHistory();
            dataModified = true;
            csvData.setCellValue(selectedRow, selectedColumn, "");
            tableView.refresh();
            updateStatus("已清除单元格内容");
        } else {
            showInfo("提示", "请先选择要清除的单元格");
        }
    }
    
    /**
     * 退出应用
     */
    @FXML
    public boolean handleExit() {
        // 如果数据被修改，询问是否保存
        if (dataModified && csvData.getRows() > 0) {
            Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
            alert.setTitle("退出确认");
            alert.setHeaderText("数据已修改");
            alert.setContentText("是否保存当前数据？");
            
            ButtonType buttonTypeSave = new ButtonType("保存", ButtonBar.ButtonData.YES);
            ButtonType buttonTypeDiscard = new ButtonType("不保存", ButtonBar.ButtonData.NO);
            ButtonType buttonTypeCancel = new ButtonType("取消", ButtonBar.ButtonData.CANCEL_CLOSE);
            
            alert.getButtonTypes().setAll(buttonTypeSave, buttonTypeDiscard, buttonTypeCancel);
            
            Optional<ButtonType> result = alert.showAndWait();
            if (result.isPresent()) {
                if (result.get() == buttonTypeSave) {
                    // 保存数据
                    if (currentFile != null) {
                        saveToFile(currentFile);
                    } else {
                        // 如果没有文件路径，执行另存为
                        FileChooser fileChooser = new FileChooser();
                        fileChooser.setTitle("保存CSV文件");
                        fileChooser.getExtensionFilters().add(
                            new FileChooser.ExtensionFilter("CSV文件", "*.csv")
                        );
                        
                        File file = fileChooser.showSaveDialog(getStage());
                        if (file != null) {
                            saveToFile(file);
                        } else {
                            return false; // 取消保存
                        }
                    }
                } else if (result.get() == buttonTypeCancel) {
                    return false; // 取消退出
                }
                // 不保存，直接退出
            } else {
                return false; // 取消退出
            }
        }
        
        // 关闭窗口
        Stage stage = getStage();
        if (stage != null) {
            stage.close();
        }
        return true;
    }
    
    /**
     * 显示关于对话框
     */
    @FXML
    public void handleAbout() {
        Alert alert = new Alert(Alert.AlertType.INFORMATION);
        alert.setTitle("关于");
        alert.setHeaderText("CSV表格编辑器");
        alert.setContentText("版本: 1.0\n一个简单的CSV表格编辑工具\n支持创建、编辑和保存CSV文件");
        alert.showAndWait();
    }
    
    /**
     * 刷新表格视图
     */
    private void refreshTable() {
        // 清空现有列
        tableView.getColumns().clear();
        
        int columnCount = csvData.getColumns();
        
        // 如果没有列，直接返回
        if (columnCount == 0) {
            tableView.setItems(csvData.getData());
            return;
        }
        
        // 添加行号列（如果启用）
        if (settings.isShowLineNumbers()) {
            TableColumn<ObservableList<CSVCell>, String> rowNumberColumn = new TableColumn<>("#");
            rowNumberColumn.setCellValueFactory(param -> {
                int rowIndex = tableView.getItems().indexOf(param.getValue());
                return new SimpleStringProperty(String.valueOf(rowIndex + 1));
            });
            rowNumberColumn.setEditable(false);
            rowNumberColumn.setSortable(false);
            rowNumberColumn.setResizable(false);
            rowNumberColumn.setPrefWidth(50);
            rowNumberColumn.setMinWidth(50);
            rowNumberColumn.setMaxWidth(50);
            rowNumberColumn.setStyle("-fx-alignment: CENTER;");
            tableView.getColumns().add(rowNumberColumn);
        }
        
        // 创建数据列
        for (int i = 0; i < columnCount; i++) {
            final int columnIndex = i;
            TableColumn<ObservableList<CSVCell>, String> column = new TableColumn<>("列 " + (i + 1));
            
            // 设置单元格值工厂
            column.setCellValueFactory(param -> {
                ObservableList<CSVCell> row = param.getValue();
                if (columnIndex < row.size()) {
                    return row.get(columnIndex).valueProperty();
                }
                return new SimpleStringProperty("");
            });
            
            // 设置自定义单元格编辑器，支持换行和高亮
            column.setCellFactory(col -> new MultiLineTableCell(columnIndex));
            column.setOnEditCommit(event -> {
                saveHistory();
                dataModified = true;
                int row = event.getTablePosition().getRow();
                csvData.setCellValue(row, columnIndex, event.getNewValue());
                
                // 单元格编辑后，重新检查该单元格的自动标记规则
                reapplyAutoMarkForCell(row, columnIndex);
                
                updatePositionLabel();
                // 刷新该单元格的显示
                tableView.refresh();
            });
            
            // 设置最小宽度，防止列太窄
            column.setMinWidth(60);
            
            // 允许用户调整列宽
            column.setResizable(true);
            
            tableView.getColumns().add(column);
        }
        
        // 设置表格数据
        tableView.setItems(csvData.getData());
        
        // 设置行工厂，用于优化无效行的显示样式和右键菜单
        tableView.setRowFactory(tv -> {
            TableRow<ObservableList<CSVCell>> row = new TableRow<>() {
                @Override
                protected void updateItem(ObservableList<CSVCell> item, boolean empty) {
                    super.updateItem(item, empty);
                    
                    if (empty || item == null) {
                        // 空行，清除所有样式和菜单
                        setStyle("-fx-background-color: rgba(0, 0, 0, 0);");
                        setContextMenu(null);
                        // 重要：清除header-row样式类，防止行重用时样式残留
                        getStyleClass().remove("header-row");
                    } else {
                        // 有效行，检查是否有行高亮
                        int rowIndex = getIndex();
                        updateRowStyle(this, rowIndex);
                        
                        // 添加右键菜单
                        ContextMenu contextMenu = createCellContextMenu(rowIndex);
                        setContextMenu(contextMenu);
                    }
                }
                
                @Override
                public void updateSelected(boolean selected) {
                    super.updateSelected(selected);
                    if (!isEmpty()) {
                        int rowIndex = getIndex();
                        updateRowStyle(this, rowIndex);
                    }
                }
            };
            
            // 添加选中监听
            row.setOnMouseClicked(event -> {
                if (!row.isEmpty()) {
                    updatePositionLabel();
                }
            });
            
            return row;
        });
        
        // 使用Platform.runLater确保在表格渲染后调整列宽
        javafx.application.Platform.runLater(() -> {
            adjustColumnWidths();
            // 应用背景图片
            applyBackgroundImage();
        });
    }
    
    /**
     * 更新行的样式（包括选中状态和高亮）
     */
    private void updateRowStyle(TableRow<ObservableList<CSVCell>> row, int rowIndex) {
        // 检查是否为标题行（首行且启用了首行为标题选项）
        boolean shouldBeHeader = settings.isFirstRowAsHeader() && rowIndex == 0;
        boolean hasHeaderClass = row.getStyleClass().contains("header-row");
        
        if (shouldBeHeader && !hasHeaderClass) {
            // 应该是表头但还没有样式类，添加
            row.getStyleClass().add("header-row");
        } else if (!shouldBeHeader && hasHeaderClass) {
            // 不应该是表头但有样式类，移除
            row.getStyleClass().remove("header-row");
        }
        
        if (row.isSelected()) {
            // 应用选中行颜色
            try {
                javafx.scene.paint.Color selectedColor = javafx.scene.paint.Color.web(settings.getSelectedRowColor());
                row.setStyle(String.format("-fx-background-color: rgba(%d, %d, %d, %.2f); -fx-text-fill: white;",
                    (int)(selectedColor.getRed() * 255),
                    (int)(selectedColor.getGreen() * 255),
                    (int)(selectedColor.getBlue() * 255),
                    selectedColor.getOpacity()));
            } catch (Exception e) {
                // 如果颜色格式错误，使用默认颜色
                row.setStyle("-fx-background-color: #3498db; -fx-text-fill: white;");
            }
        } else {
            // 未选中时，检查是否有行高亮
            javafx.scene.paint.Color rowColor = highlightManager.getRowHighlight(rowIndex) != null ?
                highlightManager.getRowHighlight(rowIndex).getColor() : null;
            
            if (rowColor != null) {
                row.setStyle(String.format("-fx-background-color: rgba(%d, %d, %d, %.2f);",
                    (int)(rowColor.getRed() * 255),
                    (int)(rowColor.getGreen() * 255),
                    (int)(rowColor.getBlue() * 255),
                    rowColor.getOpacity()));
            } else {
                row.setStyle("");
            }
        }
    }
    
    /**
     * 调整所有列的宽度，使其平均分配表格宽度
     */
    private void adjustColumnWidths() {
        if (tableView.getColumns().isEmpty()) {
            return;
        }
        
        // 获取表格的可用宽度
        double tableWidth = tableView.getWidth();
        
        // 如果表格宽度还没有计算出来，使用默认宽度
        if (tableWidth <= 0) {
            tableWidth = 900; // 使用FXML中定义的默认宽度
        }
        
        // 减去滚动条的宽度（大约15像素）和行号列的宽度（50像素）
        double usableWidth = tableWidth - 15 - 50;
        
        // 计算数据列的数量（排除行号列）
        int dataColumnCount = tableView.getColumns().size() - 1;
        
        if (dataColumnCount <= 0) {
            return;
        }
        
        // 计算每列应该占用的宽度
        double columnWidth = usableWidth / dataColumnCount;
        
        // 确保列宽不小于最小宽度
        columnWidth = Math.max(columnWidth, 60);
        
        // 为每一列设置相同的宽度（跳过行号列）
        for (int i = 1; i < tableView.getColumns().size(); i++) {
            tableView.getColumns().get(i).setPrefWidth(columnWidth);
        }
    }
    
    /**
     * 保存到文件
     */
    private void saveToFile(File file) {
        try {
            csvService.setLineEnding(settings.getLineEndingString());
            csvService.saveToFile(csvData, file);
            dataModified = false;
            updateStatus("已保存文件: " + file.getName() + " (" + 
                       csvData.getRows() + "行 x " + csvData.getColumns() + "列)");
            showInfo("保存成功", "文件已成功保存到: " + file.getAbsolutePath());
        } catch (IOException e) {
            showError("保存文件失败", 
                     "无法写入文件，请检查文件路径和权限。\n\n错误信息: " + e.getMessage());
        } catch (Exception e) {
            showError("保存文件失败", 
                     "保存CSV文件时发生错误。\n\n错误信息: " + e.getMessage());
        }
    }
    
    /**
     * 更新状态栏
     */
    private void updateStatus(String message) {
        if (statusLabel != null) {
            statusLabel.setText(message);
        }
    }
    
    /**
     * 显示错误对话框
     */
    private void showError(String title, String content) {
        Alert alert = new Alert(Alert.AlertType.ERROR);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(content);
        alert.showAndWait();
    }
    
    /**
     * 显示信息对话框
     */
    private void showInfo(String title, String content) {
        Alert alert = new Alert(Alert.AlertType.INFORMATION);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(content);
        alert.showAndWait();
    }
    
    /**
     * 显示确认对话框
     */
    private boolean showConfirmation(String title, String content) {
        Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(content);
        
        ButtonType buttonTypeYes = new ButtonType("是", ButtonBar.ButtonData.YES);
        ButtonType buttonTypeNo = new ButtonType("否", ButtonBar.ButtonData.NO);
        alert.getButtonTypes().setAll(buttonTypeYes, buttonTypeNo);
        
        Optional<ButtonType> result = alert.showAndWait();
        return result.isPresent() && result.get() == buttonTypeYes;
    }
    
    /**
     * 获取当前Stage
     */
    private Stage getStage() {
        if (tableView != null && tableView.getScene() != null) {
            return (Stage) tableView.getScene().getWindow();
        }
        return null;
    }
    
    /**
     * 确认是否放弃未保存的更改
     */
    private boolean confirmDiscardChanges() {
        Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
        alert.setTitle("确认");
        alert.setHeaderText("当前数据未保存");
        alert.setContentText("是否放弃当前的更改？");
        
        ButtonType buttonTypeYes = new ButtonType("是", ButtonBar.ButtonData.YES);
        ButtonType buttonTypeNo = new ButtonType("否", ButtonBar.ButtonData.NO);
        alert.getButtonTypes().setAll(buttonTypeYes, buttonTypeNo);
        
        Optional<ButtonType> result = alert.showAndWait();
        return result.isPresent() && result.get() == buttonTypeYes;
    }
    
    /**
     * 打开新窗口
     */
    private void openNewWindow() {
        try {
            javafx.fxml.FXMLLoader fxmlLoader = new javafx.fxml.FXMLLoader(
                getClass().getResource("/hbnu/project/ergoucsveditior/main-view.fxml")
            );
            javafx.scene.Parent root = fxmlLoader.load();
            
            Stage newStage = new Stage();
            newStage.setTitle("CSV表格编辑器 - 新建");
            newStage.setScene(new javafx.scene.Scene(root));
            newStage.show();
            
            updateStatus("已在新窗口打开");
        } catch (Exception e) {
            showError("打开新窗口失败", "无法创建新窗口: " + e.getMessage());
        }
    }
    
    /**
     * 更新位置标签，显示当前选中单元格的行列坐标
     */
    private void updatePositionLabel() {
        if (positionLabel == null) {
            return;
        }
        
        int selectedRow = tableView.getSelectionModel().getSelectedIndex();
        @SuppressWarnings("unchecked")
        TablePosition<ObservableList<CSVCell>, ?> focusedCell = 
            (TablePosition<ObservableList<CSVCell>, ?>) tableView.getFocusModel().getFocusedCell();
        
        if (selectedRow >= 0 && focusedCell != null && focusedCell.getColumn() > 0) {
            // 减1是因为第一列是行号列
            int selectedColumn = focusedCell.getColumn() - 1;
            positionLabel.setText("行: " + (selectedRow + 1) + ", 列: " + (selectedColumn + 1));
        } else if (selectedRow >= 0) {
            positionLabel.setText("行: " + (selectedRow + 1));
        } else {
            positionLabel.setText("未选中");
        }
    }
    
    /**
     * 保存当前状态到历史记录
     */
    private void saveHistory() {
        historyManager.saveState(csvData);
        updateUndoButton();
    }
    
    /**
     * 更新撤销按钮状态
     */
    private void updateUndoButton() {
        if (undoButton != null) {
            undoButton.setDisable(!historyManager.canUndo());
        }
    }
    
    /**
     * 显示设置对话框
     */
    private void showSettingsDialog() {
        Dialog<Boolean> dialog = new Dialog<>();
        dialog.setTitle("设置");
        dialog.setHeaderText("应用设置");
        
        // 设置按钮
        ButtonType saveButtonType = new ButtonType("保存", ButtonBar.ButtonData.OK_DONE);
        dialog.getDialogPane().getButtonTypes().addAll(saveButtonType, ButtonType.CANCEL);
        
        // 创建设置表单
        javafx.scene.layout.GridPane grid = new javafx.scene.layout.GridPane();
        grid.setHgap(10);
        grid.setVgap(10);
        grid.setPadding(new javafx.geometry.Insets(20, 150, 10, 10));
        
        // 字符编码
        ComboBox<String> encodingCombo = new ComboBox<>();
        encodingCombo.getItems().addAll("UTF-8", "GBK", "GB2312", "ISO-8859-1");
        encodingCombo.setValue(settings.getDefaultEncoding());
        
        // 历史记录大小
        TextField historyField = new TextField(String.valueOf(settings.getMaxHistorySize()));
        
        // 换行符类型
        ComboBox<String> lineEndingCombo = new ComboBox<>();
        lineEndingCombo.getItems().addAll("LF", "CRLF");
        lineEndingCombo.setValue(settings.getLineEnding());
        
        // 自动保存
        CheckBox autoSaveCheck = new CheckBox();
        autoSaveCheck.setSelected(settings.isAutoSaveEnabled());
        
        TextField autoSaveIntervalField = new TextField(String.valueOf(settings.getAutoSaveInterval()));
        autoSaveIntervalField.setDisable(!settings.isAutoSaveEnabled());
        
        autoSaveCheck.selectedProperty().addListener((obs, oldVal, newVal) -> {
            autoSaveIntervalField.setDisable(!newVal);
        });
        
        int row = 0;
        grid.add(new Label("默认编码:"), 0, row);
        grid.add(encodingCombo, 1, row++);
        
        grid.add(new Label("撤销历史数:"), 0, row);
        grid.add(historyField, 1, row++);
        
        grid.add(new Label("换行符:"), 0, row);
        grid.add(lineEndingCombo, 1, row++);
        
        grid.add(new Label("自动保存:"), 0, row);
        grid.add(autoSaveCheck, 1, row++);
        
        grid.add(new Label("保存间隔(分钟):"), 0, row);
        grid.add(autoSaveIntervalField, 1, row++);
        
        // 高亮冲突策略
        ComboBox<String> conflictStrategyCombo = new ComboBox<>();
        conflictStrategyCombo.getItems().addAll("覆盖策略", "随机策略");
        conflictStrategyCombo.setValue(settings.getHighlightConflictStrategy());
        
        grid.add(new Label("高亮冲突策略:"), 0, row);
        grid.add(conflictStrategyCombo, 1, row++);
        
        // 添加策略说明
        Label strategyHint = new Label("覆盖策略：后标记覆盖先标记\n随机策略：随机选择行色或列色");
        strategyHint.setStyle("-fx-font-size: 10px; -fx-text-fill: #666666;");
        strategyHint.setWrapText(true);
        grid.add(strategyHint, 1, row++);
        
        // 添加分隔符
        grid.add(new Separator(), 0, row++, 2, 1);
        
        // 自动检测分隔符
        CheckBox autoDetectDelimiterCheck = new CheckBox();
        autoDetectDelimiterCheck.setSelected(settings.isAutoDetectDelimiter());
        
        grid.add(new Label("自动检测分隔符:"), 0, row);
        grid.add(autoDetectDelimiterCheck, 1, row++);
        
        // 转义字符模式
        ComboBox<String> escapeModeCombo = new ComboBox<>();
        escapeModeCombo.getItems().addAll("重复引号", "反斜杠转义");
        escapeModeCombo.setValue(settings.getEscapeMode());
        
        grid.add(new Label("转义字符模式:"), 0, row);
        grid.add(escapeModeCombo, 1, row++);
        
        // 首行是否为标题
        CheckBox firstRowAsHeaderCheck = new CheckBox();
        firstRowAsHeaderCheck.setSelected(settings.isFirstRowAsHeader());
        
        grid.add(new Label("首行为标题:"), 0, row);
        grid.add(firstRowAsHeaderCheck, 1, row++);
        
        // 显示行号
        CheckBox showLineNumbersCheck = new CheckBox();
        showLineNumbersCheck.setSelected(settings.isShowLineNumbers());
        
        grid.add(new Label("显示行号:"), 0, row);
        grid.add(showLineNumbersCheck, 1, row++);
        
        dialog.getDialogPane().setContent(grid);
        
        // 转换结果
        dialog.setResultConverter(dialogButton -> {
            if (dialogButton == saveButtonType) {
                try {
                    settings.setDefaultEncoding(encodingCombo.getValue());
                    settings.setMaxHistorySize(Integer.parseInt(historyField.getText()));
                    settings.setLineEnding(lineEndingCombo.getValue());
                    settings.setAutoSaveEnabled(autoSaveCheck.isSelected());
                    settings.setAutoSaveInterval(Integer.parseInt(autoSaveIntervalField.getText()));
                    settings.setHighlightConflictStrategy(conflictStrategyCombo.getValue());
                    
                    // 保存新设置
                    settings.setAutoDetectDelimiter(autoDetectDelimiterCheck.isSelected());
                    settings.setEscapeMode(escapeModeCombo.getValue());
                    settings.setFirstRowAsHeader(firstRowAsHeaderCheck.isSelected());
                    settings.setShowLineNumbers(showLineNumbersCheck.isSelected());
                    
                    settings.save();
                    
                    // 重新初始化历史管理器
                    historyManager = new hbnu.project.ergoucsveditior.model.HistoryManager(
                        settings.getMaxHistorySize());
                    
                    // 更新高亮冲突策略
                    if ("随机策略".equals(conflictStrategyCombo.getValue())) {
                        highlightManager.setConflictStrategy(
                            hbnu.project.ergoucsveditior.model.HighlightManager.ConflictStrategy.随机策略);
                    } else {
                        highlightManager.setConflictStrategy(
                            hbnu.project.ergoucsveditior.model.HighlightManager.ConflictStrategy.覆盖策略);
                    }
                    
                    // 刷新表格以应用新设置
                    refreshTable();
                    
                    return true;
                } catch (NumberFormatException e1) {
                    javafx.application.Platform.runLater(() -> 
                        showError("输入错误", "请输入有效的数字")
                    );
                }
            }
            return false;
        });
        
        Optional<Boolean> result = dialog.showAndWait();
        if (result.isPresent() && result.get()) {
            updateStatus("设置已保存");
        }
    }
    
    /**
     * 显示工具栏配置对话框
     */
    private void showToolbarConfigDialog() {
        Dialog<Boolean> dialog = new Dialog<>();
        dialog.setTitle("工具栏设置");
        dialog.setHeaderText("自定义工具栏按钮和样式");
        
        // 设置按钮
        ButtonType saveButtonType = new ButtonType("保存", ButtonBar.ButtonData.OK_DONE);
        dialog.getDialogPane().getButtonTypes().addAll(saveButtonType, ButtonType.CANCEL);
        
        // 创建主布局
        javafx.scene.layout.VBox mainBox = new javafx.scene.layout.VBox(15);
        mainBox.setPadding(new javafx.geometry.Insets(20));
        
        // 样式设置区域
        TitledPane stylePane = new TitledPane();
        stylePane.setText("按钮样式");
        javafx.scene.layout.GridPane styleGrid = new javafx.scene.layout.GridPane();
        styleGrid.setHgap(10);
        styleGrid.setVgap(10);
        styleGrid.setPadding(new javafx.geometry.Insets(10));
        
        // 按钮颜色
        javafx.scene.control.ColorPicker buttonColorPicker = new javafx.scene.control.ColorPicker(
            javafx.scene.paint.Color.web(toolbarConfig.getButtonColor()));
        javafx.scene.control.ColorPicker hoverColorPicker = new javafx.scene.control.ColorPicker(
            javafx.scene.paint.Color.web(toolbarConfig.getButtonHoverColor()));
        javafx.scene.control.ColorPicker textColorPicker = new javafx.scene.control.ColorPicker(
            javafx.scene.paint.Color.web(toolbarConfig.getButtonTextColor()));
        
        styleGrid.add(new Label("按钮颜色:"), 0, 0);
        styleGrid.add(buttonColorPicker, 1, 0);
        styleGrid.add(new Label("悬停颜色:"), 0, 1);
        styleGrid.add(hoverColorPicker, 1, 1);
        styleGrid.add(new Label("文字颜色:"), 0, 2);
        styleGrid.add(textColorPicker, 1, 2);
        
        stylePane.setContent(styleGrid);
        
        // 按钮选择区域
        TitledPane buttonsPane = new TitledPane();
        buttonsPane.setText("显示的按钮");
        javafx.scene.layout.VBox buttonsBox = new javafx.scene.layout.VBox(10);
        buttonsBox.setPadding(new javafx.geometry.Insets(10));
        
        // 可用按钮列表
        javafx.scene.control.ListView<String> availableList = new javafx.scene.control.ListView<>();
        availableList.setPrefHeight(200);
        
        // 已选按钮列表
        javafx.scene.control.ListView<String> selectedList = new javafx.scene.control.ListView<>();
        selectedList.setPrefHeight(200);
        
        // 填充列表
        java.util.List<String> currentButtons = toolbarConfig.getVisibleButtons();
        for (String btn : currentButtons) {
            String displayName = "separator".equals(btn) ? "--- 分隔符 ---" : 
                hbnu.project.ergoucsveditior.model.ToolbarConfig.getButtonDisplayName(btn);
            selectedList.getItems().add(btn + ":" + displayName);
        }
        
        for (String btn : hbnu.project.ergoucsveditior.model.ToolbarConfig.getAllButtons()) {
            if (!currentButtons.contains(btn)) {
                String displayName = hbnu.project.ergoucsveditior.model.ToolbarConfig.getButtonDisplayName(btn);
                availableList.getItems().add(btn + ":" + displayName);
            }
        }
        
        // 添加分隔符选项
        if (!currentButtons.contains("separator")) {
            availableList.getItems().add("separator:--- 分隔符 ---");
        }
        
        // 操作按钮
        javafx.scene.layout.VBox controlBox = new javafx.scene.layout.VBox(5);
        Button addButton = new Button("添加 →");
        Button removeButton = new Button("← 移除");
        Button upButton = new Button("上移 ↑");
        Button downButton = new Button("下移 ↓");
        
        addButton.setPrefWidth(80);
        removeButton.setPrefWidth(80);
        upButton.setPrefWidth(80);
        downButton.setPrefWidth(80);
        
        controlBox.getChildren().addAll(addButton, removeButton, upButton, downButton);
        controlBox.setAlignment(javafx.geometry.Pos.CENTER);
        
        // 添加按钮事件
        addButton.setOnAction(e -> {
            String selected = availableList.getSelectionModel().getSelectedItem();
            if (selected != null) {
                selectedList.getItems().add(selected);
                availableList.getItems().remove(selected);
            }
        });
        
        removeButton.setOnAction(e -> {
            String selected = selectedList.getSelectionModel().getSelectedItem();
            if (selected != null) {
                availableList.getItems().add(selected);
                selectedList.getItems().remove(selected);
            }
        });
        
        upButton.setOnAction(e -> {
            int index = selectedList.getSelectionModel().getSelectedIndex();
            if (index > 0) {
                String item = selectedList.getItems().remove(index);
                selectedList.getItems().add(index - 1, item);
                selectedList.getSelectionModel().select(index - 1);
            }
        });
        
        downButton.setOnAction(e -> {
            int index = selectedList.getSelectionModel().getSelectedIndex();
            if (index >= 0 && index < selectedList.getItems().size() - 1) {
                String item = selectedList.getItems().remove(index);
                selectedList.getItems().add(index + 1, item);
                selectedList.getSelectionModel().select(index + 1);
            }
        });
        
        // 布局按钮列表
        javafx.scene.layout.HBox listsBox = new javafx.scene.layout.HBox(10);
        javafx.scene.layout.VBox availableBox = new javafx.scene.layout.VBox(5);
        availableBox.getChildren().addAll(new Label("可用按钮:"), availableList);
        javafx.scene.layout.VBox selectedBox = new javafx.scene.layout.VBox(5);
        selectedBox.getChildren().addAll(new Label("工具栏按钮:"), selectedList);
        
        listsBox.getChildren().addAll(availableBox, controlBox, selectedBox);
        buttonsBox.getChildren().add(listsBox);
        buttonsPane.setContent(buttonsBox);
        
        mainBox.getChildren().addAll(stylePane, buttonsPane);
        dialog.getDialogPane().setContent(mainBox);
        
        // 转换结果
        dialog.setResultConverter(dialogButton -> {
            if (dialogButton == saveButtonType) {
                // 保存样式配置
                toolbarConfig.setButtonColor(toHexString(buttonColorPicker.getValue()));
                toolbarConfig.setButtonHoverColor(toHexString(hoverColorPicker.getValue()));
                toolbarConfig.setButtonTextColor(toHexString(textColorPicker.getValue()));
                
                // 保存按钮列表
                java.util.List<String> buttons = new java.util.ArrayList<>();
                for (String item : selectedList.getItems()) {
                    String buttonId = item.split(":")[0];
                    buttons.add(buttonId);
                }
                toolbarConfig.setVisibleButtons(buttons);
                toolbarConfig.save();
                
                // 重新初始化工具栏
                initializeToolbar();
                updateUndoButton();
                
                return true;
            }
            return false;
        });
        
        Optional<Boolean> result = dialog.showAndWait();
        if (result.isPresent() && result.get()) {
            updateStatus("工具栏设置已保存");
        }
    }
    
    /**
     * 将颜色转换为16进制字符串
     */
    private String toHexString(javafx.scene.paint.Color color) {
        return String.format("#%02X%02X%02X",
            (int)(color.getRed() * 255),
            (int)(color.getGreen() * 255),
            (int)(color.getBlue() * 255));
    }
    
    /**
     * 显示高亮设置对话框
     */
    private void showHighlightSettingsDialog() {
        Dialog<Boolean> dialog = new Dialog<>();
        dialog.setTitle("高亮设置");
        dialog.setHeaderText("自定义高亮颜色偏好");
        
        // 设置按钮
        ButtonType saveButtonType = new ButtonType("保存", ButtonBar.ButtonData.OK_DONE);
        dialog.getDialogPane().getButtonTypes().addAll(saveButtonType, ButtonType.CANCEL);
        
        // 创建设置表单
        javafx.scene.layout.GridPane grid = new javafx.scene.layout.GridPane();
        grid.setHgap(10);
        grid.setVgap(10);
        grid.setPadding(new javafx.geometry.Insets(20, 150, 10, 10));
        
        int row = 0;
        
        // 默认单元格高亮背景色
        Label cellBgLabel = new Label("单元格背景色:");
        javafx.scene.control.ColorPicker cellBgPicker = new javafx.scene.control.ColorPicker(
            javafx.scene.paint.Color.web(settings.getDefaultCellHighlightColor())
        );
        grid.add(cellBgLabel, 0, row);
        grid.add(cellBgPicker, 1, row++);
        
        // 默认单元格文本颜色
        Label cellTextLabel = new Label("单元格文本色:");
        javafx.scene.control.ColorPicker cellTextPicker = new javafx.scene.control.ColorPicker(javafx.scene.paint.Color.BLACK);
        CheckBox cellTextEnabled = new CheckBox("启用");
        if (settings.getDefaultCellTextColor() != null && !settings.getDefaultCellTextColor().isEmpty()) {
            try {
                cellTextPicker.setValue(javafx.scene.paint.Color.web(settings.getDefaultCellTextColor()));
                cellTextEnabled.setSelected(true);
            } catch (Exception e) {
                cellTextEnabled.setSelected(false);
            }
        } else {
            cellTextEnabled.setSelected(false);
        }
        cellTextPicker.setDisable(!cellTextEnabled.isSelected());
        cellTextEnabled.selectedProperty().addListener((obs, oldVal, newVal) -> {
            cellTextPicker.setDisable(!newVal);
        });
        javafx.scene.layout.HBox cellTextBox = new javafx.scene.layout.HBox(5, cellTextPicker, cellTextEnabled);
        grid.add(cellTextLabel, 0, row);
        grid.add(cellTextBox, 1, row++);
        
        // 默认行高亮背景色
        Label rowBgLabel = new Label("行背景色:");
        javafx.scene.control.ColorPicker rowBgPicker = new javafx.scene.control.ColorPicker(
            javafx.scene.paint.Color.web(settings.getDefaultRowHighlightColor())
        );
        grid.add(rowBgLabel, 0, row);
        grid.add(rowBgPicker, 1, row++);
        
        // 默认列高亮背景色
        Label colBgLabel = new Label("列背景色:");
        javafx.scene.control.ColorPicker colBgPicker = new javafx.scene.control.ColorPicker(
            javafx.scene.paint.Color.web(settings.getDefaultColumnHighlightColor())
        );
        grid.add(colBgLabel, 0, row);
        grid.add(colBgPicker, 1, row++);
        
        // 搜索结果高亮色
        Label searchLabel = new Label("搜索高亮色:");
        javafx.scene.control.ColorPicker searchPicker = new javafx.scene.control.ColorPicker(
            javafx.scene.paint.Color.web(settings.getSearchHighlightColor())
        );
        grid.add(searchLabel, 0, row);
        grid.add(searchPicker, 1, row++);
        
        // 选中行颜色
        Label selectedLabel = new Label("选中行颜色:");
        javafx.scene.control.ColorPicker selectedPicker = new javafx.scene.control.ColorPicker(
            javafx.scene.paint.Color.web(settings.getSelectedRowColor())
        );
        grid.add(selectedLabel, 0, row);
        grid.add(selectedPicker, 1, row++);
        
        // 添加说明
        Label hint = new Label("提示：这些设置将作为默认颜色应用到新的高亮标记");
        hint.setStyle("-fx-font-size: 10px; -fx-text-fill: #666666;");
        hint.setWrapText(true);
        grid.add(hint, 0, row, 2, 1);
        
        dialog.getDialogPane().setContent(grid);
        
        // 转换结果
        dialog.setResultConverter(dialogButton -> {
            if (dialogButton == saveButtonType) {
                // 保存设置
                settings.setDefaultCellHighlightColor(toHexString(cellBgPicker.getValue()));
                if (cellTextEnabled.isSelected()) {
                    settings.setDefaultCellTextColor(toHexString(cellTextPicker.getValue()));
                } else {
                    settings.setDefaultCellTextColor("");
                }
                settings.setDefaultRowHighlightColor(toHexString(rowBgPicker.getValue()));
                settings.setDefaultColumnHighlightColor(toHexString(colBgPicker.getValue()));
                settings.setSearchHighlightColor(toHexString(searchPicker.getValue()));
                settings.setSelectedRowColor(toHexString(selectedPicker.getValue()));
                settings.save();
                
                // 更新HighlightManager的默认颜色
                highlightManager.setDefaultCellHighlightColor(cellBgPicker.getValue());
                highlightManager.setDefaultRowHighlightColor(rowBgPicker.getValue());
                highlightManager.setDefaultColumnHighlightColor(colBgPicker.getValue());
                highlightManager.setDefaultSearchHighlightColor(searchPicker.getValue());
                
                // 刷新表格以应用新的颜色设置
                tableView.refresh();
                
                return true;
            }
            return false;
        });
        
        Optional<Boolean> result = dialog.showAndWait();
        if (result.isPresent() && result.get()) {
            updateStatus("高亮设置已保存");
        }
    }
    
    /**
     * 显示主题设置对话框
     */
    private void showThemeSettingsDialog() {
        Dialog<Boolean> dialog = new Dialog<>();
        dialog.setTitle("主题设置");
        dialog.setHeaderText("自定义应用主题和背景");
        
        // 设置按钮
        ButtonType saveButtonType = new ButtonType("保存", ButtonBar.ButtonData.OK_DONE);
        dialog.getDialogPane().getButtonTypes().addAll(saveButtonType, ButtonType.CANCEL);
        
        // 创建设置表单
        javafx.scene.layout.GridPane grid = new javafx.scene.layout.GridPane();
        grid.setHgap(10);
        grid.setVgap(10);
        grid.setPadding(new javafx.geometry.Insets(20, 150, 10, 10));
        
        int row = 0;
        
        // 主题选择
        ComboBox<String> themeCombo = new ComboBox<>();
        themeCombo.getItems().addAll("浅色", "深色");
        themeCombo.setValue(settings.getTheme());
        
        grid.add(new Label("主题:"), 0, row);
        grid.add(themeCombo, 1, row++);
        
        // 添加分隔符
        grid.add(new Separator(), 0, row++, 2, 1);
        
        // 背景图片路径
        TextField backgroundImagePathField = new TextField(settings.getBackgroundImagePath());
        backgroundImagePathField.setPrefWidth(300);
        
        Button browseImageButton = new Button("选择图片...");
        browseImageButton.setOnAction(e -> {
            FileChooser fileChooser = new FileChooser();
            fileChooser.setTitle("选择背景图片");
            fileChooser.getExtensionFilters().addAll(
                new FileChooser.ExtensionFilter("图片文件", "*.png", "*.jpg", "*.jpeg", "*.gif", "*.bmp"),
                new FileChooser.ExtensionFilter("所有文件", "*.*")
            );
            File file = fileChooser.showOpenDialog(getStage());
            if (file != null) {
                backgroundImagePathField.setText(file.getAbsolutePath());
            }
        });
        
        javafx.scene.layout.HBox imageBox = new javafx.scene.layout.HBox(5);
        imageBox.getChildren().addAll(backgroundImagePathField, browseImageButton);
        
        grid.add(new Label("背景图片:"), 0, row);
        grid.add(imageBox, 1, row++);
        
        // 背景图片透明度
        javafx.scene.control.Slider opacitySlider = new javafx.scene.control.Slider(0.0, 1.0, settings.getBackgroundImageOpacity());
        opacitySlider.setMajorTickUnit(0.1);
        opacitySlider.setShowTickLabels(true);
        opacitySlider.setShowTickMarks(true);
        opacitySlider.setBlockIncrement(0.1);
        
        Label opacityLabel = new Label(String.format("%.0f%%", settings.getBackgroundImageOpacity() * 100));
        opacitySlider.valueProperty().addListener((obs, oldVal, newVal) -> {
            opacityLabel.setText(String.format("%.0f%%", newVal.doubleValue() * 100));
            // 实时预览透明度变化
            settings.setBackgroundImageOpacity(newVal.doubleValue());
            applyBackgroundImage();
        });
        
        javafx.scene.layout.HBox opacityBox = new javafx.scene.layout.HBox(10);
        opacityBox.getChildren().addAll(opacitySlider, opacityLabel);
        
        grid.add(new Label("图片透明度:"), 0, row);
        grid.add(opacityBox, 1, row++);
        
        // 背景图片适应模式
        ComboBox<String> fitModeCombo = new ComboBox<>();
        fitModeCombo.getItems().addAll("保持比例", "拉伸填充", "原始大小", "适应宽度", "适应高度");
        fitModeCombo.setValue(settings.getBackgroundImageFitMode());
        fitModeCombo.setOnAction(e -> {
            // 实时预览适应模式变化
            settings.setBackgroundImageFitMode(fitModeCombo.getValue());
            applyBackgroundImage();
        });
        
        grid.add(new Label("图片适应模式:"), 0, row);
        grid.add(fitModeCombo, 1, row++);
        
        // 添加分隔符
        grid.add(new Separator(), 0, row++, 2, 1);
        
        // 表格边框颜色
        javafx.scene.control.ColorPicker borderColorPicker = new javafx.scene.control.ColorPicker(
            javafx.scene.paint.Color.web(settings.getTableBorderColor()));
        
        grid.add(new Label("表格边框颜色:"), 0, row);
        grid.add(borderColorPicker, 1, row++);
        
        // 网格线颜色
        javafx.scene.control.ColorPicker gridColorPicker = new javafx.scene.control.ColorPicker(
            javafx.scene.paint.Color.web(settings.getTableGridColor()));
        
        grid.add(new Label("网格线颜色:"), 0, row);
        grid.add(gridColorPicker, 1, row++);
        
        // 添加说明
        Label hint = new Label("提示：调整透明度滑块可实时预览效果\n边框和网格线颜色将在重新加载表格后生效");
        hint.setStyle("-fx-font-size: 10px; -fx-text-fill: #666666;");
        hint.setWrapText(true);
        grid.add(hint, 0, row, 2, 1);
        
        dialog.getDialogPane().setContent(grid);
        
        // 转换结果
        dialog.setResultConverter(dialogButton -> {
            if (dialogButton == saveButtonType) {
                // 保存主题设置
                settings.setTheme(themeCombo.getValue());
                
                // 保存背景图片设置
                settings.setBackgroundImagePath(backgroundImagePathField.getText());
                settings.setBackgroundImageOpacity(opacitySlider.getValue());
                settings.setBackgroundImageFitMode(fitModeCombo.getValue());
                
                // 保存表格样式设置
                settings.setTableBorderColor(toHexString(borderColorPicker.getValue()));
                settings.setTableGridColor(toHexString(gridColorPicker.getValue()));
                
                settings.save();
                
                // 应用主题和表格样式
                applyTheme();
                applyTableStyles();
                
                return true;
            } else {
                // 取消时恢复原来的透明度设置
                settings.load();
                applyBackgroundImage();
            }
            return false;
        });
        
        Optional<Boolean> result = dialog.showAndWait();
        if (result.isPresent() && result.get()) {
            updateStatus("主题设置已保存");
        }
    }
    
    /**
     * 显示自动标记配置对话框
     */
    private void showAutoMarkDialog() {
        Dialog<Boolean> dialog = new Dialog<>();
        dialog.setTitle("自动标记");
        dialog.setHeaderText("配置自动标记规则");
        
        // 设置按钮
        ButtonType applyButtonType = new ButtonType("应用规则", ButtonBar.ButtonData.OK_DONE);
        ButtonType clearAllButtonType = new ButtonType("清除所有标记", ButtonBar.ButtonData.OTHER);
        ButtonType settingsButtonType = new ButtonType("颜色设置", ButtonBar.ButtonData.OTHER);
        dialog.getDialogPane().getButtonTypes().addAll(applyButtonType, clearAllButtonType, settingsButtonType, ButtonType.CANCEL);
        
        // 创建主面板
        javafx.scene.layout.BorderPane mainPane = new javafx.scene.layout.BorderPane();
        mainPane.setPrefSize(700, 500);
        
        // ========== 左侧规则列表 ==========
        javafx.scene.layout.VBox leftPane = new javafx.scene.layout.VBox(10);
        leftPane.setPadding(new javafx.geometry.Insets(10));
        leftPane.setPrefWidth(250);
        
        Label rulesLabel = new Label("已添加的规则：");
        rulesLabel.setStyle("-fx-font-weight: bold;");
        
        javafx.scene.control.ListView<hbnu.project.ergoucsveditior.model.AutoMarkRule> rulesList = 
            new javafx.scene.control.ListView<>();
        rulesList.getItems().addAll(autoMarkManager.getRules());
        rulesList.setPrefHeight(300);
        
        javafx.scene.layout.HBox ruleButtonsBox = new javafx.scene.layout.HBox(5);
        Button removeRuleBtn = new Button("删除规则");
        Button clearRulesBtn = new Button("清除全部");
        Button loadTemplateBtn = new Button("加载模板");
        Button saveTemplateBtn = new Button("保存为模板");
        
        removeRuleBtn.setOnAction(e -> {
            hbnu.project.ergoucsveditior.model.AutoMarkRule selected = rulesList.getSelectionModel().getSelectedItem();
            if (selected != null) {
                rulesList.getItems().remove(selected);
            }
        });
        
        clearRulesBtn.setOnAction(e -> {
            if (showConfirmation("确认", "是否清除所有规则？")) {
                rulesList.getItems().clear();
            }
        });
        
        loadTemplateBtn.setOnAction(e -> {
            // 显示模板选择对话框
            java.util.List<hbnu.project.ergoucsveditior.model.AutoMarkRule> templates = 
                autoMarkSettings.getRuleTemplates();
            if (templates.isEmpty()) {
                showInfo("提示", "没有保存的规则模板");
                return;
            }
            
            javafx.scene.control.ChoiceDialog<hbnu.project.ergoucsveditior.model.AutoMarkRule> templateDialog = 
                new javafx.scene.control.ChoiceDialog<>(templates.get(0), templates);
            templateDialog.setTitle("选择模板");
            templateDialog.setHeaderText("选择要加载的规则模板");
            templateDialog.setContentText("模板:");
            
            Optional<hbnu.project.ergoucsveditior.model.AutoMarkRule> templateResult = templateDialog.showAndWait();
            if (templateResult.isPresent()) {
                rulesList.getItems().add(templateResult.get());
            }
        });
        
        saveTemplateBtn.setOnAction(e -> {
            hbnu.project.ergoucsveditior.model.AutoMarkRule selected = rulesList.getSelectionModel().getSelectedItem();
            if (selected != null) {
                autoMarkSettings.addRuleTemplate(selected);
                showInfo("成功", "规则已保存为模板");
            } else {
                showInfo("提示", "请先选择一个规则");
            }
        });
        
        ruleButtonsBox.getChildren().addAll(removeRuleBtn, clearRulesBtn);
        
        javafx.scene.layout.HBox templateButtonsBox = new javafx.scene.layout.HBox(5);
        templateButtonsBox.getChildren().addAll(loadTemplateBtn, saveTemplateBtn);
        
        leftPane.getChildren().addAll(rulesLabel, rulesList, ruleButtonsBox, templateButtonsBox);
        
        // ========== 右侧规则配置 ==========
        javafx.scene.layout.VBox rightPane = new javafx.scene.layout.VBox(10);
        rightPane.setPadding(new javafx.geometry.Insets(10));
        
        Label configLabel = new Label("添加新规则：");
        configLabel.setStyle("-fx-font-weight: bold;");
        
        javafx.scene.layout.GridPane configGrid = new javafx.scene.layout.GridPane();
        configGrid.setHgap(10);
        configGrid.setVgap(10);
        
        int row = 0;
        
        // 规则名称
        TextField ruleNameField = new TextField();
        ruleNameField.setPromptText("规则名称");
        configGrid.add(new Label("规则名称:"), 0, row);
        configGrid.add(ruleNameField, 1, row++);
        
        // 规则类型
        ComboBox<String> ruleTypeCombo = new ComboBox<>();
        ruleTypeCombo.getItems().addAll(
            "数字 - 大于", "数字 - 小于", "数字 - 等于", "数字 - 是否为质数",
            "字符串 - 包含", "字符串 - 正则表达式",
            "格式校验 - 邮箱", "格式校验 - 手机号", "格式校验 - URL", "格式校验 - 身份证号",
            "空值 - 完全空值", "空值 - 仅含空白", "空值 - 长度为0"
        );
        ruleTypeCombo.setValue("数字 - 大于");
        configGrid.add(new Label("规则类型:"), 0, row);
        configGrid.add(ruleTypeCombo, 1, row++);
        
        // 规则参数
        TextField parameterField = new TextField();
        parameterField.setPromptText("参数值（如数字、字符串或正则表达式）");
        Label paramLabel = new Label("参数值:");
        configGrid.add(paramLabel, 0, row);
        configGrid.add(parameterField, 1, row++);
        
        // 根据规则类型动态调整参数输入提示
        ruleTypeCombo.setOnAction(e -> {
            String type = ruleTypeCombo.getValue();
            if (type.contains("质数") || type.contains("格式校验") || type.contains("空值")) {
                parameterField.setDisable(true);
                parameterField.setText("");
                paramLabel.setDisable(true);
            } else {
                parameterField.setDisable(false);
                paramLabel.setDisable(false);
                if (type.contains("数字")) {
                    parameterField.setPromptText("输入数字");
                } else if (type.contains("正则表达式")) {
                    parameterField.setPromptText("输入正则表达式，如 \\d{3}");
                } else {
                    parameterField.setPromptText("输入要包含的字符串");
                }
            }
        });
        
        // 标记颜色
        javafx.scene.control.ColorPicker colorPicker = new javafx.scene.control.ColorPicker(
            javafx.scene.paint.Color.web(autoMarkSettings.getNumberMarkColor()));
        configGrid.add(new Label("标记颜色:"), 0, row);
        configGrid.add(colorPicker, 1, row++);
        
        // 应用范围
        ComboBox<String> scopeCombo = new ComboBox<>();
        scopeCombo.getItems().addAll("所有列", "选中列", "指定列");
        scopeCombo.setValue("所有列");
        configGrid.add(new Label("应用范围:"), 0, row);
        configGrid.add(scopeCombo, 1, row++);
        
        // 指定列输入
        TextField columnsField = new TextField();
        columnsField.setPromptText("例如: 0,2,5 (列索引从0开始)");
        columnsField.setDisable(true);
        Label columnsLabel = new Label("指定列:");
        columnsLabel.setDisable(true);
        configGrid.add(columnsLabel, 0, row);
        configGrid.add(columnsField, 1, row++);
        
        scopeCombo.setOnAction(e -> {
            boolean isSpecified = "指定列".equals(scopeCombo.getValue());
            columnsField.setDisable(!isSpecified);
            columnsLabel.setDisable(!isSpecified);
        });
        
        // 添加规则按钮
        Button addRuleBtn = new Button("添加规则");
        addRuleBtn.setStyle("-fx-background-color: #4CAF50; -fx-text-fill: white;");
        addRuleBtn.setOnAction(e -> {
            try {
                // 验证输入
                if (ruleNameField.getText().isEmpty()) {
                    showError("错误", "请输入规则名称");
                    return;
                }
                
                String typeStr = ruleTypeCombo.getValue();
                boolean needsParam = !typeStr.contains("质数") && 
                                    !typeStr.contains("格式校验") && 
                                    !typeStr.contains("空值");
                
                if (needsParam && parameterField.getText().isEmpty()) {
                    showError("错误", "请输入参数值");
                    return;
                }
                
                // 创建规则
                hbnu.project.ergoucsveditior.model.AutoMarkRule rule = 
                    new hbnu.project.ergoucsveditior.model.AutoMarkRule();
                rule.setName(ruleNameField.getText());
                rule.setType(parseRuleType(typeStr));
                rule.setParameter(parameterField.getText());
                rule.setColor(toHexString(colorPicker.getValue()));
                rule.setScope(parseScopeType(scopeCombo.getValue()));
                
                // 处理指定列
                if ("指定列".equals(scopeCombo.getValue())) {
                    String[] colStrs = columnsField.getText().split(",");
                    int[] cols = new int[colStrs.length];
                    for (int i = 0; i < colStrs.length; i++) {
                        cols[i] = Integer.parseInt(colStrs[i].trim());
                    }
                    rule.setSpecifiedColumns(cols);
                }
                
                // 添加到列表
                rulesList.getItems().add(rule);
                
                // 清空输入
                ruleNameField.clear();
                parameterField.clear();
                
                showInfo("成功", "规则已添加");
            } catch (Exception ex) {
                showError("错误", "添加规则失败: " + ex.getMessage());
            }
        });
        
        configGrid.add(addRuleBtn, 1, row++);
        
        rightPane.getChildren().addAll(configLabel, configGrid);
        
        // 组装界面
        mainPane.setLeft(leftPane);
        mainPane.setCenter(rightPane);
        
        dialog.getDialogPane().setContent(mainPane);
        
        // 处理按钮点击
        dialog.setResultConverter(dialogButton -> {
            if (dialogButton == applyButtonType) {
                // 应用规则
                autoMarkManager.clearRules();
                for (hbnu.project.ergoucsveditior.model.AutoMarkRule rule : rulesList.getItems()) {
                    autoMarkManager.addRule(rule);
                }
                autoMarkManager.applyRules(csvData);
                refreshTable();
                updateStatus("已应用 " + rulesList.getItems().size() + " 条自动标记规则");
                return true;
            } else if (dialogButton == clearAllButtonType) {
                // 清除所有自动标记
                if (showConfirmation("确认清除", "是否清除所有自动标记？\n注意：这不会删除规则，只会清除已应用的标记。")) {
                    autoMarkManager.clearAutoMarks();
                    refreshTable();
                    updateStatus("已清除所有自动标记");
                }
                return false; // 不关闭对话框
            } else if (dialogButton == settingsButtonType) {
                // 显示颜色设置对话框
                showAutoMarkColorSettings();
                return false;
            }
            return false;
        });
        
        // 显示对话框
        dialog.showAndWait();
    }
    
    /**
     * 显示自动标记颜色设置对话框
     */
    private void showAutoMarkColorSettings() {
        Dialog<Boolean> dialog = new Dialog<>();
        dialog.setTitle("自动标记颜色设置");
        dialog.setHeaderText("设置不同类型规则的默认颜色");
        
        ButtonType saveButtonType = new ButtonType("保存", ButtonBar.ButtonData.OK_DONE);
        dialog.getDialogPane().getButtonTypes().addAll(saveButtonType, ButtonType.CANCEL);
        
        javafx.scene.layout.GridPane grid = new javafx.scene.layout.GridPane();
        grid.setHgap(10);
        grid.setVgap(10);
        grid.setPadding(new javafx.geometry.Insets(20));
        
        int row = 0;
        
        // 数字标记颜色
        javafx.scene.control.ColorPicker numberColorPicker = new javafx.scene.control.ColorPicker(
            javafx.scene.paint.Color.web(autoMarkSettings.getNumberMarkColor()));
        grid.add(new Label("数字标记默认颜色:"), 0, row);
        grid.add(numberColorPicker, 1, row++);
        
        // 字符串标记颜色
        javafx.scene.control.ColorPicker stringColorPicker = new javafx.scene.control.ColorPicker(
            javafx.scene.paint.Color.web(autoMarkSettings.getStringMarkColor()));
        grid.add(new Label("字符串标记默认颜色:"), 0, row);
        grid.add(stringColorPicker, 1, row++);
        
        // 格式校验标记颜色
        javafx.scene.control.ColorPicker formatColorPicker = new javafx.scene.control.ColorPicker(
            javafx.scene.paint.Color.web(autoMarkSettings.getFormatMarkColor()));
        grid.add(new Label("格式校验默认颜色:"), 0, row);
        grid.add(formatColorPicker, 1, row++);
        
        // 空值标记颜色
        javafx.scene.control.ColorPicker emptyColorPicker = new javafx.scene.control.ColorPicker(
            javafx.scene.paint.Color.web(autoMarkSettings.getEmptyMarkColor()));
        grid.add(new Label("空值标记默认颜色:"), 0, row);
        grid.add(emptyColorPicker, 1, row++);
        
        dialog.getDialogPane().setContent(grid);
        
        dialog.setResultConverter(dialogButton -> {
            if (dialogButton == saveButtonType) {
                autoMarkSettings.setNumberMarkColor(toHexString(numberColorPicker.getValue()));
                autoMarkSettings.setStringMarkColor(toHexString(stringColorPicker.getValue()));
                autoMarkSettings.setFormatMarkColor(toHexString(formatColorPicker.getValue()));
                autoMarkSettings.setEmptyMarkColor(toHexString(emptyColorPicker.getValue()));
                autoMarkSettings.save();
                showInfo("成功", "颜色设置已保存");
                return true;
            }
            return false;
        });
        
        dialog.showAndWait();
    }
    
    /**
     * 解析规则类型字符串为枚举
     */
    private hbnu.project.ergoucsveditior.model.AutoMarkRule.RuleType parseRuleType(String typeStr) {
        if (typeStr.contains("大于")) return hbnu.project.ergoucsveditior.model.AutoMarkRule.RuleType.NUMBER_GREATER;
        if (typeStr.contains("小于")) return hbnu.project.ergoucsveditior.model.AutoMarkRule.RuleType.NUMBER_LESS;
        if (typeStr.contains("等于")) return hbnu.project.ergoucsveditior.model.AutoMarkRule.RuleType.NUMBER_EQUAL;
        if (typeStr.contains("质数")) return hbnu.project.ergoucsveditior.model.AutoMarkRule.RuleType.NUMBER_PRIME;
        if (typeStr.contains("包含")) return hbnu.project.ergoucsveditior.model.AutoMarkRule.RuleType.STRING_CONTAINS;
        if (typeStr.contains("正则")) return hbnu.project.ergoucsveditior.model.AutoMarkRule.RuleType.STRING_REGEX;
        if (typeStr.contains("邮箱")) return hbnu.project.ergoucsveditior.model.AutoMarkRule.RuleType.FORMAT_EMAIL;
        if (typeStr.contains("手机号")) return hbnu.project.ergoucsveditior.model.AutoMarkRule.RuleType.FORMAT_PHONE;
        if (typeStr.contains("URL")) return hbnu.project.ergoucsveditior.model.AutoMarkRule.RuleType.FORMAT_URL;
        if (typeStr.contains("身份证")) return hbnu.project.ergoucsveditior.model.AutoMarkRule.RuleType.FORMAT_ID_CARD;
        if (typeStr.contains("完全空值")) return hbnu.project.ergoucsveditior.model.AutoMarkRule.RuleType.EMPTY_NULL;
        if (typeStr.contains("空白")) return hbnu.project.ergoucsveditior.model.AutoMarkRule.RuleType.EMPTY_WHITESPACE;
        if (typeStr.contains("长度为0")) return hbnu.project.ergoucsveditior.model.AutoMarkRule.RuleType.EMPTY_ZERO_LENGTH;
        return hbnu.project.ergoucsveditior.model.AutoMarkRule.RuleType.NUMBER_GREATER;
    }
    
    /**
     * 解析应用范围字符串为枚举
     */
    private hbnu.project.ergoucsveditior.model.AutoMarkRule.ApplyScope parseScopeType(String scopeStr) {
        if ("选中列".equals(scopeStr)) return hbnu.project.ergoucsveditior.model.AutoMarkRule.ApplyScope.SELECTED_COLUMN;
        if ("指定列".equals(scopeStr)) return hbnu.project.ergoucsveditior.model.AutoMarkRule.ApplyScope.SPECIFIED_COLUMNS;
        return hbnu.project.ergoucsveditior.model.AutoMarkRule.ApplyScope.ALL_COLUMNS;
    }
    
    /**
     * 清除单个单元格的自动标记
     */
    private void clearAutoMarkForCell(int row, int col) {
        if (autoMarkManager != null) {
            autoMarkManager.clearCellAutoMark(row, col);
        }
    }
    
    /**
     * 清除整行的自动标记
     */
    private void clearAutoMarkForRow(int row) {
        if (autoMarkManager != null && csvData != null) {
            autoMarkManager.clearRowAutoMark(row, csvData.getColumns());
        }
    }
    
    /**
     * 清除整列的自动标记
     */
    private void clearAutoMarkForColumn(int col) {
        if (autoMarkManager != null && csvData != null) {
            autoMarkManager.clearColumnAutoMark(col, csvData.getRows());
        }
    }
    
    /**
     * 显示列统计计算对话框
     */
    private void showColumnStatisticsDialog() {
        Dialog<Void> dialog = new Dialog<>();
        dialog.setTitle("列统计计算");
        dialog.setHeaderText("选择要统计的列并查看计算结果");
        
        // 设置按钮
        ButtonType closeButtonType = new ButtonType("关闭", ButtonBar.ButtonData.CANCEL_CLOSE);
        dialog.getDialogPane().getButtonTypes().add(closeButtonType);
        
        // 创建主面板
        javafx.scene.layout.BorderPane mainPane = new javafx.scene.layout.BorderPane();
        mainPane.setPrefSize(700, 500);
        
        // ========== 左侧列选择 ==========
        javafx.scene.layout.VBox leftPane = new javafx.scene.layout.VBox(10);
        leftPane.setPadding(new javafx.geometry.Insets(10));
        leftPane.setPrefWidth(200);
        
        Label selectLabel = new Label("选择要统计的列：");
        selectLabel.setStyle("-fx-font-weight: bold;");
        
        // 创建列选择下拉框
        ComboBox<String> columnComboBox = new ComboBox<>();
        for (int i = 0; i < csvData.getColumns(); i++) {
            String columnName = "列 " + i;
            // 尝试从第一行获取表头名称
            if (csvData.getRows() > 0) {
                String headerValue = csvData.getCellValue(0, i);
                if (headerValue != null && !headerValue.trim().isEmpty()) {
                    columnName = "列 " + i + " (" + headerValue + ")";
                }
            }
            columnComboBox.getItems().add(columnName);
        }
        columnComboBox.setPromptText("请选择列...");
        columnComboBox.setPrefWidth(180);
        
        Button calculateButton = new Button("计算统计");
        calculateButton.setStyle("-fx-background-color: #2196F3; -fx-text-fill: white;");
        calculateButton.setDisable(true);
        
        columnComboBox.setOnAction(e -> {
            calculateButton.setDisable(columnComboBox.getValue() == null);
        });
        
        // 选项：是否包含表头
        javafx.scene.control.CheckBox includeHeaderCheckBox = new javafx.scene.control.CheckBox("第一行是表头（排除在统计之外）");
        includeHeaderCheckBox.setSelected(true);
        
        leftPane.getChildren().addAll(selectLabel, columnComboBox, includeHeaderCheckBox, calculateButton);
        
        // ========== 右侧统计结果显示 ==========
        javafx.scene.layout.VBox rightPane = new javafx.scene.layout.VBox(10);
        rightPane.setPadding(new javafx.geometry.Insets(10));
        
        Label resultLabel = new Label("统计结果：");
        resultLabel.setStyle("-fx-font-weight: bold;");
        
        javafx.scene.control.TextArea resultArea = new javafx.scene.control.TextArea();
        resultArea.setEditable(false);
        resultArea.setWrapText(true);
        resultArea.setPrefHeight(400);
        resultArea.setPromptText("请选择列并点击\"计算统计\"按钮");
        resultArea.setStyle("-fx-font-family: 'Consolas', 'Monaco', monospace; -fx-font-size: 12px;");
        
        // 复制按钮
        Button copyResultButton = new Button("复制结果");
        copyResultButton.setOnAction(e -> {
            javafx.scene.input.Clipboard clipboard = javafx.scene.input.Clipboard.getSystemClipboard();
            javafx.scene.input.ClipboardContent content = new javafx.scene.input.ClipboardContent();
            content.putString(resultArea.getText());
            clipboard.setContent(content);
            showInfo("成功", "统计结果已复制到剪贴板");
        });
        
        rightPane.getChildren().addAll(resultLabel, resultArea, copyResultButton);
        
        // 计算按钮事件
        calculateButton.setOnAction(e -> {
            String selected = columnComboBox.getValue();
            if (selected != null) {
                // 提取列索引
                int columnIndex = columnComboBox.getSelectionModel().getSelectedIndex();
                boolean excludeHeader = includeHeaderCheckBox.isSelected();
                
                // 计算统计信息
                String statistics = calculateColumnStatistics(columnIndex, excludeHeader);
                resultArea.setText(statistics);
            }
        });
        
        mainPane.setLeft(leftPane);
        mainPane.setCenter(rightPane);
        
        dialog.getDialogPane().setContent(mainPane);
        dialog.showAndWait();
    }
    
    /**
     * 计算指定列的统计信息
     */
    private String calculateColumnStatistics(int columnIndex, boolean excludeHeader) {
        StringBuilder result = new StringBuilder();
        
        int startRow = excludeHeader ? 1 : 0;
        int totalRows = csvData.getRows();
        
        if (startRow >= totalRows) {
            return "没有足够的数据进行统计计算";
        }
        
        // 收集列数据
        java.util.List<String> allValues = new java.util.ArrayList<>();
        java.util.List<Double> numericValues = new java.util.ArrayList<>();
        int emptyCount = 0;
        int nonEmptyCount = 0;
        
        for (int i = startRow; i < totalRows; i++) {
            String value = csvData.getCellValue(i, columnIndex);
            allValues.add(value);
            
            if (value == null || value.trim().isEmpty()) {
                emptyCount++;
            } else {
                nonEmptyCount++;
                // 尝试解析为数字
                try {
                    double numValue = Double.parseDouble(value.trim());
                    numericValues.add(numValue);
                } catch (NumberFormatException e) {
                    // 不是数字，忽略
                }
            }
        }
        
        // 生成统计报告
        result.append("═══════════════════════════════════════\n");
        result.append("            列统计报告\n");
        result.append("═══════════════════════════════════════\n\n");
        
        // 基本信息
        result.append("【基本信息】\n");
        result.append(String.format("  列索引: %d\n", columnIndex));
        result.append(String.format("  总行数: %d\n", totalRows - startRow));
        result.append(String.format("  非空行数: %d\n", nonEmptyCount));
        result.append(String.format("  空行数: %d\n", emptyCount));
        result.append(String.format("  数值行数: %d\n", numericValues.size()));
        result.append("\n");
        
        // 如果有数值数据，计算数值统计
        if (!numericValues.isEmpty()) {
            result.append("【数值统计】\n");
            
            // 排序用于中位数和四分位数计算
            java.util.List<Double> sortedValues = new java.util.ArrayList<>(numericValues);
            java.util.Collections.sort(sortedValues);
            
            // 总和
            double sum = numericValues.stream().mapToDouble(Double::doubleValue).sum();
            result.append(String.format("  总和: %.4f\n", sum));
            
            // 平均值
            double mean = sum / numericValues.size();
            result.append(String.format("  平均值: %.4f\n", mean));
            
            // 中位数
            double median = calculateMedian(sortedValues);
            result.append(String.format("  中位数: %.4f\n", median));
            
            // 最大值和最小值
            double max = sortedValues.get(sortedValues.size() - 1);
            double min = sortedValues.get(0);
            result.append(String.format("  最大值: %.4f\n", max));
            result.append(String.format("  最小值: %.4f\n", min));
            result.append(String.format("  范围: %.4f\n", max - min));
            
            // 四分位数
            double q1 = calculatePercentile(sortedValues, 25);
            double q3 = calculatePercentile(sortedValues, 75);
            result.append(String.format("  第一四分位数 (Q1): %.4f\n", q1));
            result.append(String.format("  第三四分位数 (Q3): %.4f\n", q3));
            result.append(String.format("  四分位距 (IQR): %.4f\n", q3 - q1));
            
            // 标准差和方差
            double variance = calculateVariance(numericValues, mean);
            double stdDev = Math.sqrt(variance);
            result.append(String.format("  方差: %.4f\n", variance));
            result.append(String.format("  标准差: %.4f\n", stdDev));
            
            // 变异系数
            if (mean != 0) {
                double cv = (stdDev / Math.abs(mean)) * 100;
                result.append(String.format("  变异系数: %.2f%%\n", cv));
            }
            
            result.append("\n");
        } else {
            result.append("【数值统计】\n");
            result.append("  该列没有有效的数值数据\n\n");
        }
        
        // 文本统计
        result.append("【文本统计】\n");
        if (nonEmptyCount > 0) {
            // 最长和最短的文本
            String longest = "";
            String shortest = null;
            int totalLength = 0;
            
            for (String value : allValues) {
                if (value != null && !value.trim().isEmpty()) {
                    if (value.length() > longest.length()) {
                        longest = value;
                    }
                    if (shortest == null || value.length() < shortest.length()) {
                        shortest = value;
                    }
                    totalLength += value.length();
                }
            }
            
            result.append(String.format("  平均长度: %.2f 个字符\n", (double) totalLength / nonEmptyCount));
            result.append(String.format("  最长文本长度: %d 个字符\n", longest.length()));
            if (longest.length() <= 50) {
                result.append(String.format("  最长文本: \"%s\"\n", longest));
            } else {
                result.append(String.format("  最长文本: \"%s...\"\n", longest.substring(0, 47) + "..."));
            }
            result.append(String.format("  最短文本长度: %d 个字符\n", shortest != null ? shortest.length() : 0));
            
            // 唯一值统计
            java.util.Set<String> uniqueValues = new java.util.HashSet<>();
            for (String value : allValues) {
                if (value != null && !value.trim().isEmpty()) {
                    uniqueValues.add(value);
                }
            }
            result.append(String.format("  唯一值数量: %d\n", uniqueValues.size()));
            
            // 重复率
            double duplicateRate = (1 - (double) uniqueValues.size() / nonEmptyCount) * 100;
            result.append(String.format("  重复率: %.2f%%\n", duplicateRate));
            
        } else {
            result.append("  该列没有非空文本数据\n");
        }
        
        result.append("\n═══════════════════════════════════════\n");
        
        return result.toString();
    }
    
    /**
     * 计算中位数
     */
    private double calculateMedian(java.util.List<Double> sortedValues) {
        int size = sortedValues.size();
        if (size == 0) {
            return 0.0;
        }
        if (size % 2 == 0) {
            return (sortedValues.get(size / 2 - 1) + sortedValues.get(size / 2)) / 2.0;
        } else {
            return sortedValues.get(size / 2);
        }
    }
    
    /**
     * 计算百分位数
     */
    private double calculatePercentile(java.util.List<Double> sortedValues, double percentile) {
        if (sortedValues.isEmpty()) {
            return 0.0;
        }
        int size = sortedValues.size();
        double index = (percentile / 100.0) * (size - 1);
        int lower = (int) Math.floor(index);
        int upper = (int) Math.ceil(index);
        
        if (lower == upper) {
            return sortedValues.get(lower);
        }
        
        double weight = index - lower;
        return sortedValues.get(lower) * (1 - weight) + sortedValues.get(upper) * weight;
    }
    
    /**
     * 计算方差
     */
    private double calculateVariance(java.util.List<Double> values, double mean) {
        if (values.isEmpty()) {
            return 0.0;
        }
        double sumSquaredDiff = 0.0;
        for (double value : values) {
            double diff = value - mean;
            sumSquaredDiff += diff * diff;
        }
        return sumSquaredDiff / values.size();
    }
    
    /**
     * 显示数据库持久化对话框
     */
    private void showDatabasePersistenceDialog() {
        Dialog<Void> dialog = new Dialog<>();
        dialog.setTitle("数据库持久化");
        dialog.setHeaderText("将CSV数据导入到MySQL数据库");
        
        // 设置按钮
        ButtonType connectButtonType = new ButtonType("测试连接", ButtonBar.ButtonData.OTHER);
        ButtonType importButtonType = new ButtonType("导入数据", ButtonBar.ButtonData.OK_DONE);
        dialog.getDialogPane().getButtonTypes().addAll(connectButtonType, importButtonType, ButtonType.CANCEL);
        
        // 创建主面板
        javafx.scene.layout.BorderPane mainPane = new javafx.scene.layout.BorderPane();
        mainPane.setPrefSize(600, 500);
        
        // ========== 数据库连接配置 ==========
        javafx.scene.layout.VBox configPane = new javafx.scene.layout.VBox(15);
        configPane.setPadding(new javafx.geometry.Insets(20));
        
        Label titleLabel = new Label("数据库连接配置");
        titleLabel.setStyle("-fx-font-weight: bold; -fx-font-size: 14px;");
        
        javafx.scene.layout.GridPane grid = new javafx.scene.layout.GridPane();
        grid.setHgap(10);
        grid.setVgap(10);
        
        int row = 0;
        
        // 主机地址
        TextField hostField = new TextField("localhost");
        hostField.setPromptText("例如: localhost 或 192.168.1.100");
        grid.add(new Label("主机地址:"), 0, row);
        grid.add(hostField, 1, row++);
        
        // 端口
        TextField portField = new TextField("3306");
        portField.setPromptText("默认端口: 3306");
        grid.add(new Label("端口:"), 0, row);
        grid.add(portField, 1, row++);
        
        // 数据库名
        TextField databaseField = new TextField();
        databaseField.setPromptText("要连接的数据库名称");
        grid.add(new Label("数据库名:"), 0, row);
        grid.add(databaseField, 1, row++);
        
        // 用户名
        TextField usernameField = new TextField("root");
        usernameField.setPromptText("数据库用户名");
        grid.add(new Label("用户名:"), 0, row);
        grid.add(usernameField, 1, row++);
        
        // 密码
        javafx.scene.control.PasswordField passwordField = new javafx.scene.control.PasswordField();
        passwordField.setPromptText("数据库密码");
        grid.add(new Label("密码:"), 0, row);
        grid.add(passwordField, 1, row++);
        
        // 分隔线
        javafx.scene.control.Separator separator1 = new javafx.scene.control.Separator();
        
        // ========== 表配置 ==========
        Label tableConfigLabel = new Label("表配置");
        tableConfigLabel.setStyle("-fx-font-weight: bold; -fx-font-size: 14px;");
        
        javafx.scene.layout.GridPane tableGrid = new javafx.scene.layout.GridPane();
        tableGrid.setHgap(10);
        tableGrid.setVgap(10);
        
        int tableRow = 0;
        
        // 表名
        TextField tableNameField = new TextField();
        tableNameField.setPromptText("要创建的表名称");
        tableGrid.add(new Label("表名:"), 0, tableRow);
        tableGrid.add(tableNameField, 1, tableRow++);
        
        // 第一行作为表头
        javafx.scene.control.CheckBox useHeaderCheckBox = new javafx.scene.control.CheckBox("第一行作为列名（表头）");
        useHeaderCheckBox.setSelected(true);
        tableGrid.add(useHeaderCheckBox, 0, tableRow, 2, 1);
        tableRow++;
        
        // 如果表已存在
        ComboBox<String> existingTableCombo = new ComboBox<>();
        existingTableCombo.getItems().addAll("追加数据到现有表", "覆盖现有表", "如果存在则取消");
        existingTableCombo.setValue("如果存在则取消");
        tableGrid.add(new Label("表已存在时:"), 0, tableRow);
        tableGrid.add(existingTableCombo, 1, tableRow++);
        
        // ========== 状态信息显示 ==========
        javafx.scene.control.TextArea statusArea = new javafx.scene.control.TextArea();
        statusArea.setEditable(false);
        statusArea.setWrapText(true);
        statusArea.setPrefHeight(150);
        statusArea.setPromptText("连接状态和导入进度将在此显示");
        statusArea.setStyle("-fx-font-family: 'Consolas', 'Monaco', monospace;");
        
        Label statusLabel = new Label("操作日志:");
        statusLabel.setStyle("-fx-font-weight: bold;");
        
        configPane.getChildren().addAll(
            titleLabel, grid, separator1, 
            tableConfigLabel, tableGrid,
            new javafx.scene.control.Separator(),
            statusLabel, statusArea
        );
        
        mainPane.setCenter(configPane);
        dialog.getDialogPane().setContent(mainPane);
        
        // 测试连接按钮事件
        Button connectButton = (Button) dialog.getDialogPane().lookupButton(connectButtonType);
        connectButton.setOnAction(e -> {
            statusArea.appendText("[" + java.time.LocalDateTime.now().format(
                java.time.format.DateTimeFormatter.ofPattern("HH:mm:ss")) + "] 正在测试数据库连接...\n");
            
            String host = hostField.getText().trim();
            String port = portField.getText().trim();
            String database = databaseField.getText().trim();
            String username = usernameField.getText().trim();
            String password = passwordField.getText();
            
            if (host.isEmpty() || database.isEmpty() || username.isEmpty()) {
                statusArea.appendText("[错误] 请填写完整的连接信息！\n");
                return;
            }
            
            // 在后台线程中测试连接
            new Thread(() -> {
                boolean success = testDatabaseConnection(host, port, database, username, password, statusArea);
                javafx.application.Platform.runLater(() -> {
                    if (success) {
                        statusArea.appendText("[成功] 数据库连接测试成功！\n");
                    } else {
                        statusArea.appendText("[失败] 数据库连接失败，请检查连接信息！\n");
                    }
                });
            }).start();
            
            e.consume(); // 阻止对话框关闭
        });
        
        // 导入数据按钮事件
        Button importButton = (Button) dialog.getDialogPane().lookupButton(importButtonType);
        importButton.setOnAction(e -> {
            String host = hostField.getText().trim();
            String port = portField.getText().trim();
            String database = databaseField.getText().trim();
            String username = usernameField.getText().trim();
            String password = passwordField.getText();
            String tableName = tableNameField.getText().trim();
            boolean useHeader = useHeaderCheckBox.isSelected();
            String existingTableAction = existingTableCombo.getValue();
            
            if (host.isEmpty() || database.isEmpty() || username.isEmpty() || tableName.isEmpty()) {
                statusArea.appendText("[错误] 请填写完整的配置信息！\n");
                e.consume();
                return;
            }
            
            // 确认导入
            if (!showConfirmation("确认导入", 
                    "确定要将CSV数据导入到数据库吗？\n" +
                    "数据库: " + database + "\n" +
                    "表名: " + tableName + "\n" +
                    "数据行数: " + csvData.getRows())) {
                e.consume();
                return;
            }
            
            statusArea.appendText("\n[" + java.time.LocalDateTime.now().format(
                java.time.format.DateTimeFormatter.ofPattern("HH:mm:ss")) + "] 开始导入数据...\n");
            
            // 禁用按钮防止重复点击
            importButton.setDisable(true);
            connectButton.setDisable(true);
            
            // 在后台线程中执行导入
            new Thread(() -> {
                boolean success = importDataToDatabase(
                    host, port, database, username, password, 
                    tableName, useHeader, existingTableAction, statusArea
                );
                
                javafx.application.Platform.runLater(() -> {
                    importButton.setDisable(false);
                    connectButton.setDisable(false);
                    
                    if (success) {
                        statusArea.appendText("[完成] 数据导入成功！\n");
                        showInfo("成功", "CSV数据已成功导入到数据库！");
                    } else {
                        statusArea.appendText("[失败] 数据导入失败！\n");
                        showError("失败", "数据导入失败，请查看操作日志！");
                    }
                });
            }).start();
            
            e.consume(); // 阻止对话框立即关闭
        });
        
        dialog.showAndWait();
    }
    
    /**
     * 测试数据库连接
     */
    private boolean testDatabaseConnection(String host, String port, String database, 
                                          String username, String password, 
                                          javafx.scene.control.TextArea statusArea) {
        String url = String.format("jdbc:mysql://%s:%s/%s?useSSL=false&allowPublicKeyRetrieval=true&serverTimezone=UTC", 
                                  host, port, database);
        
        try {
            Class.forName("com.mysql.cj.jdbc.Driver");
            java.sql.Connection conn = java.sql.DriverManager.getConnection(url, username, password);
            
            javafx.application.Platform.runLater(() -> {
                statusArea.appendText("[连接] 成功连接到: " + host + ":" + port + "/" + database + "\n");
            });
            
            conn.close();
            return true;
        } catch (ClassNotFoundException ex) {
            javafx.application.Platform.runLater(() -> {
                statusArea.appendText("[错误] MySQL驱动未找到: " + ex.getMessage() + "\n");
            });
            return false;
        } catch (java.sql.SQLException ex) {
            javafx.application.Platform.runLater(() -> {
                statusArea.appendText("[错误] 连接失败: " + ex.getMessage() + "\n");
            });
            return false;
        }
    }
    
    /**
     * 将CSV数据导入到数据库
     */
    private boolean importDataToDatabase(String host, String port, String database,
                                         String username, String password,
                                         String tableName, boolean useHeader,
                                         String existingTableAction,
                                         javafx.scene.control.TextArea statusArea) {
        String url = String.format("jdbc:mysql://%s:%s/%s?useSSL=false&allowPublicKeyRetrieval=true&serverTimezone=UTC", 
                                  host, port, database);
        
        java.sql.Connection conn = null;
        try {
            // 加载驱动
            Class.forName("com.mysql.cj.jdbc.Driver");
            
            // 连接数据库
            conn = java.sql.DriverManager.getConnection(url, username, password);
            javafx.application.Platform.runLater(() -> {
                statusArea.appendText("[连接] 数据库连接成功\n");
            });
            
            // 检查表是否存在
            boolean tableExists = checkTableExists(conn, tableName, statusArea);
            
            if (tableExists) {
                javafx.application.Platform.runLater(() -> {
                    statusArea.appendText("[检测] 表 '" + tableName + "' 已存在\n");
                });
                
                if ("如果存在则取消".equals(existingTableAction)) {
                    javafx.application.Platform.runLater(() -> {
                        statusArea.appendText("[取消] 表已存在，取消导入操作\n");
                    });
                    return false;
                } else if ("覆盖现有表".equals(existingTableAction)) {
                    // 删除旧表
                    javafx.application.Platform.runLater(() -> {
                        statusArea.appendText("[操作] 删除现有表...\n");
                    });
                    dropTable(conn, tableName);
                    tableExists = false;
                }
                // "追加数据到现有表" 不需要额外操作
            }
            
            // 如果表不存在，创建新表
            if (!tableExists) {
                createTable(conn, tableName, useHeader, statusArea);
            }
            
            // 插入数据
            insertData(conn, tableName, useHeader, statusArea);
            
            conn.close();
            return true;
            
        } catch (ClassNotFoundException ex) {
            javafx.application.Platform.runLater(() -> {
                statusArea.appendText("[错误] MySQL驱动未找到: " + ex.getMessage() + "\n");
            });
            return false;
        } catch (java.sql.SQLException ex) {
            javafx.application.Platform.runLater(() -> {
                statusArea.appendText("[错误] SQL错误: " + ex.getMessage() + "\n");
            });
            return false;
        } catch (Exception ex) {
            javafx.application.Platform.runLater(() -> {
                statusArea.appendText("[错误] 导入失败: " + ex.getMessage() + "\n");
            });
            return false;
        } finally {
            if (conn != null) {
                try {
                    conn.close();
                } catch (java.sql.SQLException e) {
                    // 忽略关闭错误
                }
            }
        }
    }
    
    /**
     * 检查表是否存在
     */
    private boolean checkTableExists(java.sql.Connection conn, String tableName, 
                                     javafx.scene.control.TextArea statusArea) throws java.sql.SQLException {
        java.sql.DatabaseMetaData meta = conn.getMetaData();
        java.sql.ResultSet rs = meta.getTables(null, null, tableName, new String[]{"TABLE"});
        boolean exists = rs.next();
        rs.close();
        return exists;
    }
    
    /**
     * 删除表
     */
    private void dropTable(java.sql.Connection conn, String tableName) throws java.sql.SQLException {
        String sql = "DROP TABLE IF EXISTS `" + tableName + "`";
        java.sql.Statement stmt = conn.createStatement();
        stmt.executeUpdate(sql);
        stmt.close();
    }
    
    /**
     * 创建表
     */
    private void createTable(java.sql.Connection conn, String tableName, boolean useHeader,
                            javafx.scene.control.TextArea statusArea) throws java.sql.SQLException {
        javafx.application.Platform.runLater(() -> {
            statusArea.appendText("[创建] 正在创建表 '" + tableName + "'...\n");
        });
        
        StringBuilder sql = new StringBuilder("CREATE TABLE `" + tableName + "` (");
        sql.append("`id` INT AUTO_INCREMENT PRIMARY KEY, ");
        
        int columnCount = csvData.getColumns();
        
        for (int i = 0; i < columnCount; i++) {
            String columnName;
            if (useHeader && csvData.getRows() > 0) {
                // 使用第一行作为列名
                columnName = csvData.getCellValue(0, i);
                if (columnName == null || columnName.trim().isEmpty()) {
                    columnName = "column_" + i;
                }
                // 清理列名，移除特殊字符
                columnName = columnName.replaceAll("[^a-zA-Z0-9_\\u4e00-\\u9fa5]", "_");
            } else {
                columnName = "column_" + i;
            }
            
            sql.append("`").append(columnName).append("` TEXT");
            if (i < columnCount - 1) {
                sql.append(", ");
            }
        }
        
        sql.append(")");
        
        java.sql.Statement stmt = conn.createStatement();
        stmt.executeUpdate(sql.toString());
        stmt.close();
        
        javafx.application.Platform.runLater(() -> {
            statusArea.appendText("[成功] 表创建完成\n");
        });
    }
    
    /**
     * 插入数据
     */
    private void insertData(java.sql.Connection conn, String tableName, boolean useHeader,
                           javafx.scene.control.TextArea statusArea) throws java.sql.SQLException {
        javafx.application.Platform.runLater(() -> {
            statusArea.appendText("[导入] 开始插入数据...\n");
        });
        
        int startRow = useHeader ? 1 : 0;
        int totalRows = csvData.getRows();
        int columnCount = csvData.getColumns();
        
        // 构建插入SQL
        StringBuilder sql = new StringBuilder("INSERT INTO `" + tableName + "` (");
        
        // 添加列名
        for (int i = 0; i < columnCount; i++) {
            String columnName;
            if (useHeader && csvData.getRows() > 0) {
                columnName = csvData.getCellValue(0, i);
                if (columnName == null || columnName.trim().isEmpty()) {
                    columnName = "column_" + i;
                }
                columnName = columnName.replaceAll("[^a-zA-Z0-9_\\u4e00-\\u9fa5]", "_");
            } else {
                columnName = "column_" + i;
            }
            sql.append("`").append(columnName).append("`");
            if (i < columnCount - 1) {
                sql.append(", ");
            }
        }
        
        sql.append(") VALUES (");
        for (int i = 0; i < columnCount; i++) {
            sql.append("?");
            if (i < columnCount - 1) {
                sql.append(", ");
            }
        }
        sql.append(")");
        
        java.sql.PreparedStatement pstmt = conn.prepareStatement(sql.toString());
        
        int insertedCount = 0;
        int batchSize = 100; // 每100行提交一次
        
        for (int row = startRow; row < totalRows; row++) {
            for (int col = 0; col < columnCount; col++) {
                String value = csvData.getCellValue(row, col);
                pstmt.setString(col + 1, value);
            }
            pstmt.addBatch();
            insertedCount++;
            
            // 批量提交
            if (insertedCount % batchSize == 0) {
                pstmt.executeBatch();
                int finalInsertedCount = insertedCount;
                javafx.application.Platform.runLater(() -> {
                    statusArea.appendText("[进度] 已插入 " + finalInsertedCount + " / " + (totalRows - startRow) + " 行\n");
                });
            }
        }
        
        // 提交剩余的数据
        if (insertedCount % batchSize != 0) {
            pstmt.executeBatch();
        }
        
        pstmt.close();
        
        int finalInsertedCount = insertedCount;
        javafx.application.Platform.runLater(() -> {
            statusArea.appendText("[完成] 共插入 " + finalInsertedCount + " 行数据\n");
        });
    }
    
    /**
     * 重新应用自动标记规则到指定单元格
     * 当单元格内容被编辑后调用
     */
    private void reapplyAutoMarkForCell(int row, int col) {
        if (autoMarkManager == null || autoMarkManager.getRules().isEmpty()) {
            return;
        }
        
        String cellValue = csvData.getCellValue(row, col);
        
        // 检查所有规则
        boolean matched = false;
        for (hbnu.project.ergoucsveditior.model.AutoMarkRule rule : autoMarkManager.getRules()) {
            if (!rule.isEnabled()) {
                continue;
            }
            
            // 检查列是否在规则应用范围内
            boolean inScope = false;
            switch (rule.getScope()) {
                case ALL_COLUMNS:
                    inScope = true;
                    break;
                case SPECIFIED_COLUMNS:
                    if (rule.getSpecifiedColumns() != null) {
                        for (int specCol : rule.getSpecifiedColumns()) {
                            if (specCol == col) {
                                inScope = true;
                                break;
                            }
                        }
                    }
                    break;
                case SELECTED_COLUMN:
                    // 选中列模式，这里简化为所有列
                    inScope = true;
                    break;
            }
            
            if (!inScope) {
                continue;
            }
            
            // 检查是否匹配规则
            if (autoMarkManager.matchesRule(rule, cellValue)) {
                // 匹配，应用颜色
                autoMarkManager.getAutoMarkColor(row, col); // 触发内部更新
                matched = true;
                break; // 只应用第一个匹配的规则
            }
        }
        
        // 如果不匹配任何规则，清除该单元格的自动标记
        if (!matched) {
            // 重新应用所有规则以更新autoMarkColors
            autoMarkManager.applyRules(csvData);
        }
    }
    
    /**
     * 显示快捷键设置对话框
     */
    private void showKeyBindingsDialog() {
        Dialog<Boolean> dialog = new Dialog<>();
        dialog.setTitle("快捷键设置");
        dialog.setHeaderText("自定义快捷键");
        
        // 设置按钮
        ButtonType saveButtonType = new ButtonType("保存", ButtonBar.ButtonData.OK_DONE);
        dialog.getDialogPane().getButtonTypes().addAll(saveButtonType, ButtonType.CANCEL);
        
        // 创建快捷键设置表单
        javafx.scene.layout.GridPane grid = new javafx.scene.layout.GridPane();
        grid.setHgap(10);
        grid.setVgap(10);
        grid.setPadding(new javafx.geometry.Insets(20, 150, 10, 10));
        
        // 获取所有操作并创建输入字段
        String[] actions = hbnu.project.ergoucsveditior.model.KeyBindings.getAllActions();
        java.util.Map<String, TextField> fields = new java.util.HashMap<>();
        
        int row = 0;
        for (String action : actions) {
            Label label = new Label(hbnu.project.ergoucsveditior.model.KeyBindings.getActionDisplayName(action) + ":");
            TextField field = new TextField();
            
            javafx.scene.input.KeyCombination binding = keyBindings.getBinding(action);
            if (binding != null) {
                field.setText(binding.toString());
            }
            
            field.setPromptText("点击并按下快捷键组合");
            field.setEditable(false);
            
            // 监听键盘事件来捕获快捷键
            field.setOnKeyPressed(event -> {
                if (event.getCode() != javafx.scene.input.KeyCode.UNDEFINED) {
                    StringBuilder sb = new StringBuilder();
                    if (event.isControlDown()) {
                        sb.append("Ctrl+");
                    }
                    if (event.isShiftDown()) {
                        sb.append("Shift+");
                    }
                    if (event.isAltDown()) {
                        sb.append("Alt+");
                    }
                    
                    javafx.scene.input.KeyCode code = event.getCode();
                    if (code != javafx.scene.input.KeyCode.CONTROL && 
                        code != javafx.scene.input.KeyCode.SHIFT && 
                        code != javafx.scene.input.KeyCode.ALT) {
                        sb.append(code.getName());
                        field.setText(sb.toString());
                    }
                    event.consume();
                }
            });
            
            // 添加清除按钮
            Button clearButton = new Button("清除");
            clearButton.setOnAction(e -> field.setText(""));
            
            grid.add(label, 0, row);
            grid.add(field, 1, row);
            grid.add(clearButton, 2, row);
            
            fields.put(action, field);
            row++;
        }
        
        javafx.scene.control.ScrollPane scrollPane = new javafx.scene.control.ScrollPane(grid);
        scrollPane.setFitToWidth(true);
        scrollPane.setPrefHeight(400);
        
        dialog.getDialogPane().setContent(scrollPane);
        
        // 转换结果
        dialog.setResultConverter(dialogButton -> {
            if (dialogButton == saveButtonType) {
                try {
                    for (java.util.Map.Entry<String, TextField> entry : fields.entrySet()) {
                        String action = entry.getKey();
                        String keyString = entry.getValue().getText();
                        
                        if (keyString != null && !keyString.trim().isEmpty()) {
                            try {
                                javafx.scene.input.KeyCombination keyCombination = 
                                    javafx.scene.input.KeyCombination.valueOf(keyString);
                                keyBindings.setBinding(action, keyCombination);
                            } catch (IllegalArgumentException e) {
                                // 忽略无效的快捷键
                            }
                        } else {
                            keyBindings.removeBinding(action);
                        }
                    }
                    keyBindings.save();
                    
                    // 重新应用快捷键
                    setupKeyboardShortcuts();
                    
                    return true;
                } catch (Exception e) {
                    javafx.application.Platform.runLater(() -> 
                        showError("保存失败", "保存快捷键设置时发生错误: " + e.getMessage())
                    );
                }
            }
            return false;
        });
        
        Optional<Boolean> result = dialog.showAndWait();
        if (result.isPresent() && result.get()) {
            updateStatus("快捷键设置已保存");
        }
    }
    
    /**
     * 显示文本颜色选择器对话框
     */
    private void showTextColorPickerDialog(int row, int col, String type) {
        Dialog<javafx.scene.paint.Color> dialog = new Dialog<>();
        dialog.setTitle("选择文本颜色");
        dialog.setHeaderText("为" + ("cell".equals(type) ? "单元格" : "row".equals(type) ? "行" : "列") + "选择文本颜色");
        
        ButtonType applyButtonType = new ButtonType("应用", ButtonBar.ButtonData.OK_DONE);
        dialog.getDialogPane().getButtonTypes().addAll(applyButtonType, ButtonType.CANCEL);
        
        javafx.scene.layout.VBox vbox = new javafx.scene.layout.VBox(15);
        vbox.setPadding(new javafx.geometry.Insets(20));
        
        // 颜色选择器
        javafx.scene.control.ColorPicker colorPicker = new javafx.scene.control.ColorPicker(javafx.scene.paint.Color.BLACK);
        
        // 检查是否已有文本颜色设置
        if ("cell".equals(type)) {
            HighlightInfo cellInfo = highlightManager.getCellHighlight(row, col);
            if (cellInfo != null && cellInfo.getTextColor() != null) {
                colorPicker.setValue(cellInfo.getTextColor());
            }
        }
        
        vbox.getChildren().addAll(new Label("选择文本颜色:"), colorPicker);
        dialog.getDialogPane().setContent(vbox);
        
        dialog.setResultConverter(dialogButton -> {
            if (dialogButton == applyButtonType) {
                return colorPicker.getValue();
            }
            return null;
        });
        
        Optional<javafx.scene.paint.Color> result = dialog.showAndWait();
        result.ifPresent(textColor -> {
            if ("cell".equals(type)) {
                // 获取现有的背景色，如果没有则使用默认色
                HighlightInfo cellInfo = highlightManager.getCellHighlight(row, col);
                javafx.scene.paint.Color bgColor = cellInfo != null ? 
                    cellInfo.getBackgroundColor() : highlightManager.getDefaultCellHighlightColor();
                highlightManager.setCellHighlight(row, col, bgColor, textColor);
                updateStatus("已设置单元格文本颜色: 行 " + (row + 1) + ", 列 " + (col + 1));
            } else if (type.equals("row")) {
                HighlightInfo rowInfo = highlightManager.getRowHighlight(row);
                javafx.scene.paint.Color bgColor = rowInfo != null ? 
                    rowInfo.getBackgroundColor() : highlightManager.getDefaultRowHighlightColor();
                highlightManager.setRowHighlight(row, bgColor, textColor);
                updateStatus("已设置第 " + (row + 1) + " 行文本颜色");
            } else if (type.equals("column")) {
                HighlightInfo colInfo = highlightManager.getColumnHighlight(col);
                javafx.scene.paint.Color bgColor = colInfo != null ? 
                    colInfo.getBackgroundColor() : highlightManager.getDefaultColumnHighlightColor();
                highlightManager.setColumnHighlight(col, bgColor, textColor);
                updateStatus("已设置第 " + (col + 1) + " 列文本颜色");
            }
            tableView.refresh();
        });
    }
    
    /**
     * 显示颜色选择器对话框
     */
    private void showColorPickerDialog(int row, int col, String type) {
        Dialog<javafx.scene.paint.Color> dialog = new Dialog<>();
        dialog.setTitle("选择高亮颜色");
        dialog.setHeaderText("为" + (type.equals("cell") ? "单元格" : type.equals("row") ? "行" : "列") + "选择高亮颜色");
        
        ButtonType applyButtonType = new ButtonType("应用", ButtonBar.ButtonData.OK_DONE);
        dialog.getDialogPane().getButtonTypes().addAll(applyButtonType, ButtonType.CANCEL);
        
        javafx.scene.layout.VBox vbox = new javafx.scene.layout.VBox(15);
        vbox.setPadding(new javafx.geometry.Insets(20));
        
        // 颜色选择器
        javafx.scene.control.ColorPicker colorPicker = new javafx.scene.control.ColorPicker();
        
        // 设置默认颜色
        if (type.equals("cell")) {
            colorPicker.setValue(highlightManager.getDefaultCellHighlightColor());
        } else if (type.equals("row")) {
            colorPicker.setValue(highlightManager.getDefaultRowHighlightColor());
        } else if (type.equals("column")) {
            colorPicker.setValue(highlightManager.getDefaultColumnHighlightColor());
        }
        
        // 预设颜色按钮
        javafx.scene.layout.FlowPane presetColors = new javafx.scene.layout.FlowPane(10, 10);
        Label presetLabel = new Label("快速选择:");
        
        javafx.scene.paint.Color[] colors = {
            javafx.scene.paint.Color.rgb(255, 255, 153, 0.7),  // 黄色
            javafx.scene.paint.Color.rgb(255, 200, 200, 0.7),  // 粉红色
            javafx.scene.paint.Color.rgb(200, 255, 200, 0.7),  // 浅绿色
            javafx.scene.paint.Color.rgb(200, 200, 255, 0.7),  // 浅蓝色
            javafx.scene.paint.Color.rgb(255, 200, 100, 0.7),  // 橙色
            javafx.scene.paint.Color.rgb(200, 150, 255, 0.7)   // 紫色
        };
        
        for (javafx.scene.paint.Color color : colors) {
            Button colorBtn = new Button();
            colorBtn.setPrefSize(30, 30);
            colorBtn.setStyle(String.format("-fx-background-color: rgba(%d, %d, %d, %.2f);",
                (int)(color.getRed() * 255),
                (int)(color.getGreen() * 255),
                (int)(color.getBlue() * 255),
                color.getOpacity()));
            colorBtn.setOnAction(e -> colorPicker.setValue(color));
            presetColors.getChildren().add(colorBtn);
        }
        
        vbox.getChildren().addAll(new Label("自定义颜色:"), colorPicker, presetLabel, presetColors);
        dialog.getDialogPane().setContent(vbox);
        
        dialog.setResultConverter(dialogButton -> {
            if (dialogButton == applyButtonType) {
                return colorPicker.getValue();
            }
            return null;
        });
        
        Optional<javafx.scene.paint.Color> result = dialog.showAndWait();
        result.ifPresent(color -> {
            if ("cell".equals(type)) {
                highlightManager.setCellHighlight(row, col, color);
                updateStatus("已标记单元格: 行 " + (row + 1) + ", 列 " + (col + 1));
            } else if ("row".equals(type)) {
                highlightManager.setRowHighlight(row, color);
                updateStatus("已标记第 " + (row + 1) + " 行");
            } else if ("column".equals(type)) {
                highlightManager.setColumnHighlight(col, color);
                updateStatus("已标记第 " + (col + 1) + " 列");
            }
            tableView.refresh();
        });
    }
    
    /**
     * 显示搜索对话框
     */
    private void showSearchDialog() {
        Dialog<Void> dialog = new Dialog<>();
        dialog.setTitle("搜索");
        dialog.setHeaderText("查找表格内容");
        
        // 设置按钮
        ButtonType closeButtonType = new ButtonType("关闭", ButtonBar.ButtonData.CANCEL_CLOSE);
        dialog.getDialogPane().getButtonTypes().add(closeButtonType);
        
        // 创建搜索表单
        javafx.scene.layout.VBox vbox = new javafx.scene.layout.VBox(10);
        vbox.setPadding(new javafx.geometry.Insets(20));
        
        // 搜索文本框
        javafx.scene.layout.HBox searchBox = new javafx.scene.layout.HBox(10);
        Label searchLabel = new Label("查找内容:");
        TextField searchField = new TextField();
        searchField.setPrefWidth(300);
        searchField.setText(lastSearchText);
        searchBox.getChildren().addAll(searchLabel, searchField);
        
        // 选项
        CheckBox caseSensitiveCheck = new CheckBox("区分大小写");
        caseSensitiveCheck.setSelected(lastSearchCaseSensitive);
        
        CheckBox fuzzyMatchCheck = new CheckBox("模糊匹配");
        fuzzyMatchCheck.setSelected(true);
        
        CheckBox regexMatchCheck = new CheckBox("正则表达式");
        regexMatchCheck.setSelected(false);
        
        // 选项说明
        Label infoLabel = new Label("提示：支持多个值搜索（用逗号或分号分隔）");
        infoLabel.setStyle("-fx-text-fill: #888888; -fx-font-size: 11px;");
        
        // 搜索结果标签
        Label resultLabel = new Label("");
        resultLabel.setStyle("-fx-text-fill: #666666;");
        
        // 按钮
        javafx.scene.layout.HBox buttonBox = new javafx.scene.layout.HBox(10);
        Button searchButton = new Button("查找全部");
        Button nextButton = new Button("下一个");
        Button prevButton = new Button("上一个");
        Button clearButton = new Button("清除搜索");
        
        nextButton.setDisable(true);
        prevButton.setDisable(true);
        
        buttonBox.getChildren().addAll(searchButton, nextButton, prevButton, clearButton);
        
        vbox.getChildren().addAll(searchBox, caseSensitiveCheck, fuzzyMatchCheck, 
                                   regexMatchCheck, infoLabel, resultLabel, buttonBox);
        dialog.getDialogPane().setContent(vbox);
        
        // 搜索按钮事件
        searchButton.setOnAction(e -> {
            String searchText = searchField.getText();
            if (searchText == null || searchText.trim().isEmpty()) {
                resultLabel.setText("请输入搜索内容");
                return;
            }
            
            lastSearchText = searchText;
            lastSearchCaseSensitive = caseSensitiveCheck.isSelected();
            
            // 执行搜索
            try {
                searchResults = performSearch(searchText, caseSensitiveCheck.isSelected(), 
                                             fuzzyMatchCheck.isSelected(), 
                                             regexMatchCheck.isSelected());
            } catch (java.util.regex.PatternSyntaxException ex) {
                resultLabel.setText("正则表达式语法错误: " + ex.getMessage());
                return;
            }
            
            if (searchResults.isEmpty()) {
                resultLabel.setText("未找到匹配结果");
                nextButton.setDisable(true);
                prevButton.setDisable(true);
                currentSearchIndex = -1;
            } else {
                resultLabel.setText("找到 " + searchResults.size() + " 个结果");
                nextButton.setDisable(false);
                prevButton.setDisable(false);
                currentSearchIndex = 0;
                // 跳转到第一个结果
                jumpToSearchResult(searchResults.get(0));
            }
        });
        
        // 下一个按钮事件
        nextButton.setOnAction(e -> {
            if (searchResults != null && !searchResults.isEmpty()) {
                currentSearchIndex = (currentSearchIndex + 1) % searchResults.size();
                jumpToSearchResult(searchResults.get(currentSearchIndex));
                resultLabel.setText("结果 " + (currentSearchIndex + 1) + " / " + searchResults.size());
            }
        });
        
        // 上一个按钮事件
        prevButton.setOnAction(e -> {
            if (searchResults != null && !searchResults.isEmpty()) {
                currentSearchIndex = (currentSearchIndex - 1 + searchResults.size()) % searchResults.size();
                jumpToSearchResult(searchResults.get(currentSearchIndex));
                resultLabel.setText("结果 " + (currentSearchIndex + 1) + " / " + searchResults.size());
            }
        });
        
        // 清除按钮事件
        clearButton.setOnAction(e -> {
            // 清除搜索结果
            searchResults = null;
            currentSearchIndex = -1;
            lastSearchText = "";
            
            // 清空搜索框
            searchField.setText("");
            
            // 刷新表格以移除高亮
            tableView.refresh();
            
            // 更新界面
            resultLabel.setText("已清除搜索");
            nextButton.setDisable(true);
            prevButton.setDisable(true);
            
            updateStatus("已清除搜索高亮");
        });
        
        // Enter键触发搜索
        searchField.setOnAction(e -> searchButton.fire());
        
        // 添加对话框关闭监听器，关闭时清除搜索高亮
        dialog.setOnCloseRequest(event -> {
            searchResults = null;
            currentSearchIndex = -1;
            lastSearchText = "";
            tableView.refresh();
            updateStatus("就绪");
        });
        
        dialog.show();
    }
    
    /**
     * 执行搜索
     */
    private java.util.List<SearchResult> performSearch(String searchText, boolean caseSensitive, 
                                                       boolean fuzzy, boolean useRegex) {
        java.util.List<SearchResult> results = new java.util.ArrayList<>();
        
        if (useRegex) {
            // 正则表达式搜索
            try {
                int flags = caseSensitive ? 0 : java.util.regex.Pattern.CASE_INSENSITIVE;
                java.util.regex.Pattern pattern = java.util.regex.Pattern.compile(searchText, flags);
                
                for (int row = 0; row < csvData.getRows(); row++) {
                    for (int col = 0; col < csvData.getColumns(); col++) {
                        String cellValue = csvData.getCellValue(row, col);
                        if (cellValue == null) {
                            continue;
                        }
                        
                        java.util.regex.Matcher matcher = pattern.matcher(cellValue);
                        if (matcher.find()) {
                            results.add(new SearchResult(row, col, cellValue));
                        }
                    }
                }
            } catch (java.util.regex.PatternSyntaxException e) {
                throw e; // 重新抛出异常，让调用方处理
            }
        } else {
            // 普通搜索：支持多个搜索值，用逗号或分号分隔
            String[] searchTerms = searchText.split("[,;，；]");
            
            for (int row = 0; row < csvData.getRows(); row++) {
                for (int col = 0; col < csvData.getColumns(); col++) {
                    String cellValue = csvData.getCellValue(row, col);
                    if (cellValue == null) {
                        continue;
                    }
                    
                    // 检查是否匹配任一搜索词
                    for (String term : searchTerms) {
                        term = term.trim();
                        if (term.isEmpty()) {
                            continue;
                        }
                        
                        boolean matches = false;
                        if (fuzzy) {
                            // 模糊匹配
                            if (caseSensitive) {
                                matches = cellValue.contains(term);
                            } else {
                                matches = cellValue.toLowerCase().contains(term.toLowerCase());
                            }
                        } else {
                            // 精确匹配
                            if (caseSensitive) {
                                matches = cellValue.equals(term);
                            } else {
                                matches = cellValue.equalsIgnoreCase(term);
                            }
                        }
                        
                        if (matches) {
                            results.add(new SearchResult(row, col, cellValue));
                            break; // 同一单元格只记录一次
                        }
                    }
                }
            }
        }
        
        return results;
    }
    
    /**
     * 跳转到搜索结果
     */
    private void jumpToSearchResult(SearchResult result) {
        // 选中对应的行
        tableView.getSelectionModel().select(result.row);
        
        // 滚动到该行
        tableView.scrollTo(result.row);
        
        // 设置焦点到对应的单元格（列索引+1是因为有行号列）
        tableView.getFocusModel().focus(result.row, tableView.getColumns().get(result.column + 1));
        
        // 高亮显示
        tableView.requestFocus();
        
        updateStatus("跳转到: 行 " + (result.row + 1) + ", 列 " + (result.column + 1) + 
                    " - " + result.value);
    }
    
    /**
     * 搜索结果类
     */
    private static class SearchResult {
        int row;
        int column;
        String value;
        
        SearchResult(int row, int column, String value) {
            this.row = row;
            this.column = column;
            this.value = value;
        }
    }
    
    /**
     * 创建单元格右键菜单
     */
    private ContextMenu createCellContextMenu(int rowIndex) {
        ContextMenu contextMenu = new ContextMenu();
        
        // 基本操作
        MenuItem copyItem = new MenuItem("复制");
        copyItem.setOnAction(e -> handleCopy());
        
        MenuItem pasteItem = new MenuItem("粘贴");
        pasteItem.setOnAction(e -> handlePaste());
        
        MenuItem clearCellItem = new MenuItem("清除内容");
        clearCellItem.setOnAction(e -> handleClearCell());
        
        SeparatorMenuItem separator1 = new SeparatorMenuItem();
        
        // 高亮操作
        MenuItem highlightCellItem = new MenuItem("标记单元格颜色");
        highlightCellItem.setOnAction(e -> {
            @SuppressWarnings("unchecked")
            TablePosition<ObservableList<CSVCell>, ?> focusedCell = 
                (TablePosition<ObservableList<CSVCell>, ?>) tableView.getFocusModel().getFocusedCell();
            if (focusedCell != null && focusedCell.getColumn() > 0) {
                showColorPickerDialog(rowIndex, focusedCell.getColumn() - 1, "cell");
            }
        });
        
        MenuItem highlightCellTextItem = new MenuItem("设置单元格文本颜色");
        highlightCellTextItem.setOnAction(e -> {
            @SuppressWarnings("unchecked")
            TablePosition<ObservableList<CSVCell>, ?> focusedCell = 
                (TablePosition<ObservableList<CSVCell>, ?>) tableView.getFocusModel().getFocusedCell();
            if (focusedCell != null && focusedCell.getColumn() > 0) {
                showTextColorPickerDialog(rowIndex, focusedCell.getColumn() - 1, "cell");
            }
        });
        
        MenuItem highlightRowItem = new MenuItem("标记整行颜色");
        highlightRowItem.setOnAction(e -> {
            showColorPickerDialog(rowIndex, 0, "row");
        });
        
        MenuItem highlightColumnItem = new MenuItem("标记整列颜色");
        highlightColumnItem.setOnAction(e -> {
            @SuppressWarnings("unchecked")
            TablePosition<ObservableList<CSVCell>, ?> focusedCell = 
                (TablePosition<ObservableList<CSVCell>, ?>) tableView.getFocusModel().getFocusedCell();
            if (focusedCell != null && focusedCell.getColumn() > 0) {
                showColorPickerDialog(rowIndex, focusedCell.getColumn() - 1, "column");
            }
        });
        
        SeparatorMenuItem separator2 = new SeparatorMenuItem();
        
        MenuItem clearCellHighlightItem = new MenuItem("清除单元格背景色");
        clearCellHighlightItem.setOnAction(e -> {
            @SuppressWarnings("unchecked")
            TablePosition<ObservableList<CSVCell>, ?> focusedCell = 
                (TablePosition<ObservableList<CSVCell>, ?>) tableView.getFocusModel().getFocusedCell();
            if (focusedCell != null && focusedCell.getColumn() > 0) {
                int colIndex = focusedCell.getColumn() - 1;
                highlightManager.clearCellHighlight(rowIndex, colIndex);
                // 同时清除该单元格的自动标记
                clearAutoMarkForCell(rowIndex, colIndex);
                tableView.refresh();
            }
        });
        
        MenuItem clearCellTextColorItem = new MenuItem("清除单元格文本颜色");
        clearCellTextColorItem.setOnAction(e -> {
            @SuppressWarnings("unchecked")
            TablePosition<ObservableList<CSVCell>, ?> focusedCell = 
                (TablePosition<ObservableList<CSVCell>, ?>) tableView.getFocusModel().getFocusedCell();
            if (focusedCell != null && focusedCell.getColumn() > 0) {
                int col = focusedCell.getColumn() - 1;
                // 获取现有的背景色，保留背景色，只清除文本颜色
                HighlightInfo cellInfo = highlightManager.getCellHighlight(rowIndex, col);
                if (cellInfo != null) {
                    highlightManager.setCellHighlight(rowIndex, col, cellInfo.getBackgroundColor(), null);
                    tableView.refresh();
                    updateStatus("已清除单元格文本颜色");
                }
            }
        });
        
        MenuItem clearRowHighlightItem = new MenuItem("清除行高亮");
        clearRowHighlightItem.setOnAction(e -> {
            highlightManager.clearRowHighlight(rowIndex);
            // 同时清除该行的所有自动标记
            clearAutoMarkForRow(rowIndex);
            tableView.refresh();
        });
        
        MenuItem clearColumnHighlightItem = new MenuItem("清除列高亮");
        clearColumnHighlightItem.setOnAction(e -> {
            @SuppressWarnings("unchecked")
            TablePosition<ObservableList<CSVCell>, ?> focusedCell = 
                (TablePosition<ObservableList<CSVCell>, ?>) tableView.getFocusModel().getFocusedCell();
            if (focusedCell != null && focusedCell.getColumn() > 0) {
                int colIndex = focusedCell.getColumn() - 1;
                highlightManager.clearColumnHighlight(colIndex);
                // 同时清除该列的所有自动标记
                clearAutoMarkForColumn(colIndex);
                tableView.refresh();
            }
        });
        
        contextMenu.getItems().addAll(
            copyItem,
            pasteItem,
            clearCellItem,
            separator1,
            highlightCellItem,
            highlightCellTextItem,
            highlightRowItem,
            highlightColumnItem,
            separator2,
            clearCellHighlightItem,
            clearCellTextColorItem,
            clearRowHighlightItem,
            clearColumnHighlightItem
        );
        
        return contextMenu;
    }
    
    /**
     * 支持多行的表格单元格（支持高亮显示）
     */
    private class MultiLineTableCell extends TableCell<ObservableList<CSVCell>, String> {
        private TextArea textArea;
        private javafx.scene.text.TextFlow textFlow;
        private int columnIndex;
        
        public MultiLineTableCell(int columnIndex) {
            this.columnIndex = columnIndex;
            textFlow = new javafx.scene.text.TextFlow();
            textFlow.setMaxWidth(Double.MAX_VALUE);
            
            textArea = new TextArea();
            textArea.setWrapText(true);
            textArea.setPrefRowCount(3);
            
            // 设置键盘事件
            textArea.setOnKeyPressed(event -> {
                // 检查是否匹配单元格换行快捷键（默认Ctrl+Enter）
                javafx.scene.input.KeyCombination cellNewlineBinding = 
                    keyBindings.getBinding(hbnu.project.ergoucsveditior.model.KeyBindings.ACTION_CELL_NEWLINE);
                
                if (cellNewlineBinding != null && cellNewlineBinding.match(event)) {
                    // Ctrl+Enter 换行
                    int caretPos = textArea.getCaretPosition();
                    textArea.insertText(caretPos, settings.getLineEndingString());
                    event.consume();
                } else if (event.getCode() == javafx.scene.input.KeyCode.ENTER && !event.isControlDown() && !event.isShiftDown()) {
                    // 普通Enter键提交编辑
                    commitEdit(textArea.getText());
                    event.consume();
                } else if (event.getCode() == javafx.scene.input.KeyCode.ESCAPE) {
                    // ESC键取消编辑
                    cancelEdit();
                    event.consume();
                } else if (event.getCode() == javafx.scene.input.KeyCode.TAB) {
                    // Tab键也提交编辑并移到下一个单元格
                    commitEdit(textArea.getText());
                    event.consume();
                }
            });
            
            // 支持粘贴多行内容并保持格式
            textArea.setOnKeyReleased(event -> {
                // 粘贴操作在keyReleased时处理，确保内容已经插入
                if (event.isControlDown() && event.getCode() == javafx.scene.input.KeyCode.V) {
                    // 粘贴的内容已经自动插入到TextArea中，保持原有换行符
                    // TextArea会自动处理粘贴的多行文本
                }
            });
        }
        
        @Override
        protected void updateItem(String item, boolean empty) {
            super.updateItem(item, empty);
            
            if (empty || item == null) {
                setText(null);
                setGraphic(null);
                setStyle("");
            } else {
                int rowIndex = getIndex();
                
                // 先获取手动高亮颜色
                javafx.scene.paint.Color backgroundColor = highlightManager.getFinalHighlightColor(rowIndex, columnIndex);
                javafx.scene.paint.Color textColor = highlightManager.getFinalTextColor(rowIndex, columnIndex);
                
                // 如果没有手动高亮，检查自动标记颜色
                if (backgroundColor == null) {
                    String autoMarkColorHex = autoMarkManager.getAutoMarkColor(rowIndex, columnIndex);
                    if (autoMarkColorHex != null && !autoMarkColorHex.isEmpty()) {
                        try {
                            backgroundColor = javafx.scene.paint.Color.web(autoMarkColorHex);
                        } catch (Exception e) {
                            // 颜色解析失败，忽略
                        }
                    }
                }
                
                StringBuilder styleBuilder = new StringBuilder();
                if (backgroundColor != null) {
                    styleBuilder.append(String.format("-fx-background-color: rgba(%d, %d, %d, %.2f);",
                        (int)(backgroundColor.getRed() * 255),
                        (int)(backgroundColor.getGreen() * 255),
                        (int)(backgroundColor.getBlue() * 255),
                        backgroundColor.getOpacity()));
                }
                if (textColor != null) {
                    styleBuilder.append(String.format("-fx-text-fill: rgba(%d, %d, %d, %.2f);",
                        (int)(textColor.getRed() * 255),
                        (int)(textColor.getGreen() * 255),
                        (int)(textColor.getBlue() * 255),
                        textColor.getOpacity()));
                }
                setStyle(styleBuilder.toString());
                
                if (isEditing()) {
                    textArea.setText(item);
                    setText(null);
                    setGraphic(textArea);
                    textArea.requestFocus();
                } else {
                    // 显示文本，支持搜索结果高亮
                    displayTextWithHighlight(item, rowIndex, textColor);
                }
            }
        }
        
        /**
         * 显示带高亮的文本
         */
        private void displayTextWithHighlight(String text, int rowIndex, javafx.scene.paint.Color textColor) {
            textFlow.getChildren().clear();
            
            // 检查是否需要高亮搜索结果
            boolean hasSearchHighlight = false;
            if (searchResults != null && !searchResults.isEmpty() && lastSearchText != null && !lastSearchText.isEmpty()) {
                for (SearchResult result : searchResults) {
                    if (result.row == rowIndex && result.column == columnIndex) {
                        hasSearchHighlight = true;
                        break;
                    }
                }
            }
            
            if (hasSearchHighlight && lastSearchText != null) {
                // 高亮搜索匹配的文本
                highlightSearchText(text, lastSearchText, textColor);
            } else {
                // 普通显示
                javafx.scene.text.Text textNode = new javafx.scene.text.Text(text);
                if (textColor != null) {
                    textNode.setFill(textColor);
                }
                textFlow.getChildren().add(textNode);
            }
            
            setText(null);
            setGraphic(textFlow);
        }
        
        /**
         * 高亮搜索文本
         */
        private void highlightSearchText(String fullText, String searchText, javafx.scene.paint.Color defaultTextColor) {
            if (fullText == null || searchText == null || searchText.isEmpty()) {
                javafx.scene.text.Text textNode = new javafx.scene.text.Text(fullText);
                if (defaultTextColor != null) {
                    textNode.setFill(defaultTextColor);
                }
                textFlow.getChildren().add(textNode);
                return;
            }
            
            String lowerFullText = fullText.toLowerCase();
            String lowerSearchText = searchText.toLowerCase();
            
            int lastIndex = 0;
            int index = lowerFullText.indexOf(lowerSearchText);
            
            while (index >= 0) {
                // 添加匹配前的文本
                if (index > lastIndex) {
                    javafx.scene.text.Text beforeText = new javafx.scene.text.Text(
                        fullText.substring(lastIndex, index));
                    if (defaultTextColor != null) {
                        beforeText.setFill(defaultTextColor);
                    }
                    textFlow.getChildren().add(beforeText);
                }
                
                // 添加匹配的文本（高亮）
                javafx.scene.text.Text matchText = new javafx.scene.text.Text(
                    fullText.substring(index, index + searchText.length()));
                matchText.setStyle("-fx-fill: #000000; -fx-font-weight: bold;");
                matchText.setFill(javafx.scene.paint.Color.BLACK);
                
                // 使用背景高亮（从设置中获取搜索高亮颜色）
                javafx.scene.paint.Color searchColor = javafx.scene.paint.Color.web(settings.getSearchHighlightColor());
                javafx.scene.layout.StackPane highlightPane = new javafx.scene.layout.StackPane(matchText);
                highlightPane.setStyle(String.format("-fx-background-color: rgba(%d, %d, %d, %.2f); -fx-padding: 1px;",
                    (int)(searchColor.getRed() * 255),
                    (int)(searchColor.getGreen() * 255),
                    (int)(searchColor.getBlue() * 255),
                    searchColor.getOpacity()));
                textFlow.getChildren().add(highlightPane);
                
                lastIndex = index + searchText.length();
                index = lowerFullText.indexOf(lowerSearchText, lastIndex);
            }
            
            // 添加剩余文本
            if (lastIndex < fullText.length()) {
                javafx.scene.text.Text afterText = new javafx.scene.text.Text(
                    fullText.substring(lastIndex));
                if (defaultTextColor != null) {
                    afterText.setFill(defaultTextColor);
                }
                textFlow.getChildren().add(afterText);
            }
        }
        
        @Override
        public void startEdit() {
            super.startEdit();
            String text = getItem();
            textArea.setText(text != null ? text : "");
            setText(null);
            setGraphic(textArea);
            textArea.selectAll();
            textArea.requestFocus();
        }
        
        @Override
        public void cancelEdit() {
            super.cancelEdit();
            String text = getItem();
            if (text != null) {
                int rowIndex = getIndex();
                javafx.scene.paint.Color textColor = highlightManager.getFinalTextColor(rowIndex, columnIndex);
                displayTextWithHighlight(text, rowIndex, textColor);
            }
        }
    }
    
    /**
     * 应用主题
     */
    private void applyTheme() {
        if (tableView != null && tableView.getScene() != null) {
            javafx.scene.Parent root = tableView.getScene().getRoot();
            
            // 应用深色/浅色主题
            if ("深色".equals(settings.getTheme())) {
                root.getStyleClass().add("theme-dark");
            } else {
                root.getStyleClass().remove("theme-dark");
            }
            
            // 应用背景图片
            applyBackgroundImage();
        }
    }
    
    /**
     * 应用图片适应模式
     */
    private void applyImageFitMode(javafx.scene.image.ImageView imageView, javafx.scene.image.Image image) {
        String fitMode = settings.getBackgroundImageFitMode();
        double windowWidth = rootPane.getWidth();
        double windowHeight = rootPane.getHeight();
        double imageWidth = image.getWidth();
        double imageHeight = image.getHeight();
        
        switch (fitMode) {
            case "保持比例":
                // 保持比例，适应窗口大小
                imageView.setPreserveRatio(true);
                imageView.setFitWidth(windowWidth);
                imageView.setFitHeight(windowHeight);
                break;
                
            case "拉伸填充":
                // 拉伸填充整个窗口，不保持比例
                imageView.setPreserveRatio(false);
                imageView.setFitWidth(windowWidth);
                imageView.setFitHeight(windowHeight);
                break;
                
            case "原始大小":
                // 显示原始大小
                imageView.setPreserveRatio(true);
                imageView.setFitWidth(imageWidth);
                imageView.setFitHeight(imageHeight);
                break;
                
            case "适应宽度":
                // 适应宽度，保持比例
                imageView.setPreserveRatio(true);
                imageView.setFitWidth(windowWidth);
                break;
                
            case "适应高度":
                // 适应高度，保持比例
                imageView.setPreserveRatio(true);
                imageView.setFitHeight(windowHeight);
                break;
                
            default:
                // 默认保持比例
                imageView.setPreserveRatio(true);
                imageView.setFitWidth(windowWidth);
                imageView.setFitHeight(windowHeight);
                break;
        }
    }
    
    /**
     * 应用背景图片
     */
    private void applyBackgroundImage() {
        String imagePath = settings.getBackgroundImagePath();
        double opacity = settings.getBackgroundImageOpacity();
        
        if (rootPane != null && rootPane.getScene() != null) {
            if (imagePath != null && !imagePath.isEmpty()) {
                try {
                    File imageFile = new File(imagePath);
                    if (imageFile.exists()) {
                        // 为了实现透明度，我们需要使用一个背景层
                        if (backgroundPane == null) {
                            // 首次创建背景层
                            initBackgroundPane();
                        }
                        
                        // 加载图片
                        javafx.scene.image.Image image = new javafx.scene.image.Image(imageFile.toURI().toString());
                        if (backgroundImageView == null) {
                            backgroundImageView = new javafx.scene.image.ImageView();
                            // 设置图片适应方式
                            backgroundImageView.setSmooth(true);
                            backgroundImageView.setCache(true);
                        }
                        
                        backgroundImageView.setImage(image);
                        backgroundImageView.setOpacity(opacity);
                        
                        // 根据设置应用不同的图片适应模式
                        applyImageFitMode(backgroundImageView, image);
                        
                        // 监听窗口大小变化，动态调整图片大小（避免重复添加监听器）
                        if (!backgroundImageView.getProperties().containsKey("listenersAdded")) {
                            rootPane.widthProperty().addListener((obs, oldVal, newVal) -> {
                                if (backgroundImageView != null && backgroundImageView.getImage() != null) {
                                    applyImageFitMode(backgroundImageView, backgroundImageView.getImage());
                                }
                            });
                            
                            rootPane.heightProperty().addListener((obs, oldVal, newVal) -> {
                                if (backgroundImageView != null && backgroundImageView.getImage() != null) {
                                    applyImageFitMode(backgroundImageView, backgroundImageView.getImage());
                                }
                            });
                            
                            // 标记已添加监听器
                            backgroundImageView.getProperties().put("listenersAdded", true);
                        }
                        
                        // 如果背景层还未添加ImageView，则添加
                        if (!backgroundPane.getChildren().contains(backgroundImageView)) {
                            backgroundPane.getChildren().add(0, backgroundImageView);
                        }
                        
                        // 设置TableView和RootPane为透明背景，以便看到背景图片
                        if (tableView != null) {
                            tableView.setStyle("-fx-background-color: rgba(255, 255, 255, 0.85);");
                        }
                        if (rootPane != null) {
                            rootPane.setStyle("-fx-background-color: rgba(0, 0, 0, 0);");
                        }
                        
                    }
                } catch (Exception e) {
                    System.out.println("无法加载背景图片: " + e.getMessage());
                    e.printStackTrace();
                }
            } else {
                // 清除背景图片
                if (backgroundImageView != null && backgroundPane != null) {
                    backgroundPane.getChildren().remove(backgroundImageView);
                    // 清理监听器标记
                    backgroundImageView.getProperties().remove("listenersAdded");
                    backgroundImageView = null;
                }
                // 恢复默认背景
                if (tableView != null) {
                    tableView.setStyle("");
                }
                if (rootPane != null) {
                    rootPane.setStyle("");
                }
            }
        }
    }
    
    /**
     * 初始化背景层
     */
    private void initBackgroundPane() {
        if (rootPane != null && rootPane.getScene() != null) {
            javafx.scene.Parent currentRoot = rootPane.getScene().getRoot();
            
            // 如果当前根节点不是StackPane，则创建一个StackPane来包装
            if (!(currentRoot instanceof javafx.scene.layout.StackPane) || currentRoot == rootPane) {
                backgroundPane = new javafx.scene.layout.StackPane();
                // 使用rgba设置透明背景
                backgroundPane.setStyle("-fx-background-color: rgba(0, 0, 0, 0);");
                
                // 将BorderPane从Scene中移除
                javafx.scene.Scene scene = rootPane.getScene();
                
                // 将BorderPane添加到StackPane中
                backgroundPane.getChildren().add(rootPane);
                
                // 将StackPane设置为Scene的根节点
                scene.setRoot(backgroundPane);
            } else {
                backgroundPane = (javafx.scene.layout.StackPane) currentRoot;
            }
        }
    }
    
    /**
     * 显示导出设置对话框
     */
    private void showExportSettingsDialog() {
        Dialog<Boolean> dialog = new Dialog<>();
        dialog.setTitle("导出设置");
        dialog.setHeaderText("配置导出选项");
        
        ButtonType saveButtonType = new ButtonType("保存", ButtonBar.ButtonData.OK_DONE);
        dialog.getDialogPane().getButtonTypes().addAll(saveButtonType, ButtonType.CANCEL);
        
        // 创建选项卡面板
        TabPane tabPane = new TabPane();
        
        // Excel导出选项卡
        Tab excelTab = new Tab("Excel选项");
        excelTab.setClosable(false);
        javafx.scene.layout.GridPane excelGrid = new javafx.scene.layout.GridPane();
        excelGrid.setHgap(10);
        excelGrid.setVgap(10);
        excelGrid.setPadding(new javafx.geometry.Insets(20));
        
        int excelRow = 0;
        CheckBox excelHeaderCheck = new CheckBox();
        excelHeaderCheck.setSelected(exportSettings.isExcelIncludeHeader());
        excelGrid.add(new Label("包含标题行:"), 0, excelRow);
        excelGrid.add(excelHeaderCheck, 1, excelRow++);
        
        CheckBox excelAutoSizeCheck = new CheckBox();
        excelAutoSizeCheck.setSelected(exportSettings.isExcelAutoSizeColumns());
        excelGrid.add(new Label("自动调整列宽:"), 0, excelRow);
        excelGrid.add(excelAutoSizeCheck, 1, excelRow++);
        
        TextField excelSheetNameField = new TextField(exportSettings.getExcelSheetName());
        excelGrid.add(new Label("工作表名称:"), 0, excelRow);
        excelGrid.add(excelSheetNameField, 1, excelRow++);
        
        CheckBox excelHighlightCheck = new CheckBox();
        excelHighlightCheck.setSelected(exportSettings.isExcelApplyHighlight());
        excelGrid.add(new Label("应用高亮颜色:"), 0, excelRow);
        excelGrid.add(excelHighlightCheck, 1, excelRow++);
        
        excelTab.setContent(excelGrid);
        
        // PDF导出选项卡
        Tab pdfTab = new Tab("PDF选项");
        pdfTab.setClosable(false);
        javafx.scene.layout.GridPane pdfGrid = new javafx.scene.layout.GridPane();
        pdfGrid.setHgap(10);
        pdfGrid.setVgap(10);
        pdfGrid.setPadding(new javafx.geometry.Insets(20));
        
        int pdfRow = 0;
        TextField pdfTitleField = new TextField(exportSettings.getPdfTitle());
        pdfGrid.add(new Label("文档标题:"), 0, pdfRow);
        pdfGrid.add(pdfTitleField, 1, pdfRow++);
        
        ComboBox<String> pdfPageSizeCombo = new ComboBox<>();
        pdfPageSizeCombo.getItems().addAll("A4", "A3", "Letter", "Legal");
        pdfPageSizeCombo.setValue(exportSettings.getPdfPageSize());
        pdfGrid.add(new Label("页面大小:"), 0, pdfRow);
        pdfGrid.add(pdfPageSizeCombo, 1, pdfRow++);
        
        ComboBox<String> pdfOrientationCombo = new ComboBox<>();
        pdfOrientationCombo.getItems().addAll("Portrait", "Landscape");
        pdfOrientationCombo.setValue(exportSettings.getPdfOrientation());
        pdfGrid.add(new Label("页面方向:"), 0, pdfRow);
        pdfGrid.add(pdfOrientationCombo, 1, pdfRow++);
        
        Spinner<Integer> pdfFontSizeSpinner = new Spinner<>(6, 20, exportSettings.getPdfFontSize());
        pdfGrid.add(new Label("字体大小:"), 0, pdfRow);
        pdfGrid.add(pdfFontSizeSpinner, 1, pdfRow++);
        
        CheckBox pdfHeaderCheck = new CheckBox();
        pdfHeaderCheck.setSelected(exportSettings.isPdfIncludeHeader());
        pdfGrid.add(new Label("包含标题行:"), 0, pdfRow);
        pdfGrid.add(pdfHeaderCheck, 1, pdfRow++);
        
        CheckBox pdfHighlightCheck = new CheckBox();
        pdfHighlightCheck.setSelected(exportSettings.isPdfApplyHighlight());
        pdfGrid.add(new Label("应用高亮颜色:"), 0, pdfRow);
        pdfGrid.add(pdfHighlightCheck, 1, pdfRow++);
        
        pdfTab.setContent(pdfGrid);
        
        tabPane.getTabs().addAll(excelTab, pdfTab);
        dialog.getDialogPane().setContent(tabPane);
        
        dialog.setResultConverter(dialogButton -> {
            if (dialogButton == saveButtonType) {
                // 保存Excel设置
                exportSettings.setExcelIncludeHeader(excelHeaderCheck.isSelected());
                exportSettings.setExcelAutoSizeColumns(excelAutoSizeCheck.isSelected());
                exportSettings.setExcelSheetName(excelSheetNameField.getText());
                exportSettings.setExcelApplyHighlight(excelHighlightCheck.isSelected());
                
                // 保存PDF设置
                exportSettings.setPdfTitle(pdfTitleField.getText());
                exportSettings.setPdfPageSize(pdfPageSizeCombo.getValue());
                exportSettings.setPdfOrientation(pdfOrientationCombo.getValue());
                exportSettings.setPdfFontSize(pdfFontSizeSpinner.getValue());
                exportSettings.setPdfIncludeHeader(pdfHeaderCheck.isSelected());
                exportSettings.setPdfApplyHighlight(pdfHighlightCheck.isSelected());
                
                exportSettings.save();
                return true;
            }
            return false;
        });
        
        Optional<Boolean> result = dialog.showAndWait();
        if (result.isPresent() && result.get()) {
            updateStatus("导出设置已保存");
        }
    }
    
    /**
     * 带进度条的通用导出方法
     */
    private void exportWithProgress(File file, String format) {
        // 检测文件是否被锁定
        if (file.exists() && isFileLocked(file)) {
            showError("文件被锁定", "文件已被其他程序占用，请关闭后重试");
            return;
        }
        
        // 对于大文件显示进度条
        boolean showProgress = csvData.getRows() >= 10000;
        
        if (showProgress) {
            // 创建进度对话框
            Dialog<Void> progressDialog = new Dialog<>();
            progressDialog.setTitle("导出中");
            progressDialog.setHeaderText("正在导出数据，请稍候...");
            
            javafx.scene.control.ProgressIndicator progressIndicator = 
                new javafx.scene.control.ProgressIndicator();
            progressIndicator.setProgress(-1); // 不确定进度模式
            
            javafx.scene.layout.VBox vbox = new javafx.scene.layout.VBox(10);
            vbox.setPadding(new javafx.geometry.Insets(20));
            vbox.setAlignment(javafx.geometry.Pos.CENTER);
            vbox.getChildren().addAll(progressIndicator, new Label("正在导出 " + csvData.getRows() + " 行数据..."));
            
            progressDialog.getDialogPane().setContent(vbox);
            progressDialog.getDialogPane().getButtonTypes().clear();
            
            // 在后台线程执行导出
            javafx.concurrent.Task<Boolean> task = new javafx.concurrent.Task<>() {
                @Override
                protected Boolean call() throws Exception {
                    return performExport(file, format);
                }
            };
            
            task.setOnSucceeded(event -> {
                progressDialog.close();
                if (task.getValue()) {
                    showExportSuccessDialog(file, format);
                }
            });
            
            task.setOnFailed(event -> {
                progressDialog.close();
                Throwable exception = task.getException();
                showError("导出失败", "导出时发生错误：" + exception.getMessage());
            });
            
            Thread thread = new Thread(task);
            thread.setDaemon(true);
            thread.start();
            
            progressDialog.show();
        } else {
            // 小文件直接导出
            try {
                if (performExport(file, format)) {
                    showExportSuccessDialog(file, format);
                }
            } catch (Exception e) {
                showError("导出失败", "导出时发生错误：" + e.getMessage());
            }
        }
    }
    
    /**
     * 执行实际的导出操作
     */
    private boolean performExport(File file, String format) throws Exception {
        switch (format) {
            case "CSV":
                csvService.setLineEnding(settings.getLineEndingString());
                csvService.saveToFile(csvData, file);
                javafx.application.Platform.runLater(() -> updateStatus("已导出为CSV: " + file.getName()));
                return true;
            case "TXT":
                exportToTXT(file);
                return true;
            case "HTML":
                exportToHTML(file);
                return true;
            case "Excel":
                exportToExcel(file);
                return true;
            case "PDF":
                exportToPDF(file);
                return true;
            case "Markdown":
                exportToMarkdown(file);
                return true;
            default:
                throw new Exception("不支持的格式: " + format);
        }
    }
    
    /**
     * 检查文件是否被锁定
     */
    private boolean isFileLocked(File file) {
        try (java.io.RandomAccessFile raf = new java.io.RandomAccessFile(file, "rw");
             java.nio.channels.FileChannel channel = raf.getChannel()) {
            // 尝试获取文件锁
            java.nio.channels.FileLock lock = channel.tryLock();
            if (lock != null) {
                lock.release();
                return false;
            }
            return true;
        } catch (Exception e) {
            return true;
        }
    }
    
    /**
     * 显示导出成功对话框
     */
    private void showExportSuccessDialog(File file, String format) {
        Dialog<ButtonType> dialog = new Dialog<>();
        dialog.setTitle("导出成功");
        dialog.setHeaderText(format + " 文件导出成功！");
        dialog.setContentText("文件已保存到：\n" + file.getAbsolutePath());
        
        ButtonType openFileButton = new ButtonType("打开文件");
        ButtonType openFolderButton = new ButtonType("打开文件夹");
        ButtonType closeButton = new ButtonType("关闭", ButtonBar.ButtonData.CANCEL_CLOSE);
        
        dialog.getDialogPane().getButtonTypes().addAll(openFileButton, openFolderButton, closeButton);
        
        dialog.showAndWait().ifPresent(response -> {
            if (response == openFileButton) {
                // 打开文件
                try {
                    java.awt.Desktop.getDesktop().open(file);
                } catch (Exception e) {
                    showError("打开失败", "无法打开文件：" + e.getMessage());
                }
            } else if (response == openFolderButton) {
                // 打开文件夹
                try {
                    java.awt.Desktop.getDesktop().open(file.getParentFile());
                } catch (Exception e) {
                    showError("打开失败", "无法打开文件夹：" + e.getMessage());
                }
            }
        });
    }
    
    /**
     * 导出到TXT文件
     */
    private void exportToTXT(File file) throws Exception {
        try {
            java.io.FileWriter writer = new java.io.FileWriter(file, java.nio.charset.StandardCharsets.UTF_8);
            
            // 写入数据，使用制表符分隔
            for (int row = 0; row < csvData.getRows(); row++) {
                for (int col = 0; col < csvData.getColumns(); col++) {
                    String value = csvData.getCellValue(row, col);
                    writer.write(value != null ? value : "");
                    if (col < csvData.getColumns() - 1) {
                        writer.write("\t"); // 使用制表符分隔
                    }
                }
                writer.write(System.lineSeparator());
            }
            
            writer.close();
            javafx.application.Platform.runLater(() -> updateStatus("已导出为TXT: " + file.getName()));
        } catch (Exception e) {
            throw new Exception("导出TXT失败: " + e.getMessage(), e);
        }
    }
    
    /**
     * 导出到HTML文件
     */
    private void exportToHTML(File file) throws Exception {
        try {
            java.io.FileWriter writer = new java.io.FileWriter(file, java.nio.charset.StandardCharsets.UTF_8);
            
            // 写入HTML头部
            writer.write("<!DOCTYPE html>\n");
            writer.write("<html>\n");
            writer.write("<head>\n");
            writer.write("    <meta charset=\"UTF-8\">\n");
            writer.write("    <title>CSV导出</title>\n");
            writer.write("    <style>\n");
            writer.write("        table { border-collapse: collapse; width: 100%; }\n");
            writer.write("        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }\n");
            writer.write("        th { background-color: #4CAF50; color: white; }\n");
            writer.write("        tr:nth-child(even) { background-color: #f2f2f2; }\n");
            writer.write("    </style>\n");
            writer.write("</head>\n");
            writer.write("<body>\n");
            writer.write("    <h2>CSV数据导出</h2>\n");
            writer.write("    <table>\n");
            
            // 写入表头（如果需要）
            if (settings.isFirstRowAsHeader() && csvData.getRows() > 0) {
                writer.write("        <tr>\n");
                for (int col = 0; col < csvData.getColumns(); col++) {
                    String value = csvData.getCellValue(0, col);
                    writer.write("            <th>" + escapeHtml(value) + "</th>\n");
                }
                writer.write("        </tr>\n");
                
                // 从第二行开始写数据
                for (int row = 1; row < csvData.getRows(); row++) {
                    writer.write("        <tr>\n");
                    for (int col = 0; col < csvData.getColumns(); col++) {
                        String value = csvData.getCellValue(row, col);
                        String style = "";
                        
                        // 应用高亮颜色
                        javafx.scene.paint.Color color = highlightManager.getFinalHighlightColor(row, col);
                        if (color != null) {
                            style = String.format(" style=\"background-color: rgb(%d, %d, %d);\"",
                                (int)(color.getRed() * 255),
                                (int)(color.getGreen() * 255),
                                (int)(color.getBlue() * 255));
                        }
                        
                        writer.write("            <td" + style + ">" + escapeHtml(value) + "</td>\n");
                    }
                    writer.write("        </tr>\n");
                }
            } else {
                // 写入所有数据
                for (int row = 0; row < csvData.getRows(); row++) {
                    writer.write("        <tr>\n");
                    for (int col = 0; col < csvData.getColumns(); col++) {
                        String value = csvData.getCellValue(row, col);
                        String style = "";
                        
                        // 应用高亮颜色
                        javafx.scene.paint.Color color = highlightManager.getFinalHighlightColor(row, col);
                        if (color != null) {
                            style = String.format(" style=\"background-color: rgb(%d, %d, %d);\"",
                                (int)(color.getRed() * 255),
                                (int)(color.getGreen() * 255),
                                (int)(color.getBlue() * 255));
                        }
                        
                        writer.write("            <td" + style + ">" + escapeHtml(value) + "</td>\n");
                    }
                    writer.write("        </tr>\n");
                }
            }
            
            writer.write("    </table>\n");
            writer.write("</body>\n");
            writer.write("</html>\n");
            
            writer.close();
            javafx.application.Platform.runLater(() -> updateStatus("已导出为HTML: " + file.getName()));
        } catch (Exception e) {
            throw new Exception("导出HTML失败: " + e.getMessage(), e);
        }
    }
    
    /**
     * 导出到Markdown文件
     */
    private void exportToMarkdown(File file) throws Exception {
        try {
            java.io.FileWriter writer = new java.io.FileWriter(file, java.nio.charset.StandardCharsets.UTF_8);
            
            // 写入Markdown表格
            // 计算每列的最大宽度
            int[] columnWidths = new int[csvData.getColumns()];
            for (int col = 0; col < csvData.getColumns(); col++) {
                columnWidths[col] = ("列 " + (col + 1)).length();
                for (int row = 0; row < csvData.getRows(); row++) {
                    String value = csvData.getCellValue(row, col);
                    if (value != null) {
                        // 中文字符按2个字符宽度计算
                        int width = getStringDisplayWidth(value);
                        columnWidths[col] = Math.max(columnWidths[col], width);
                    }
                }
                // 最小宽度为3
                columnWidths[col] = Math.max(columnWidths[col], 3);
            }
            
            // 写入表头（如果需要）
            if (settings.isFirstRowAsHeader() && csvData.getRows() > 0) {
                writer.write("|");
                for (int col = 0; col < csvData.getColumns(); col++) {
                    String value = csvData.getCellValue(0, col);
                    String paddedValue = padString(value != null ? value : "", columnWidths[col]);
                    writer.write(" " + paddedValue + " |");
                }
                writer.write("\n");
                
                // 写入分隔行
                writer.write("|");
                for (int col = 0; col < csvData.getColumns(); col++) {
                    writer.write(" " + "-".repeat(columnWidths[col]) + " |");
                }
                writer.write("\n");
                
                // 写入数据（从第二行开始）
                for (int row = 1; row < csvData.getRows(); row++) {
                    writer.write("|");
                    for (int col = 0; col < csvData.getColumns(); col++) {
                        String value = csvData.getCellValue(row, col);
                        String paddedValue = padString(value != null ? value : "", columnWidths[col]);
                        writer.write(" " + paddedValue + " |");
                    }
                    writer.write("\n");
                }
            } else {
                // 写入默认表头
                writer.write("|");
                for (int col = 0; col < csvData.getColumns(); col++) {
                    String header = "列 " + (col + 1);
                    String paddedValue = padString(header, columnWidths[col]);
                    writer.write(" " + paddedValue + " |");
                }
                writer.write("\n");
                
                // 写入分隔行
                writer.write("|");
                for (int col = 0; col < csvData.getColumns(); col++) {
                    writer.write(" " + "-".repeat(columnWidths[col]) + " |");
                }
                writer.write("\n");
                
                // 写入所有数据
                for (int row = 0; row < csvData.getRows(); row++) {
                    writer.write("|");
                    for (int col = 0; col < csvData.getColumns(); col++) {
                        String value = csvData.getCellValue(row, col);
                        String paddedValue = padString(value != null ? value : "", columnWidths[col]);
                        writer.write(" " + paddedValue + " |");
                    }
                    writer.write("\n");
                }
            }
            
            writer.close();
            javafx.application.Platform.runLater(() -> updateStatus("已导出为Markdown: " + file.getName()));
        } catch (Exception e) {
            throw new Exception("导出Markdown失败: " + e.getMessage(), e);
        }
    }
    
    /**
     * 计算字符串显示宽度（中文字符算2个宽度）
     */
    private int getStringDisplayWidth(String str) {
        int width = 0;
        for (char c : str.toCharArray()) {
            if (c >= 0x4E00 && c <= 0x9FFF) { // 中文字符范围
                width += 2;
            } else {
                width += 1;
            }
        }
        return width;
    }
    
    /**
     * 填充字符串到指定宽度
     */
    private String padString(String str, int width) {
        int currentWidth = getStringDisplayWidth(str);
        if (currentWidth >= width) {
            return str;
        }
        return str + " ".repeat(width - currentWidth);
    }
    
    /**
     * HTML转义
     */
    private String escapeHtml(String text) {
        if (text == null || text.isEmpty()) {
            return "";
        }
        return text.replace("&", "&amp;")
                   .replace("<", "&lt;")
                   .replace(">", "&gt;")
                   .replace("\"", "&quot;")
                   .replace("'", "&#39;");
    }
    
    /**
     * 导出到Excel文件
     */
    private void exportToExcel(File file) throws Exception {
        hbnu.project.ergoucsveditior.service.ExcelExportService excelService = 
            new hbnu.project.ergoucsveditior.service.ExcelExportService();
        excelService.exportToExcel(csvData, file, exportSettings, highlightManager);
        javafx.application.Platform.runLater(() -> updateStatus("已导出为Excel: " + file.getName()));
    }
    
    /**
     * 导出到PDF文件
     */
    private void exportToPDF(File file) throws Exception {
        hbnu.project.ergoucsveditior.service.PDFExportService pdfService = 
            new hbnu.project.ergoucsveditior.service.PDFExportService();
        pdfService.exportToPDF(csvData, file, exportSettings, highlightManager);
        javafx.application.Platform.runLater(() -> updateStatus("已导出为PDF: " + file.getName()));
    }
    
    /**
     * 应用表格样式（边框、网格线颜色等）
     */
    private void applyTableStyles() {
        if (tableView != null) {
            String borderColor = settings.getTableBorderColor();
            String gridColor = settings.getTableGridColor();
            
            // 应用边框和网格线颜色
            String tableStyle = String.format(
                "-fx-border-color: %s; " +
                "-fx-border-width: 1px;",
                borderColor
            );
            
            // 如果有背景图片，保持透明背景
            if (settings.getBackgroundImagePath() != null && !settings.getBackgroundImagePath().isEmpty()) {
                tableStyle = "-fx-background-color: rgba(255, 255, 255, 0.85); " + tableStyle;
            }
            
            tableView.setStyle(tableStyle);
            
            // 动态添加CSS样式来设置表格单元格边框颜色
            String dynamicCSS = String.format(
                ".table-view .table-cell { -fx-border-color: %s; }",
                gridColor
            );
            
            // 移除之前的动态样式
            tableView.getStylesheets().removeIf(url -> url.contains("dynamic-table-cell"));
            
            // 添加新的动态样式
            try {
                java.io.File tempCSS = java.io.File.createTempFile("dynamic-table-cell", ".css");
                try (java.io.FileWriter writer = new java.io.FileWriter(tempCSS)) {
                    writer.write(dynamicCSS);
                }
                tableView.getStylesheets().add(tempCSS.toURI().toString());
            } catch (Exception e) {
                System.out.println("无法创建动态CSS样式: " + e.getMessage());
            }
            
            tableView.refresh();
        }
    }
    
    /**
     * 在初始化时应用主题
     */
    private void applyInitialTheme() {
        javafx.application.Platform.runLater(() -> {
            if (tableView != null && tableView.getScene() != null) {
                applyTheme();
                applyTableStyles();
            }
        });
    }
}
