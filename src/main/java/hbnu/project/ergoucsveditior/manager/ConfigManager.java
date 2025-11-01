package hbnu.project.ergoucsveditior.manager;

import java.io.*;
import java.nio.file.*;

/**
 * 配置文件管理器
 * 统一管理应用程序的配置文件路径和读写
 */
public class ConfigManager {
    // 应用程序配置目录名称
    private static final String APP_CONFIG_DIR = "ErgouCSVEditor";
    
    // 配置文件名称常量
    public static final String SETTINGS_FILE = "csv_editor_settings.properties";
    public static final String KEYBINDINGS_FILE = "csv_editor_keybindings.properties";
    public static final String TOOLBAR_FILE = "csv_editor_toolbar.properties";
    public static final String AUTOMARK_SETTINGS_FILE = "csv_editor_automark_settings.properties";
    public static final String AUTOMARK_RULES_FILE = "csv_editor_automark_rules.dat";
    
    // 配置目录路径
    private static Path configDirPath;
    
    static {
        initConfigDirectory();
    }
    
    /**
     * 初始化配置目录
     */
    private static void initConfigDirectory() {
        try {
            // 获取用户主目录
            String userHome = System.getProperty("user.home");
            
            // 根据操作系统确定配置目录位置
            String os = System.getProperty("os.name").toLowerCase();
            if (os.contains("win")) {
                // Windows: %APPDATA%/ErgouCSVEditor
                String appData = System.getenv("APPDATA");
                if (appData != null) {
                    configDirPath = Paths.get(appData, APP_CONFIG_DIR);
                } else {
                    configDirPath = Paths.get(userHome, APP_CONFIG_DIR);
                }
            } else if (os.contains("mac")) {
                // macOS: ~/Library/Application Support/ErgouCSVEditor
                configDirPath = Paths.get(userHome, "Library", "Application Support", APP_CONFIG_DIR);
            } else {
                // Linux/Unix: ~/.config/ErgouCSVEditor
                String xdgConfigHome = System.getenv("XDG_CONFIG_HOME");
                if (xdgConfigHome != null) {
                    configDirPath = Paths.get(xdgConfigHome, APP_CONFIG_DIR);
                } else {
                    configDirPath = Paths.get(userHome, ".config", APP_CONFIG_DIR);
                }
            }
            
            // 创建配置目录（如果不存在）
            if (!Files.exists(configDirPath)) {
                Files.createDirectories(configDirPath);
                System.out.println("配置目录已创建: " + configDirPath);
            }
            
            // 创建rules子目录
            Path rulesDir = configDirPath.resolve("rules");
            if (!Files.exists(rulesDir)) {
                Files.createDirectories(rulesDir);
            }
            
        } catch (IOException e) {
            System.err.println("无法创建配置目录: " + e.getMessage());
            // 降级到当前目录
            configDirPath = Paths.get(".");
        }
    }
    
    /**
     * 获取配置目录路径
     */
    public static Path getConfigDir() {
        return configDirPath;
    }
    
    /**
     * 获取配置文件的完整路径
     * @param fileName 配置文件名
     * @return 配置文件路径
     */
    public static File getConfigFile(String fileName) {
        return configDirPath.resolve(fileName).toFile();
    }
    
    /**
     * 获取rules目录下的文件路径
     * @param fileName 文件名
     * @return 文件路径
     */
    public static File getRulesFile(String fileName) {
        return configDirPath.resolve("rules").resolve(fileName).toFile();
    }
    
    /**
     * 检查配置文件是否存在
     * @param fileName 配置文件名
     * @return 如果存在返回true
     */
    public static boolean configFileExists(String fileName) {
        return getConfigFile(fileName).exists();
    }
    
    /**
     * 从resources目录复制默认配置文件到用户配置目录
     * @param fileName 配置文件名
     * @return 是否成功复制
     */
    public static boolean copyDefaultConfig(String fileName) {
        try {
            File targetFile = getConfigFile(fileName);
            
            // 如果目标文件已存在，不覆盖
            if (targetFile.exists()) {
                return true;
            }
            
            // 尝试从resources读取默认配置
            InputStream inputStream = ConfigManager.class.getClassLoader()
                    .getResourceAsStream(fileName);
            
            if (inputStream == null) {
                // 尝试从resources子目录读取
                inputStream = ConfigManager.class.getClassLoader()
                        .getResourceAsStream("setting_properties/" + fileName);
            }
            
            if (inputStream == null) {
                inputStream = ConfigManager.class.getClassLoader()
                        .getResourceAsStream("toolbar_properties/" + fileName);
            }
            
            if (inputStream != null) {
                // 复制文件
                Files.copy(inputStream, targetFile.toPath());
                inputStream.close();
                System.out.println("已从resources复制默认配置: " + fileName);
                return true;
            } else {
                System.out.println("resources中未找到默认配置文件: " + fileName);
                return false;
            }
        } catch (IOException e) {
            System.err.println("复制默认配置文件失败: " + e.getMessage());
            return false;
        }
    }
    
    /**
     * 获取配置文件的输入流
     * 优先从用户配置目录读取，如果不存在则从resources读取
     * @param fileName 配置文件名
     * @return 输入流，如果文件不存在返回null
     */
    public static InputStream getConfigInputStream(String fileName) throws IOException {
        File configFile = getConfigFile(fileName);
        
        // 如果用户配置文件存在，直接读取
        if (configFile.exists()) {
            return new FileInputStream(configFile);
        }
        
        // 否则尝试从resources读取
        InputStream inputStream = ConfigManager.class.getClassLoader()
                .getResourceAsStream(fileName);
        
        if (inputStream == null) {
            inputStream = ConfigManager.class.getClassLoader()
                    .getResourceAsStream("setting_properties/" + fileName);
        }
        
        if (inputStream == null) {
            inputStream = ConfigManager.class.getClassLoader()
                    .getResourceAsStream("toolbar_properties/" + fileName);
        }
        
        // 如果从resources读取到了，复制一份到用户配置目录
        if (inputStream != null) {
            copyDefaultConfig(fileName);
        }
        
        return inputStream;
    }
    
    /**
     * 获取配置文件的输出流
     * @param fileName 配置文件名
     * @return 输出流
     */
    public static OutputStream getConfigOutputStream(String fileName) throws IOException {
        File configFile = getConfigFile(fileName);
        
        // 确保父目录存在
        File parentDir = configFile.getParentFile();
        if (parentDir != null && !parentDir.exists()) {
            parentDir.mkdirs();
        }
        
        return new FileOutputStream(configFile);
    }
    
    /**
     * 获取配置目录路径的字符串表示
     */
    public static String getConfigDirPath() {
        return configDirPath.toString();
    }
    
    /**
     * 打开配置目录（用系统文件管理器打开）
     */
    public static void openConfigDirectory() {
        try {
            java.awt.Desktop.getDesktop().open(configDirPath.toFile());
        } catch (IOException e) {
            System.err.println("无法打开配置目录: " + e.getMessage());
        }
    }
}

