package com.na.contract.utils;

import com.na.common.exceptions.NaBusinessException;
import com.na.common.result.enums.NaStatus;
import com.na.common.utils.LicenseValidator;
import com.na.common.utils.NaCommonUtil;
import com.na.common.utils.NaFileReadUtil;
import com.openhtmltopdf.pdfboxout.PdfRendererBuilder;
import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;

import org.springframework.stereotype.Component;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.nio.file.attribute.FileAttribute;
import java.nio.file.attribute.PosixFilePermission;
import java.nio.file.attribute.PosixFilePermissions;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Component
public class NaHtmlUtil {

    @Autowired
    private NaAutoContractConfig naAutoContractConfig;


    /**
     * 将原始 Map 的 key 转换为模板格式：{{key}}
     * 例如：name -> {{name}}
     */
//    public static Map<String, Object> convertMap(Map<String, Object> sourceMap) {
//        if (sourceMap != null && !sourceMap.isEmpty()) {
//            Map<String, Object> resultMap = new HashMap<>();
//            for (Map.Entry<String, Object> entry : sourceMap.entrySet()) {
//                String newKey = "{{" + entry.getKey() + "}}";
//                resultMap.put(newKey, entry.getValue());
//            }
//            return resultMap;
//        } else {
//            return Collections.emptyMap();
//        }
//    }

    /**
     * 将 HTML 字符串中的 {{key}} 占位符替换为 map 中对应的值。
     * 示例：若 HTML 为 "你好，{{name}}！"，map 中 "name" = "张三"，则输出 "你好，张三！"
     *
     * @param html HTML 模板字符串，包含 {{key}} 占位符
     * @param sourceMap  变量映射，key 为变量名（不带大括号），value 为替换值
     * @return 替换完成后的 HTML 字符串
     */
    private String replaceAll(String html, Map<String, Object> sourceMap) {
        // 判断输入是否合法
        if (html != null && !html.isEmpty() && sourceMap != null && !sourceMap.isEmpty()) {
            // 匹配 {{key}} 格式的模板变量
            Pattern pattern = Pattern.compile("\\{\\{(.*?)\\}\\}");
            Matcher matcher = pattern.matcher(html);
            StringBuffer result = new StringBuffer();

            // 遍历所有匹配到的模板变量
            while (matcher.find()) {
                String key = matcher.group(1).trim(); // 提取 key，例如 name
                // 从 map 中获取对应的值，如果不存在则使用空字符串
                String replacement = sourceMap.getOrDefault(key, "").toString();
                // 安全替换，避免 replacement 中包含 $、\ 等特殊字符
                matcher.appendReplacement(result, Matcher.quoteReplacement(replacement));
            }

            // 添加剩余文本
            matcher.appendTail(result);
            return result.toString();
        } else {
            // 输入为空，直接返回原始 HTML
            return html;
        }
    }


    public Boolean renderHtmlToFile(String htmlTempFilePath,
                                    Map<String, Object> sourceMap,
                                    String targetFilePath) {
        try {
            // 1. 检查配置对象
            if (naAutoContractConfig == null) {
                throw new NullPointerException("naAutoContractConfig is null");
            }

            // 2. 检查 License Key
            String key = naAutoContractConfig.getKey();
            if (StringUtils.isEmpty(key)) {
                throw new NaBusinessException(NaStatus.AUTHORIZATION_EXPIRED, null);
            }
            if (!LicenseValidator.isValidLicense(key)) {
                throw new NaBusinessException(NaStatus.AUTHORIZATION_EXPIRED, null);
            }

            // 3. 检查 HTML 模板路径
            if (StringUtils.isBlank(htmlTempFilePath) || !Files.exists(Paths.get(htmlTempFilePath))) {
                throw new FileNotFoundException("HTML 模板文件不存在: " + htmlTempFilePath);
            }

            // 4. 读取 HTML 模板
            String html = new String(Files.readAllBytes(Paths.get(htmlTempFilePath)), StandardCharsets.UTF_8);

            // 5. 替换模板变量
            if (sourceMap == null) {
                return false;
            }
            html = replaceAll(html, sourceMap);

            // 6. 检查字体文件
            String fontPathRegular = NaFileReadUtil.getFileAbsolutePath("fonts/NotoSerifSC-Regular.ttf");
            String fontPathBold = NaFileReadUtil.getFileAbsolutePath("fonts/NotoSerifSC-Bold.ttf");
            if (StringUtils.isBlank(fontPathRegular) || !new File(fontPathRegular).exists()) {
                throw new FileNotFoundException("找不到 Regular 字体文件: " + fontPathRegular);
            }
            if (StringUtils.isBlank(fontPathBold) || !new File(fontPathBold).exists()) {
                throw new FileNotFoundException("找不到 Bold 字体文件: " + fontPathBold);
            }

            // 7. 创建输出文件
            if (!NaCommonUtil.isWindows()) {
                Path path = Paths.get(targetFilePath);
                Files.deleteIfExists(path); // 删除旧文件，避免 FileAlreadyExistsException

                Set<PosixFilePermission> perms = PosixFilePermissions.fromString("rwxr-xr-x");
                FileAttribute<Set<PosixFilePermission>> attr = PosixFilePermissions.asFileAttribute(perms);
                path = Files.createFile(path, attr);

                try (OutputStream os = Files.newOutputStream(path)) {
                    PdfRendererBuilder builder = new PdfRendererBuilder();
                    builder.useFastMode();
                    builder.useFont(new File(fontPathRegular), "Noto Serif SC", 400, PdfRendererBuilder.FontStyle.NORMAL, true);
                    builder.useFont(new File(fontPathBold), "Noto Serif SC", 700, PdfRendererBuilder.FontStyle.NORMAL, true);
                    builder.withHtmlContent(html, "https://");
                    builder.toStream(os);
                    System.out.println("开始执行 builder.run()");
                    builder.run();
                    System.out.println("结束执行 builder.run()");
                }
            } else {
                // Windows 系统直接用普通方式创建
                try (OutputStream os = new FileOutputStream(targetFilePath)) {
                    PdfRendererBuilder builder = new PdfRendererBuilder();
                    builder.useFastMode();
                    builder.useFont(new File(fontPathRegular), "Noto Serif SC", 400, PdfRendererBuilder.FontStyle.NORMAL, true);
                    builder.useFont(new File(fontPathBold), "Noto Serif SC", 700, PdfRendererBuilder.FontStyle.NORMAL, true);
                    builder.withHtmlContent(html, "https://");
                    builder.toStream(os);
                    builder.run();
                }
            }

            return true;
        }catch (Exception e) {
            System.out.println(e);
            return false;
        }
    }

//    public static void main(String[] args) {
//        String htmlTempFilePath = "D:\\hetong.html";
//        Map<String, Object> sourceMap = new HashMap<>();
//        String targetFilePath = "D:\\hetong.pdf";
//        sourceMap.put("userIdB","1111");
//        sourceMap.put("nameB","2222");
////        sourceMap = convertMap(sourceMap);
//        renderHtmlToFile(htmlTempFilePath,sourceMap,targetFilePath);
//    }
}
