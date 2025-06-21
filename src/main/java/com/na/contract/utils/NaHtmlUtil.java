package com.na.contract.utils;

import com.na.common.exceptions.NaBusinessException;
import com.na.common.result.enums.NaStatus;
import com.na.common.utils.LicenseValidator;
import com.na.common.utils.NaFileReadUtil;
import com.openhtmltopdf.pdfboxout.PdfRendererBuilder;
import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;

import org.springframework.stereotype.Component;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.Map;
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
            String key = naAutoContractConfig != null ? naAutoContractConfig.getKey() : null;
            if (StringUtils.isEmpty(key)) {
                throw new NaBusinessException(NaStatus.AUTHORIZATION_EXPIRED,null);
            }

           if(!LicenseValidator.isValidLicense(key)) {
               throw new NaBusinessException(NaStatus.AUTHORIZATION_EXPIRED,null);
           }

            // 读取 HTML 文件
            String html = new String(Files.readAllBytes(Paths.get(htmlTempFilePath)), StandardCharsets.UTF_8);

            // 替换模板变量
            html = replaceAll(html, sourceMap);


            // 创建输出文件流
            try (OutputStream os = new FileOutputStream(targetFilePath)) {
                PdfRendererBuilder builder = new PdfRendererBuilder();
                builder.useFastMode();

                // 加载字体，名字必须和 HTML 中设置的一致
                File fontFileRegular = new File(NaFileReadUtil.getFileAbsolutePath("fonts/NotoSerifSC-Regular.ttf"));
                File fontFileBold = new File(NaFileReadUtil.getFileAbsolutePath("fonts/NotoSerifSC-Bold.ttf"));

                builder.useFont(fontFileRegular, "Noto Serif SC", 400, PdfRendererBuilder.FontStyle.NORMAL, true);
                builder.useFont(fontFileBold, "Noto Serif SC", 700, PdfRendererBuilder.FontStyle.NORMAL, true);

                // 必须添加这行：让页眉页脚也能使用注册字体
//                builder.usePdfAConformance(PdfRendererBuilder.PdfAConformance.PDFA_1_A); // ⚠️或不加，用下面方式强制嵌入

                // 确保嵌入字体到页眉页脚
//                builder.useDefaultPageSize(210, 297, PdfRendererBuilder.PageSizeUnits.MM); // A4

                System.out.println("Regular font exists: " + fontFileRegular.exists());
                System.out.println("Bold font exists: " + fontFileBold.exists());


                // 设置 baseUri，支持网络图片加载
                builder.withHtmlContent(html, "https://");

                builder.toStream(os);
                builder.run();
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
