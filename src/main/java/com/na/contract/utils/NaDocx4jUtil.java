package com.na.contract.utils;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.springframework.core.io.ClassPathResource;

import com.deepoove.poi.XWPFTemplate;
import org.springframework.core.io.ClassPathResource;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Map;

/**
 * @author pg
 * @date 2026年03月13日
 */
public class NaDocx4jUtil {

    public static void fillTempWord(String templateName,
                              Map<String, Object> data,
                              String filePath) throws Exception {

        // 创建目录
        File file = new File(filePath);
        File parent = file.getParentFile();
        if (parent != null && !parent.exists()) {
            boolean success = parent.mkdirs();
            if (!success) {
                throw new RuntimeException("创建目录失败：" + parent.getAbsolutePath());
            }
        }

        // 读取模板
        ClassPathResource resource = new ClassPathResource(templateName);

        try (InputStream inputStream = resource.getInputStream();
             XWPFTemplate template = XWPFTemplate.compile(inputStream).render(data);
             FileOutputStream out = new FileOutputStream(file)) {

            template.write(out);
        }
    }

    public static void convert(String docxPath, String pdfPath) {

        if (docxPath == null || docxPath.trim().isEmpty()) {
            throw new IllegalArgumentException("Word文件路径不能为空");
        }

        if (pdfPath == null || pdfPath.trim().isEmpty()) {
            throw new IllegalArgumentException("PDF输出路径不能为空");
        }

        File docxFile = new File(docxPath);

        // 判断Word文件是否存在
        if (!docxFile.exists()) {
            throw new RuntimeException("Word文件不存在: " + docxPath);
        }

        // 判断是否是文件
        if (!docxFile.isFile()) {
            throw new RuntimeException("Word路径不是文件: " + docxPath);
        }

        // 判断文件类型
        if (!docxPath.toLowerCase().endsWith(".docx")) {
            throw new RuntimeException("仅支持docx文件: " + docxPath);
        }

        File pdfFile = new File(pdfPath);

        // 判断输出目录
        File parentDir = pdfFile.getParentFile();
        if (parentDir != null && !parentDir.exists()) {
            parentDir.mkdirs(); // 自动创建目录
        }

        try {

            WordprocessingMLPackage wordMLPackage =
                    WordprocessingMLPackage.load(docxFile);

            try (FileOutputStream os = new FileOutputStream(pdfFile)) {
                Docx4J.toPDF(wordMLPackage, os);
            }

        } catch (Exception e) {
            throw new RuntimeException("Word转PDF失败: " + e.getMessage(), e);
        }
    }

    public static boolean deleteFile(String filePath) {

        if (filePath == null || filePath.trim().isEmpty()) {
            return false;
        }

        File file = new File(filePath);

        // 文件不存在
        if (!file.exists()) {
            return false;
        }

        // 如果是文件直接删除
        if (file.isFile()) {
            return file.delete();
        }

        // 如果是目录递归删除
        return deleteDirectory(file);
    }

    private static boolean deleteDirectory(File dir) {

        File[] files = dir.listFiles();

        if (files != null) {
            for (File file : files) {

                if (file.isFile()) {
                    file.delete();
                } else {
                    deleteDirectory(file);
                }

            }
        }

        return dir.delete();
    }
}
