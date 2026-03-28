package com.na.contract.utils;

import com.na.common.utils.NaCommonUtil;
import com.na.common.utils.NaIDUtil;
import com.na.contract.dto.NaResponse;
import com.na.contract.dto.NaWordToPdfDTO;
import org.apache.commons.lang3.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Base64;

/**
 * Word 转 PDF 工具类（使用 LibreOffice）
 * 支持 Base64 输入/输出
 */
public class NaWordToPdfUtil {

    public static NaResponse wordToPdf(NaWordToPdfDTO dto) throws Exception{
        if(dto == null || !dto.isPdfFile() || StringUtils.isEmpty(dto.getBase64())){
            return NaResponse.builder()
                    .code(-1)
                    .msg("参数错误")
                    .build();
        }
        String pdfPath = dto.getAbsPath();
        // 确保输出目录存在
        ensureParentDirectoryExists(pdfPath);

        // 拼写 word 临时输出路径
        String wordPath = pdfPath.replace(".pdf", ".docx");

        // 写入 Base64 到 Word 临时文件
        writeBase64ToFile(dto.getBase64(), wordPath);

        // 调用 LibreOffice 转换
        executeLibreOfficeConvert(wordPath, pdfPath,dto.getLibreoffice());

        // 将 PDF 转 Base64
        String pdfBase64 = fileToBase64(pdfPath);

        // 7. 删除临时文件
        deleteFile(wordPath);
        deleteFile(pdfPath);

        return NaResponse.builder()
                .code(0)
                .msg("转换成功")
                .data(pdfBase64)
                .build();
    }


    /** 写 Base64 到文件 */
    public static void writeBase64ToFile(String base64, String filePath) throws IOException {
        byte[] bytes = Base64.getDecoder().decode(base64);
        Path path = Paths.get(filePath);
        ensureParentDirectoryExists(filePath);
        Files.write(path, bytes);
    }

    /** 确保父目录存在 */
    public static void ensureParentDirectoryExists(String filePath) {
        File parent = new File(filePath).getParentFile();
        if (!parent.exists()) {
            boolean created = parent.mkdirs();
            if (!created) {
                System.err.println("创建目录失败: " + parent.getAbsolutePath());
            }
        }
    }

    /** 调用 LibreOffice 将 Word 转 PDF */
    public static void executeLibreOfficeConvert(String wordPath, String pdfPath,String libreoffice) throws IOException, InterruptedException {
        File pdfFile = new File(pdfPath);
        File parent = pdfFile.getParentFile();

        ProcessBuilder pb = new ProcessBuilder(
                libreoffice,
                "--headless",
                "--convert-to", "pdf",
                "--outdir", parent.getAbsolutePath(),
                wordPath
        );

        // 打印执行命令
        System.out.println("执行命令: " + String.join(" ", pb.command()));

        pb.redirectErrorStream(true); // 将错误输出合并到标准输出
        Process process = pb.start();

        // 打印 LibreOffice 输出
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()))) {
            reader.lines().forEach(System.out::println);
        }

        int exitCode = process.waitFor();
        if (exitCode != 0) {
            throw new RuntimeException("LibreOffice 转换失败，退出码：" + exitCode);
        }
    }

    /** 文件 → Base64 */
    public static String fileToBase64(String pdfPath) throws IOException {
        Path path = Paths.get(pdfPath);
        byte[] bytes = Files.readAllBytes(path);
        return Base64.getEncoder().encodeToString(bytes);
    }

    /** 删除文件（存在则删除） */
    public static void deleteFile(String path) {
        File file = new File(path);
        if (file.exists()) {
            boolean success = file.delete();
            if (!success) {
                System.err.println("删除文件失败: " + path);
            }
        }
    }

    /**
     * base64 转 InputStream（智能去掉前缀，无前缀则不解）
     */
    public static InputStream base64ToInputStream(String base64) {
        if (base64 == null) {
            return null;
        }
        // 只有包含逗号时，才去掉前缀 data:...base64,
        if (base64.contains(",")) {
            base64 = base64.split(",")[1];
        }
        // 解码
        byte[] bytes = Base64.getDecoder().decode(base64);
        // 返回流
        return new ByteArrayInputStream(bytes);
    }
}
