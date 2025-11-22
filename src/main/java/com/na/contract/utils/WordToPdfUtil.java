package com.na.contract.utils;

import com.na.contract.dto.NaResponse;
import com.na.contract.dto.NaWordToPdfDTO;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Base64;

/**
 * @author pg
 */
public class WordToPdfUtil {
    public static NaResponse wordToPdf(NaWordToPdfDTO dto) throws Exception {
        // 1. 拼写 Word 路径
        String wordPath = String.format("/na/contract/temp/%s/%s.docx", "1", "2");
//        String wordPath = String.format("D:\\%s\\%s.docx", "1", "2");

        // 2. 将 Base64 写入 Word 文件
        writeBase64ToFile(dto.getBase64(), wordPath);

        // 3. 拼写 PDF 输出路径
        String pdfPath = wordPath.replace(".docx", ".pdf");

        // 4. 确保输出目录存在
        File pdfFile = new File(pdfPath);
        File parent = pdfFile.getParentFile();
        if (!parent.exists()) {
            parent.mkdirs();
        }

        // 5. 调用 LibreOffice 转换
        ProcessBuilder pb = new ProcessBuilder(
                "/opt/libreoffice25.8/program/soffice",
                "--headless",
                "--convert-to", "pdf",
                "--outdir", parent.getAbsolutePath(),
                wordPath
        );
        // 打印执行的命令
        System.out.println(String.join(" ", pb.command()));
        pb.redirectErrorStream(true);
        Process process = pb.start();
        int exitCode = process.waitFor();
        if (exitCode != 0) {
            throw new RuntimeException("LibreOffice 转换失败，退出码：" + exitCode);
        }

        // 6. 将 PDF 转 Base64
        String pdfBase64 = pdfToBase64(pdfPath);

        // 7. 删除临时文件
        deleteFile(wordPath);
        deleteFile(pdfPath);

        // 8. 返回结果
        NaResponse response = new NaResponse();
        response.setCode(0);
        response.setMsg("转换成功");
        response.setData(pdfBase64);
        return response;
    }

    // 写 Base64 到文件
    public static void writeBase64ToFile(String base64, String filePath) throws Exception {
        byte[] bytes = Base64.getDecoder().decode(base64);
        File file = new File(filePath);
        File parent = file.getParentFile();
        if (!parent.exists()) {
            parent.mkdirs();
        }
        try (FileOutputStream fos = new FileOutputStream(file)) {
            fos.write(bytes);
            fos.flush();
        }
    }

    // PDF → Base64 (Java 8 兼容)
    private static String pdfToBase64(String pdfPath) throws Exception {
        File file = new File(pdfPath);
        try (FileInputStream fis = new FileInputStream(file);
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            byte[] buffer = new byte[8192];
            int bytesRead;
            while ((bytesRead = fis.read(buffer)) != -1) {
                baos.write(buffer, 0, bytesRead);
            }
            return Base64.getEncoder().encodeToString(baos.toByteArray());
        }
    }

    // 删除文件
    private static void deleteFile(String path) {
        File file = new File(path);
        if (file.exists()) {
            boolean success = file.delete();
            if (!success) {
                System.err.println("删除文件失败: " + path);
            }
        }
    }
}
