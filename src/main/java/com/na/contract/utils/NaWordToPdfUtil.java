package com.na.contract.utils;

import com.na.common.utils.NaCommonUtil;
import com.na.common.utils.NaIDUtil;
import com.na.contract.dto.NaResponse;
import com.na.contract.dto.NaWordToPdfDTO;
import org.apache.commons.lang3.StringUtils;

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

    /** LibreOffice 可执行路径 */
    private static final String WIN_LIBREOFFICE_PATH = "C:\\Program Files\\LibreOffice\\program\\soffice.exe";
    private static final String LINUX_LIBREOFFICE_PATH = "/opt/libreoffice25.8/program/soffice";

    /** 临时目录根路径 */
    private static final String WIN_TEMP_DIR = "C:\\na\\contract\\temp";
    private static final String LINUX_TEMP_DIR = "/na/contract/temp";

    /**
     * 将 Word 文件（Base64）转换为 PDF（Base64）
     *
     * @param dto 输入数据对象，包含 Base64 内容
     * @return NaResponse 返回 Base64 PDF
     */
    public static NaResponse wordToPdf(NaWordToPdfDTO dto) throws Exception {
        // 1. 拼写临时 Word 路径
        String wordPath = buildTempPath(dto, NaIDUtil.getRandomNo(7), ".docx");

        // 2. 写入 Base64 到 Word 文件
        writeBase64ToFile(dto.getBase64(), wordPath);

        // 3. 拼写 PDF 输出路径
        String pdfPath = wordPath.replace(".docx", ".pdf");

        // 4. 确保输出目录存在
        ensureParentDirectoryExists(pdfPath);

        // 5. 调用 LibreOffice 转换
        executeLibreOfficeConvert(wordPath, pdfPath);

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

    // ========================= 辅助方法 =========================

    /** 拼接临时文件路径 */
    private static String buildTempPath(NaWordToPdfDTO dto, String fileName, String extension) {
        String folder = dto.getAbsPath();
        if(StringUtils.isNotEmpty(folder)){
            return String.format("%s/%s%s", folder, fileName, extension);
        }
        folder = "tmp";
        return String.format("%s/%s/%s%s", NaCommonUtil.isWindows() ? WIN_TEMP_DIR : LINUX_TEMP_DIR, folder, fileName, extension);
    }

    /** 写 Base64 到文件 */
    private static void writeBase64ToFile(String base64, String filePath) throws IOException {
        byte[] bytes = Base64.getDecoder().decode(base64);
        Path path = Paths.get(filePath);
        ensureParentDirectoryExists(filePath);
        Files.write(path, bytes);
    }

    /** 确保父目录存在 */
    private static void ensureParentDirectoryExists(String filePath) {
        File parent = new File(filePath).getParentFile();
        if (!parent.exists()) {
            boolean created = parent.mkdirs();
            if (!created) {
                System.err.println("创建目录失败: " + parent.getAbsolutePath());
            }
        }
    }

    /** 调用 LibreOffice 将 Word 转 PDF */
    private static void executeLibreOfficeConvert(String wordPath, String pdfPath) throws IOException, InterruptedException {
        File pdfFile = new File(pdfPath);
        File parent = pdfFile.getParentFile();

        ProcessBuilder pb = new ProcessBuilder(
                NaCommonUtil.isWindows() ? WIN_LIBREOFFICE_PATH : LINUX_LIBREOFFICE_PATH,
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

    /** PDF → Base64 */
    private static String pdfToBase64(String pdfPath) throws IOException {
        Path path = Paths.get(pdfPath);
        byte[] bytes = Files.readAllBytes(path);
        return Base64.getEncoder().encodeToString(bytes);
    }

    /** 删除文件（存在则删除） */
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
