package com.na.contract.utils;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.util.Units;

import java.io.*;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Word 模板填充工具（兼容 JDK8）
 *
 * 功能：
 *  - 文本占位符替换 ${key}
 *  - 图片占位符替换 ${imgKey}（imageParams 中传 InputStream）
 *  - 表格循环：在表格任意单元格使用 ${table:tableName} 标识（表格中任意一行），
 *    程序会把该行的下一行作为模板行（template row），并用 tableParams.get(tableName) 数据插入多行。
 *
 * 模板约定：
 *  - 文本占位符： ${key}
 *  - 图片占位符： ${imgKey} （imageParams 储存 InputStream）
 *  - 表格循环： 在某单元格中包含 ${table:orderTable}，则认为该表格需要循环，
 *    该单元格所在行为“标识行”，标识行的下一行为“模板行”，模板行内使用 ${colName} 占位。
 *
 * 注意：
 *  - 如果占位符被拆分到多个 runs（Word 常见），代码会合并后替换并重写段落（单 run 写回）。
 *  - imageParams 的 InputStream 由调用方负责关闭（方法内部不会关闭传入的 InputStream）。
 */
public class WordBase64ReplaceUtil {

    // 表格占位符正则，匹配 ${table:xxx}，xxx 捕获为 group(1)
    private static final Pattern TABLE_PLACEHOLDER_PATTERN = Pattern.compile("\\$\\{table:([^}]+)}");
    // 文本/图片占位符正则（用于段落级别按 key 替换时可选）
    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("\\$\\{([^}]+)}");

    /**
     * 将 Base64 的 Word 文档进行占位符替换，返回替换后的 Base64 Word
     *
     * @param templateBase64 Word 模板 Base64，不可为空
     * @param textParams     文本占位符 Map<key, value>，可为 null
     * @param imageParams    图片占位符 Map<imgKey, InputStream>，可为 null（InputStream 由调用方管理关闭）
     * @param tableParams    表格循环 Map<tableName, List<Map<列名, 值>>>，可为 null
     * @return 替换后的 Word 的 Base64 字符串
     */
    public static String fillTemplate(
            String templateBase64,
            Map<String, Object> textParams,
            Map<String, InputStream> imageParams,
            Map<String, List<Map<String, Object>>> tableParams
    ) throws Exception {

        if (templateBase64 == null || templateBase64.trim().isEmpty()) {
            throw new IllegalArgumentException("模板 Base64 不能为空");
        }

        // null -> empty maps 保证后续无需判空
        textParams = textParams == null ? Collections.emptyMap() : textParams;
        imageParams = imageParams == null ? Collections.emptyMap() : imageParams;
        tableParams = tableParams == null ? Collections.emptyMap() : tableParams;

        byte[] wordBytes = Base64.getDecoder().decode(templateBase64);
        try (XWPFDocument doc = new XWPFDocument(new ByteArrayInputStream(wordBytes))) {

            // 1) 段落级文本/图片替换（文档自由段落）
            for (XWPFParagraph p : doc.getParagraphs()) {
                replaceParagraphTextAndImages(p, textParams, imageParams);
            }

            // 2) 表格处理（包括表格中的段落）
            //    注意：表格里既可能包含普通占位符，也可能包含循环表格占位符
            for (XWPFTable table : doc.getTables()) {
                processTable(table, textParams, imageParams, tableParams);
            }

            // 3) 输出 Base64
            try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
                doc.write(baos);
                return Base64.getEncoder().encodeToString(baos.toByteArray());
            }
        }
    }

    // ---------------------------
    // 段落替换（文本 + 图片）
    // ---------------------------
    private static void replaceParagraphTextAndImages(XWPFParagraph paragraph,
                                                      Map<String, Object> textParams,
                                                      Map<String, InputStream> imageParams) throws Exception {
        if (paragraph == null) {return;}

        List<XWPFRun> runs = paragraph.getRuns();
        if (runs == null || runs.isEmpty()) {return;}

        // 合并所有 run 文本，以便支持跨 run 的占位符
        StringBuilder sb = new StringBuilder();
        for (XWPFRun run : runs) {
            String t = run.getText(0);
            sb.append(t == null ? "" : t);
        }
        String merged = sb.toString();

        // 替换文本占位符 ${key} -> value（value 为 null 则替换为空字符串）
        String replaced = merged;
        for (Map.Entry<String, Object> entry : textParams.entrySet()) {
            String key = "${" + entry.getKey() + "}";
            String value = entry.getValue() == null ? "" : entry.getValue().toString();
            if (replaced.contains(key)) {
                replaced = replaced.replace(key, value);
            }
        }

        // 替换图片占位符（如果存在），注意：此处只支持替换为单张图片并写入对应位置
        // 图片占位符必须以 ${imgKey} 形式存在
        boolean hadImage = false;
        for (Map.Entry<String, InputStream> imgEntry : imageParams.entrySet()) {
            String key = "${" + imgEntry.getKey() + "}";
            if (replaced.contains(key)) {
                // 将该占位符删除（用空字符串），然后在该段落末尾插入图片（或你可以改为在占位处插入）
                replaced = replaced.replace(key, "");
                hadImage = true;

                // 清空原 runs 并写回文本（先清空，再写）
                clearRunsAndSetText(paragraph, replaced);

                // 在段落末尾插入图片（宽高为 150x150，可按需调整或做参数）
                XWPFRun imageRun = paragraph.createRun();
                try (InputStream in = imgEntry.getValue()) {
                    if (in != null) {
                        // 默认当 PNG 处理；如果你需要根据图片类型动态设置，需传入额外信息
                        imageRun.addPicture(in, XWPFDocument.PICTURE_TYPE_PNG, imgEntry.getKey(),
                                Units.toEMU(150), Units.toEMU(150));
                    }
                }
                // 可能存在多个图片占位符，继续循环处理
            }
        }

        if (!hadImage) {
            // 若没有图片替换，直接将替换后的文本写回（统一为单 run，避免 run 拆分问题）
            clearRunsAndSetText(paragraph, replaced);
        }
    }

    /**
     * 清除段落所有 run 的文本内容，并将 newText 写回为单个 run（保留样式不可行时会丢失样式）
     * 如果需要保留样式，需要逐 run 复制样式，这里为了健壮性采用简单写法。
     */
    private static void clearRunsAndSetText(XWPFParagraph paragraph, String newText) {
        List<XWPFRun> runs = paragraph.getRuns();
        if (runs != null) {
            // 清空原 runs 文本
            for (XWPFRun r : runs) {
                r.setText("", 0);
            }
        }
        // 写入新文本为一个 run
        XWPFRun newRun = paragraph.createRun();
        newRun.setText(newText == null ? "" : newText);
    }

    // ---------------------------
    // 表格处理：支持任意行上的 ${table:xxx} 标识
    // ---------------------------
    private static void processTable(XWPFTable table,
                                     Map<String, Object> textParams,
                                     Map<String, InputStream> imageParams,
                                     Map<String, List<Map<String, Object>>> tableParams) throws Exception {
        if (table == null) {return;}

        // 1) 搜索表格中第一个包含 ${table:xxx} 的单元格（支持任意行/列）
        int rows = table.getRows().size();
        for (int r = 0; r < rows; r++) {
            XWPFTableRow row = table.getRow(r);
            if (row == null) {continue;}

            boolean found = false;
            String tableName = null;

            for (XWPFTableCell cell : row.getTableCells()) {
                if (cell == null) {continue;}
                String cellText = cell.getText();
                if (cellText == null) {continue;}

                Matcher matcher = TABLE_PLACEHOLDER_PATTERN.matcher(cellText);
                if (matcher.find()) {
                    tableName = matcher.group(1);
                    found = true;
                    break;
                }
            }

            if (!found) {
                // 未在该行找到标识，处理下一行
                continue;
            }

            // 找到占位符所在行 —— r
            // 模板行约定为：标识行的下一行（r+1）
            int templateRowIndex = r + 1;
            if (templateRowIndex >= table.getRows().size()) {
                // 没有模板行，清理占位符后继续
                cleanPlaceholderInRow(row);
                continue;
            }

            // 从 tableParams 获取数据列表
            if (!tableParams.containsKey(tableName)) {
                // 没有数据也要清理标识（只删除 ${table:xxx}，保留其他文字）
                cleanPlaceholderInRow(row);
                // 不删除模板行（保留模板）
                continue;
            }

            List<Map<String, Object>> dataList = tableParams.get(tableName);
            if (dataList == null || dataList.isEmpty()) {
                // 清理标识并保留模板行
                cleanPlaceholderInRow(row);
                continue;
            }

            // 获取模板行
            XWPFTableRow templateRow = table.getRow(templateRowIndex);
            // 插入数据行：从 templateRowIndex 开始插入（插入位置每次都是 templateRowIndex，插入后原 templateRow 向下移）
            for (int i = 0; i < dataList.size(); i++) {
                Map<String, Object> rowData = dataList.get(i);
                // 在 templateRowIndex 处插入新行
                XWPFTableRow createdRow = table.insertNewTableRow(templateRowIndex + i);
                // 为了简单稳健，这里按模板单元格数量创建新单元格并填文本（不复制单元格样式）
                int cellCount = templateRow.getTableCells().size();
                for (int c = 0; c < cellCount; c++) {
                    XWPFTableCell srcCell = templateRow.getCell(c);
                    XWPFTableCell dstCell = createdRow.addNewTableCell();

                    // 获取模板单元格的完整文本（可能包含 ${col}）
                    String templateText = srcCell.getText();
                    templateText = templateText == null ? "" : templateText;

                    // 先用表格级的占位符（rowData）替换
                    for (Map.Entry<String, Object> e : rowData.entrySet()) {
                        String key = "${" + e.getKey() + "}";
                        String value = e.getValue() == null ? "" : e.getValue().toString();
                        templateText = templateText.replace(key, value);
                    }

                    // 然后用全局 textParams 做二次替换（如果模板单元格中也可能使用全局变量）
                    for (Map.Entry<String, Object> globalEntry : textParams.entrySet()) {
                        String gk = "${" + globalEntry.getKey() + "}";
                        String gv = globalEntry.getValue() == null ? "" : globalEntry.getValue().toString();
                        if (templateText.contains(gk)) {
                            templateText = templateText.replace(gk, gv);
                        }
                    }

                    // 写入目标单元格（写为单段落单 run）
                    dstCell.removeParagraph(0); // 移除默认段落
                    XWPFParagraph p = dstCell.addParagraph();
                    XWPFRun run = p.createRun();
                    run.setText(templateText);
                }
            }

            // 删除模板行（原 templateRow 向下移动到 templateRowIndex + dataList.size()，所以删除 templateRowIndex + dataList.size()）
            table.removeRow(templateRowIndex + dataList.size());

            // 清理标识行（只删除 ${table:xxx}，保留其他文字）
            cleanPlaceholderInRow(row);

            // 处理完当前表格的一个循环后直接返回（假设一个表格只处理第一个匹配的循环）
            return;
        }

        // 如果遍历完没有找到任何 ${table:xxx}，则把表格当成普通表格处理（替换其中的文本/图片占位）
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for (XWPFParagraph p : cell.getParagraphs()) {
                    replaceParagraphTextAndImages(p, textParams, imageParams);
                }
            }
        }
    }

    /**
     * 清理行中所有段落的 ${table:xxx} 占位符（保留行内其它文字）
     * 处理方法：合并段落所有 run 文本 -> replace -> 清空原 runs -> 写入新单 run
     */
    private static void cleanPlaceholderInRow(XWPFTableRow row) {
        if (row == null) {return;}

        for (XWPFTableCell cell : row.getTableCells()) {
            for (XWPFParagraph paragraph : cell.getParagraphs()) {
                List<XWPFRun> runs = paragraph.getRuns();
                if (runs == null || runs.isEmpty()) {continue;}

                // 合并原文本
                StringBuilder sb = new StringBuilder();
                for (XWPFRun run : runs) {
                    String t = run.getText(0);
                    sb.append(t == null ? "" : t);
                }
                String merged = sb.toString();

                // 删除占位符 ${table:xxx}
                String cleaned = merged.replaceAll("\\$\\{table:[^}]+}", "");

                // 如果相同，则不改写
                if (Objects.equals(merged, cleaned)) {continue;}

                // 清空原 runs
                for (XWPFRun run : runs) {run.setText("", 0);}

                // 写回清理后的文本（单 run）
                XWPFRun newRun = paragraph.createRun();
                newRun.setText(cleaned);
            }
        }
    }

    // ---------------------------
    // Helper：把 HTTP/HTTPS 图像转为 InputStream 的示例（供调用方直接使用）
    // ---------------------------
    /**
     * 从 URL（http/https）读取图片为 InputStream。调用者负责关闭返回的 InputStream。
     * 示例：
     *   try (InputStream in = downloadImage("https://...")) {
     *       imageParams.put("signature", in);
     *       fillTemplate(...);
     *   }
     */
    public static InputStream downloadImage(String imageUrl) throws IOException {
        URL u = new URL(imageUrl);
        return u.openStream();
    }

    // ---------------------------
    // Base64 辅助方法（对外）
    // ---------------------------
    public static InputStream base64ToInputStream(String base64) {
        return new ByteArrayInputStream(Base64.getDecoder().decode(base64));
    }

    public static void main(String[] args) throws Exception {
// 1. 文本占位符
        Map<String, Object> textParams = new HashMap<>();
        textParams.put("contractNo", "HT-2025-001");
        textParams.put("userName", "张三");
        textParams.put("date", "2025-11-22");
        textParams.put("amount", "50000");

        // 2. 图片占位符
        InputStream signatureInput = new FileInputStream("D:\\Desktop\\1231312.jpg");
        Map<String, InputStream> imageParams = new HashMap<>();
        imageParams.put("signature", signatureInput);

        // 3. 表格循环
        List<Map<String, Object>> orderList = new ArrayList<>();

        Map<String, Object> row1 = new HashMap<>();
        row1.put("item", "产品A");
        row1.put("price", "100");
        row1.put("qty", "2");
        orderList.add(row1);

        Map<String, Object> row2 = new HashMap<>();
        row2.put("item", "产品B");
        row2.put("price", "200");
        row2.put("qty", "1");
        orderList.add(row2);

        Map<String, List<Map<String, Object>>> tableParams = new HashMap<>();
        tableParams.put("orderTable", orderList);


        List<Map<String, Object>> orderList2 = new ArrayList<>();
        Map<String, Object> row12 = new HashMap<>();
        row12.put("item", "产品A2");
        row12.put("price", "100");
        row12.put("qty", "2");
        orderList2.add(row12);

        Map<String, Object> row22 = new HashMap<>();
        row22.put("item", "产品B2");
        row22.put("price", "200");
        row22.put("qty", "1");
        orderList2.add(row22);

        tableParams.put("orderTable2", orderList2);

        String path = "D:\\tt.docx";
        byte[] fileBytes = Files.readAllBytes(Paths.get(path));
        // word → base64
        String base64 = Base64.getEncoder().encodeToString(fileBytes);

        // 4. 模板 Base64（可以通过工具先生成 Base64）
        String templateBase64 = base64;

        // 调用 fillTemplate
        String resultBase64 = WordBase64ReplaceUtil.fillTemplate(templateBase64, textParams, imageParams, tableParams);

        // 写入文件
        byte[] resultBytes = java.util.Base64.getDecoder().decode(resultBase64);
        java.io.FileOutputStream fos = new java.io.FileOutputStream("D:\\out.docx");
        fos.write(resultBytes);
        fos.close();

        System.out.println("Word生成完成：filled_contract.docx");
    }
}
