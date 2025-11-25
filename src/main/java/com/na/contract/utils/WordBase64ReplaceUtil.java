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

    // 表格循环占位符，例如：${table:orderTable}
    private static final Pattern TABLE_PLACEHOLDER_PATTERN = Pattern.compile("\\$\\{table:([^}]+)}");
    // 文本/图片占位符正则（用于段落级别按 key 替换时可选）
    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("\\$\\{([^}]+)}");

    /**
     * 将 Base64 的 Word 文档进行占位符替换，返回替换后的 Base64 Word
     *
     * @param templateBase64 Word 模板 Base64，不可为空
     * @param textParams     文本占位符  {@code Map<key, value>}，可为 null
     * @param imageParams    图片占位符 {@code Map<imgKey, InputStream>}，可为 null（InputStream 由调用方管理关闭）
     * @param tableParams    表格循环 {@code Map<tableName, List<Map<列名, 值>>>}，可为 null
     * @return 替换后的 Word 的 Base64 字符串
     * @throws Exception 转换失败异常
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

        // 避免空指针
        textParams = textParams == null ? Collections.emptyMap() : textParams;
        imageParams = imageParams == null ? Collections.emptyMap() : imageParams;
        tableParams = tableParams == null ? Collections.emptyMap() : tableParams;

        byte[] bytes = Base64.getDecoder().decode(templateBase64);

        try (XWPFDocument doc = new XWPFDocument(new ByteArrayInputStream(bytes))) {

            // -----------------------
            // 1. 处理文档自由段落
            // -----------------------
            for (XWPFParagraph p : doc.getParagraphs()) {
                replaceParagraphTextAndImages(p, textParams, imageParams);
            }

            // -----------------------
            // 2. 处理所有表格
            // -----------------------
            for (XWPFTable table : doc.getTables()) {
                processTable(table, textParams, imageParams, tableParams);
            }

            // -----------------------
            // 3. 输出 Base64
            // -----------------------
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            doc.write(baos);
            return Base64.getEncoder().encodeToString(baos.toByteArray());
        }
    }


    // ---------------------------
// 段落替换（文本 + 图片，支持多 run 和换行）
// ---------------------------
    private static void replaceParagraphTextAndImages(XWPFParagraph paragraph,
                                                      Map<String, Object> textParams,
                                                      Map<String, InputStream> imageParams) throws Exception {
        if (paragraph == null || paragraph.getRuns() == null) {return;}

        // 1. 合并段落所有 run 文本
        StringBuilder merged = new StringBuilder();
        for (XWPFRun r : paragraph.getRuns()) {
            merged.append(r.getText(0) == null ? "" : r.getText(0));
        }

        String text = merged.toString();

        // 2. 替换文本占位符
        for (Map.Entry<String, Object> entry : textParams.entrySet()) {
            String key = "${" + entry.getKey() + "}";
            if (text.contains(key)) {
                Object value = entry.getValue();
                text = text.replace(key, value == null ? "" : value.toString());
            }
        }

        // 3. 替换图片占位符
        for (Map.Entry<String, InputStream> entry : imageParams.entrySet()) {
            String imgKey = "${" + entry.getKey() + "}";
            if (text.contains(imgKey)) {
                String[] parts = text.split(Pattern.quote(imgKey), -1);

                // 占位符前文本
                clearRuns(paragraph);
                XWPFRun beforeRun = paragraph.createRun();
                insertTextWithLineBreaks(beforeRun, parts[0]);

                // 图片
                XWPFRun imgRun = paragraph.createRun();
                try (InputStream in = entry.getValue()) {
                    imgRun.addPicture(in, XWPFDocument.PICTURE_TYPE_PNG, entry.getKey(),
                            Units.toEMU(150), Units.toEMU(150));
                }

                // 占位符后文本
                if (parts.length > 1 && !parts[1].isEmpty()) {
                    XWPFRun afterRun = paragraph.createRun();
                    insertTextWithLineBreaks(afterRun, parts[1]);
                }

                return; // 图片处理后退出
            }
        }

        // 4. 写回文本（支持换行）
        clearRuns(paragraph);
        XWPFRun run = paragraph.createRun();
        insertTextWithLineBreaks(run, text);
    }

    /**
     * 将文本中的 \n 转为 Word 换行
     */
    private static void insertTextWithLineBreaks(XWPFRun run, String text) {
        String[] lines = text.split("\\n", -1);
        for (int i = 0; i < lines.length; i++) {
            if (i > 0) {run.addBreak();}
            run.setText(lines[i], i == 0 ? 0 : -1);
        }
    }


    /**
     * 清除段落所有 run 的文本内容，并将 newText 写回为单个 run（保留样式不可行时会丢失样式）
     * 如果需要保留样式，需要逐 run 复制样式，这里为了健壮性采用简单写法。
     */
    private static void clearRuns(XWPFParagraph p) {
        if (p.getRuns() == null) {return;}
        for (XWPFRun r : p.getRuns()) {
            r.setText("", 0);
        }
    }




    // ============================================================================================
    //   表格处理（表格循环 + 非循环区域的普通变量替换）
    // ============================================================================================
    private static void processTable(
            XWPFTable table,
            Map<String, Object> textParams,
            Map<String, InputStream> imageParams,
            Map<String, List<Map<String, Object>>> tableParams
    ) throws Exception {

        if (table == null){ return;}

        int rowCount = table.getRows().size();

        for (int r = 0; r < rowCount; r++) {

            XWPFTableRow row = table.getRow(r);
            if (row == null) {continue;}

            boolean foundTableFlag = false;
            String tableName = null;

            // -------------------------------
            // ⭐ 扫描每个单元格：
            //   ① 替换普通变量（关键修复点！）
            //   ② 查找 ${table:xxx}
            // -------------------------------
            for (XWPFTableCell cell : row.getTableCells()) {

                replaceTextInTableCell(cell, textParams);  // ⭐ 修复点：普通变量替换

                String text = cell.getText();
                if (text == null) {continue;}

                Matcher m = TABLE_PLACEHOLDER_PATTERN.matcher(text);
                if (m.find()) {
                    tableName = m.group(1);
                    foundTableFlag = true;
                    break;
                }
            }

            if (!foundTableFlag) {continue;}

            // -------------------------------
            // 找到 ${table:xxx} 所在行 r
            // 下一行 (r+1) 作为模板行
            // -------------------------------
            int templateRowIndex = r + 1;
            if (templateRowIndex >= table.getRows().size()) {
                // 没有模板行，清理占位符后继续
                cleanPlaceholderInRow(row);
                continue;
            }

            XWPFTableRow templateRow = table.getRow(templateRowIndex);

            // 获取循环数据
            List<Map<String, Object>> dataList = tableParams.get(tableName);

            // 无数据时，只清理占位符
            if (dataList == null || dataList.isEmpty()) {
                // 没有数据也要清理标识（只删除 ${table:xxx}，保留其他文字）
                cleanPlaceholderInRow(row);
                // 不删除模板行（保留模板）
                continue;
            }

            // -------------------------------
            // 插入数据行
            // -------------------------------
            for (int i = 0; i < dataList.size(); i++) {
                // 在 templateRowIndex 处插入新行
                XWPFTableRow newRow = table.insertNewTableRow(templateRowIndex + i);
                // 为了简单稳健，这里按模板单元格数量创建新单元格并填文本（不复制单元格样式）
                int cellCount = templateRow.getTableCells().size();
                Map<String, Object> rowData = dataList.get(i);

                for (int c = 0; c < cellCount; c++) {

                    XWPFTableCell srcCell = templateRow.getCell(c);
                    XWPFTableCell dstCell = newRow.addNewTableCell();
                    // 获取模板单元格的完整文本（可能包含 ${col}）
                    String txt = srcCell.getText();
                    txt = txt == null ? "" : txt;

                    // 表格级变量替换
                    for (Map.Entry<String, Object> e : rowData.entrySet()) {
                        txt = txt.replace("${" + e.getKey() + "}",
                                e.getValue() == null ? "" : e.getValue().toString());
                    }

                    // 全局变量替换（如 ${taskTotal}）
                    for (Map.Entry<String, Object> e : textParams.entrySet()) {
                        txt = txt.replace("${" + e.getKey() + "}",
                                e.getValue() == null ? "" : e.getValue().toString());
                    }

                    // 写入，然后用全局 textParams 做二次替换（如果模板单元格中也可能使用全局变量）
                    dstCell.removeParagraph(0);
                    XWPFParagraph p = dstCell.addParagraph();
                    XWPFRun run = p.createRun();
                    run.setText(txt);

                    // 再执行一次完整段落替换（保留兼容性）
                    replaceParagraphTextAndImages(p, textParams, imageParams);
                }
            }

            // 删除模板行
            table.removeRow(templateRowIndex + dataList.size());

            // 清理 ${table:xxx}
            cleanPlaceholderInRow(row);

            return;
        }

        // -------------------------------
        // 表格没有循环占位符 → 普通表格
        // -------------------------------
        for (XWPFTableRow r : table.getRows()) {
            for (XWPFTableCell cell : r.getTableCells()) {
                for (XWPFParagraph p : cell.getParagraphs()) {
                    replaceParagraphTextAndImages(p, textParams, imageParams);
                }
            }
        }
    }



    /**
     * 删除表格行中的所有 ${table:xxx} 占位符，保留其他文字
     */
    private static void cleanPlaceholderInRow(XWPFTableRow row) {
        if (row == null) {return;}

        for (XWPFTableCell cell : row.getTableCells()) {
            for (XWPFParagraph p : cell.getParagraphs()) {

                List<XWPFRun> runs = p.getRuns();
                if (runs == null || runs.isEmpty()) {continue;}

                StringBuilder sb = new StringBuilder();
                for (XWPFRun r : runs) {
                    sb.append(r.getText(0) == null ? "" : r.getText(0));
                }

                String merged = sb.toString();
                String cleaned = merged.replaceAll("\\$\\{table:[^}]+}", "");

                if (!cleaned.equals(merged)) {
                    clearRuns(p);
                    XWPFRun nr = p.createRun();
                    nr.setText(cleaned);
                }
            }
        }
    }



    /**
     * ⭐ 新增：表格单元格内的文本变量替换（不含图片）
     * 解决 `${taskTotal}` 不在段落层、而在表格层不被替换的问题
     */
    private static void replaceTextInTableCell(XWPFTableCell cell, Map<String, Object> textParams) {

        for (XWPFParagraph p : cell.getParagraphs()) {

            List<XWPFRun> runs = p.getRuns();
            if (runs == null || runs.isEmpty()) {continue;}

            // 合并文本
            StringBuilder sb = new StringBuilder();
            for (XWPFRun r : runs) {sb.append(r.getText(0) == null ? "" : r.getText(0));}
            String merged = sb.toString();

            // 替换
            String replaced = merged;
            for (Map.Entry<String, Object> e : textParams.entrySet()) {
                replaced = replaced.replace("${" + e.getKey() + "}",
                        e.getValue() == null ? "" : e.getValue().toString());
            }

            if (!merged.equals(replaced)) {
                clearRuns(p);
                XWPFRun nr = p.createRun();
                nr.setText(replaced);
            }
        }
    }



    // ============================================================================================
    // 图片下载（供调用者使用）
    // ============================================================================================
    public static InputStream downloadImage(String url) throws IOException {
        return new URL(url).openStream();
    }

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
