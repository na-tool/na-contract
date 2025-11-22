package com.na.contract.utils;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.util.Units;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;

public class WordBase64ReplaceUtil {

    /**
     * 替换文本、图片、表格循环
     * @param templateBase64 Word 模板 Base64
     * @param textParams 文本占位符 Map<key,value>
     * @param imageParams 图片占位符 Map<key, InputStream>
     * @param tableParams 表格循环 Map<tableName, List<Map<列名, 值>>>
     * @return 替换后的 Word Base64
     */
    public static String fillTemplate(
            String templateBase64,
            Map<String, Object> textParams,
            Map<String, InputStream> imageParams,
            Map<String, List<Map<String, Object>>> tableParams
    ) throws Exception {
        if (templateBase64 == null || templateBase64.isEmpty()) {
            throw new IllegalArgumentException("模板 Base64 不能为空");
        }

        textParams = textParams == null ? Collections.emptyMap() : textParams;
        imageParams = imageParams == null ? Collections.emptyMap() : imageParams;
        tableParams = tableParams == null ? Collections.emptyMap() : tableParams;

        // Base64 -> Word
        byte[] wordBytes = Base64.getDecoder().decode(templateBase64);
        try (XWPFDocument doc = new XWPFDocument(new ByteArrayInputStream(wordBytes))) {

            // 替换段落文本
            for (XWPFParagraph paragraph : doc.getParagraphs()) {
                replaceParagraphText(paragraph, textParams, imageParams);
            }

            // 替换表格内容
            for (XWPFTable table : doc.getTables()) {
                replaceTable(table, textParams, imageParams, tableParams);
            }

            // 输出 Base64
            try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
                doc.write(baos);
                return Base64.getEncoder().encodeToString(baos.toByteArray());
            }
        }
    }

    /** 替换段落内容，包括文本和图片 */
    private static void replaceParagraphText(XWPFParagraph paragraph, Map<String, Object> textParams, Map<String, InputStream> imageParams) throws Exception {
        if (paragraph.getRuns() == null) {return;}

        for (int i = 0; i < paragraph.getRuns().size(); i++) {
            XWPFRun run = paragraph.getRuns().get(i);
            String text = run.getText(0);
            if (text == null) {continue;}

            // 替换文本
            for (Map.Entry<String, Object> entry : textParams.entrySet()) {
                String key = "${" + entry.getKey() + "}";
                String value = entry.getValue() == null ? "" : entry.getValue().toString();
                if (text.contains(key)) {
                    text = text.replace(key, value);
                }
            }
            run.setText(text, 0);

            // 替换图片
            for (Map.Entry<String, InputStream> imgEntry : imageParams.entrySet()) {
                String key = "${" + imgEntry.getKey() + "}";
                if (text.contains(key)) {
                    run.setText("", 0);
                    run.addPicture(
                            imgEntry.getValue(),
                            XWPFDocument.PICTURE_TYPE_PNG,
                            imgEntry.getKey(),
                            Units.toEMU(150),
                            Units.toEMU(150)
                    );
                }
            }
        }
    }

    /** 替换表格文本、循环表格 */
    private static void replaceTable(XWPFTable table, Map<String, Object> textParams, Map<String, InputStream> imageParams, Map<String, List<Map<String, Object>>> tableParams) throws Exception {
        if (table.getRows().isEmpty()) {return;}

        // 判断是否循环表格
        XWPFTableRow firstRow = table.getRow(0);
        boolean isLoop = false;
        String tableName = null;

        for (XWPFTableCell cell : firstRow.getTableCells()) {
            String cellText = cell.getText();
            if (cellText != null && cellText.contains("${table:")) {
                isLoop = true;
                tableName = cellText.substring(cellText.indexOf("${table:") + 8, cellText.indexOf("}"));
                break;
            }
        }

        if (isLoop && tableParams.containsKey(tableName)) {
            List<Map<String, Object>> rowsData = tableParams.get(tableName);
            if (rowsData != null && rowsData.size() > 0) {
                XWPFTableRow templateRow = table.getRow(1); // 第二行模板
                for (Map<String, Object> rowData : rowsData) {
                    XWPFTableRow newRow = table.createRow();
                    for (int i = 0; i < templateRow.getTableCells().size(); i++) {
                        XWPFTableCell templateCell = templateRow.getCell(i);
                        XWPFTableCell newCell = newRow.getCell(i);
                        if (newCell == null) {newCell = newRow.createCell();}
                        newCell.setText(templateCell.getText() != null ? templateCell.getText() : "");

                        // 替换占位符
                        for (Map.Entry<String, Object> entry : rowData.entrySet()) {
                            String key = "${" + entry.getKey() + "}";
                            String value = entry.getValue() == null ? "" : entry.getValue().toString();
                            newCell.setText(newCell.getText().replace(key, value));
                        }
                    }
                }
                // 移除模板行
                table.removeRow(1);
            }
        } else {
            // 普通表格文本替换
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        replaceParagraphText(paragraph, textParams, imageParams);
                    }
                }
            }
        }
    }

    // Base64 -> InputStream
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
