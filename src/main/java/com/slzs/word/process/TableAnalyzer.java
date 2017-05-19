package com.slzs.word.process;

import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.slzs.util.ObjectUtil;
import com.slzs.word.model.WordData;

/**
 * 表格处理
 * @author 北京拓尔思信息技术股份有限公司
 * @author slzs
 * 2017年4月14日 下午3:00:33
 */
class TableAnalyzer {
    private DataFillAnalyzer dfAnalyzer;
    private ParagraphAnalyzer paragraphAnalyzer;
    private StyleAnalyzer styleAnalyzer;

    public TableAnalyzer(DataFillAnalyzer dfAnalyzer,ParagraphAnalyzer paragraphAnalyzer,StyleAnalyzer styleAnalyzer) {
        this.dfAnalyzer = dfAnalyzer;
        this.paragraphAnalyzer = paragraphAnalyzer;
        this.styleAnalyzer = styleAnalyzer;
    }
    /**
     * 表格标签解析
     * @author: slzs
     * 2015-2-5 下午2:12:38
     * @param document
     * @param tableMap
     */
    void analy(XWPFDocument document, Map<String, List<WordData>> tableMap) {
        // 表格标签解析

        // 文档表格中段落标记梳理
        List<XWPFTable> tableList = document.getTables();
        for (XWPFTable table : tableList) { // 循环表格处理

            List<WordData> rowDataList = null; // 多行数据
            List<XWPFTableCell> copyCellList = null;
            List<XWPFTableRow> rowList = table.getRows();
            XWPFTableRow copyRow = null;
            if (ObjectUtil.isNotEmpty(rowList)) {
                for (int i = 0; i < rowList.size(); i++) { // 循环处理行
                    XWPFTableRow row = rowList.get(i);
                    List<XWPFTableCell> cellList = row.getTableCells();
                    XWPFTableCell firstCell = cellList.get(0);
                    String firstCellText = firstCell.getText();
                    if (firstCellText.matches(".*\\$\\{.*:rows\\}.*")) {
                        boolean hasData = false;
                        if (tableMap != null) {
                            for (String key : tableMap.keySet()) {
                                if (firstCellText.contains("${" + key + ":rows}")) {
                                    hasData = true;
                                    rowDataList = tableMap.get(key);
                                    copyCellList = row.getTableCells(); // 复制本行
                                    copyRow = row;
                                    break;
                                }
                            }
                        }
                        if (!hasData) {
                            // table.removeRow(i--); 可删除一行
                            for (XWPFTableCell cell : cellList) {
                                // 设置整行为空数据
                                while (cell.getParagraphs().size() > 0) {
                                    cell.removeParagraph(0); // 删除数据
                                }
                                cell.addParagraph(); // 增加空段占位，否则格式不符无法打开
                            }
                        }
                    }
                }
            }

            if (ObjectUtil.isNotEmpty(rowDataList) && copyRow != null && ObjectUtil.isNotEmpty(copyCellList)) {
                // 根据数据复制行数
                for (int i = 1; i < rowDataList.size(); i++) {
                    XWPFTableRow newRow = table.createRow();
                    newRow.setHeight(copyRow.getHeight()); // 设置相同行高
                    List<XWPFTableCell> cellList = newRow.getTableCells();
                    for (int c = 0; c < cellList.size(); c++) {
                        XWPFTableCell cell = cellList.get(c);
                        XWPFTableCell copyCell = copyCellList.get(c);
                        while (cell.getParagraphs().size() > 0) {
                            cell.removeParagraph(0); // 清理多余段落
                        }
                        List<XWPFParagraph> pList = copyCell.getParagraphs();
                        for (XWPFParagraph p : pList) { // 复制单元格内段落
                            XWPFParagraph newP = cell.addParagraph();
                            styleAnalyzer.styleClone(p, newP);// 克隆段落样式
                            List<XWPFRun> runList = p.getRuns();
                            if (ObjectUtil.isNotEmpty(runList)) {
                                for (XWPFRun run : runList) { // 复制段落内字符
                                    XWPFRun newRun = newP.createRun();
                                    newRun.getCTR().set(run.getCTR().copy());
                                }
                            }
                        }

                    }
                }
            }
        }
    }
    


    /**
     * 表格数据内容处理
     * 
     * @author: slzs 
     * 2015-2-12 下午2:09:10
     * @param document
     * @param table
     * @param data
     * 
     */
    void setTableContent(XWPFDocument document, XWPFTable table, WordData data) {
        List<XWPFTableRow> rowList = table.getRows();
        for (XWPFTableRow row : rowList) { // 循环处理行
            List<XWPFTableCell> cellList = row.getTableCells();
            XWPFTableCell firstCell = cellList.get(0);
            for (XWPFTableCell cell : cellList) { // 循环处理列
                List<XWPFParagraph> cellPghList = cell.getParagraphs();
                for (XWPFParagraph cellPgp : cellPghList) { // 处理单元格内段落
                    paragraphAnalyzer.setParagraphContent(document, cellPgp, data);
                }
            }

            if (data.hasTable()) {
                Map<String, List<WordData>> tableMap = data.getTableMap();
                for (String tableKey : tableMap.keySet()) {
                    if (firstCell.getText().contains("${" + tableKey + ":rows}")) {
                        tableNext(data, tableKey);// 下一行数据
                        break;
                    }
                }
            }
        }
    }


    /**
     * 加载指定表格下一行数据
     * @author: slzs
     * 2015-2-5 下午3:28:40
     * @param data
     * @param tableKey
     * @return WordData
     */
    private WordData tableNext(WordData data, String tableKey) {
        if (data.hasTable()) {
            // 表格填充数据
            Map<String, List<WordData>> tableMap = data.getTableMap();
            List<WordData> wordDataList = tableMap.get(tableKey);
            dfAnalyzer.mergeMove(data, tableKey, wordDataList);
        }
        return data;
    }
    
    /**
     * 全部表格加载下一行数据
     * 
     * @author: slzs 
     * 2015-2-5 下午3:28:40
     * @param data
     * @return WordData
     * 
     */
    public WordData tableAllNext(WordData data) {
        if (data.hasTable()) {
            // 表格填充数据
            Map<String, List<WordData>> tableMap = data.getTableMap();
            Set<String> keySet = tableMap.keySet();
            for (String tableKey : keySet) {

                List<WordData> wordDataList = tableMap.get(tableKey);

                dfAnalyzer.mergeMove(data, tableKey, wordDataList);

            }
        }
        return data;
    }

}
