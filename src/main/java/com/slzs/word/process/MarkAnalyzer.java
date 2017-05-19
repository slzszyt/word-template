package com.slzs.word.process;

import java.util.List;

import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.slzs.util.StringUtil;

import lombok.extern.log4j.Log4j2;

/**
 * 标记处理
 * @author 北京拓尔思信息技术股份有限公司
 * @author slzs
 * 2017年4月14日 下午1:17:43
 */
@Log4j2
public class MarkAnalyzer {

    private boolean       markClear;

    private UBBAnalyzer   ubb;

    private StyleAnalyzer styleAnalyzer;
    
    /**
     * @param styleAnalyzer
     */
    public MarkAnalyzer(StyleAnalyzer styleAnalyzer) {
        this.styleAnalyzer = styleAnalyzer;
    }

    /**
     * 梳理文档${标记}，保留标记整体样式
     * word字符存储时有各种情况会将连续字符存到不同的XWPFRun里，对处理造成困难
     * 本方法将连续的标记梳理到相同的run里，便于处理
     * @author: slzs
     * 2014-4-16 下午1:34:48
     * @param document 
     * 
     */
      public void markComb(XWPFDocument document) {

        if (log.isDebugEnabled()) {
            log.debug("梳理文档${标记}开始...");
        }

        // 文档段落标记梳理
        List<XWPFParagraph> paragraphList = document.getParagraphs();
        markComb(paragraphList);

        // 文档表格中段落标记梳理
        List<XWPFTable> tableList = document.getTables();
        for (XWPFTable table : tableList) { // 循环表格处理
            List<XWPFTableRow> rowList = table.getRows();
            for (XWPFTableRow row : rowList) { // 循环处理行
                List<XWPFTableCell> cellList = row.getTableCells();
                for (XWPFTableCell cell : cellList) { // 循环处理列
                    markComb(cell.getParagraphs()); // 每个单元格中的标记梳理
                }
            }
        }

        if (log.isDebugEnabled()) {
            log.debug("梳理文档${标记}结束...");
        }
    }

    /**
     * 梳理段落标记
     * @author: slzs
     * 2015-2-5 下午2:04:38
     * @param paragraphList 
     * 
     */
    private void markComb(List<XWPFParagraph> paragraphList) {

        boolean markBeginMissing = false; // 出现一半开始符
        int markBeginRunIndex = -1; // 开始标记符出现的run位置
        StringBuffer markKeySB = null; // 叠加标记字符串

        List<XWPFRun> runList;
        XWPFParagraph paragraph;
        XWPFRun markRun;

        for (int i = 0; i < paragraphList.size(); i++) {
            paragraph = paragraphList.get(i);
            runList = paragraph.getRuns();
            for (int p = 0; p < runList.size(); p++) {
                markRun = runList.get(p);
                final String runTempText = markRun.toString().trim(); // 当前循环run文本
                if (StringUtil.isNotEmpty(runTempText)) {
                    int markEndStrIndex;
                    int markBeginStrIndex;
                    if (markBeginRunIndex > -1) {
                        /* 找到开始标记，重新组织标记*/
                        markEndStrIndex = runTempText.lastIndexOf("}");// 结束标记位置
                        if (markEndStrIndex > -1) {// 标记结束
                            if (runTempText.length() > ++markEndStrIndex) {
                                markKeySB.append(runTempText.substring(0, markEndStrIndex));
                                markRun.setText(runTempText.substring(markEndStrIndex), 0); //截去当前位置标记文本
                            } else {
                                markKeySB.append(runTempText);
                                //                                paragraph.removeRun(p--);//去掉当前位置标记
                                markRun.setText("", 0);
                            }
                            // 重写标记
                            markRun = runList.get(markBeginRunIndex);
                            markRun.setText(markKeySB.toString(), 0); // setText是替换文本并保留原样式不变的最佳方案
                            markBeginRunIndex = -1;
                        } else {
                            //                            paragraph.removeRun(p--);
                            markRun.setText("", 0);
                            markKeySB.append(runTempText);
                        }
                    } else {
                        /* 定位开始标记 */
                        markBeginStrIndex = runTempText.lastIndexOf("${");
                        if (markBeginStrIndex > -1) {
                            // 标记开始后是否有结束符
                            markEndStrIndex = runTempText.substring(markBeginStrIndex).lastIndexOf("}");
                            if (markEndStrIndex < 0) {
                                markKeySB = new StringBuffer(runTempText);
                                markBeginRunIndex = p; // 记录run索引
                            }
                        } else if (markBeginMissing && runTempText.startsWith("{")) {
                            // 标记开始
                            markEndStrIndex = runTempText.lastIndexOf("}");
                            markBeginStrIndex = runTempText.lastIndexOf("${");// 最后一个开始符
                            if (markEndStrIndex < 0 || markEndStrIndex < markBeginStrIndex) { // 没有结束或结束符在开始符之前
                                markKeySB = new StringBuffer("$");
                                markKeySB.append(runTempText);
                                //                                paragraph.removeRun(p--); // 移除当前位置
                                markRun.setText("", 0);
                                markBeginRunIndex = p - 1; // 记录run索引
                            }
                        }
                    }
                    markBeginMissing = runTempText.endsWith("$"); // 文本最后包含部分标记
                }
            }
        }
    }

    /**
     * 清理文档${标记}
     * @author: slzs
     * 2014-4-16 下午8:36:12
     * @param document 
     */
    public void markClear(XWPFDocument document) {
        List<IBodyElement> bodyList = document.getBodyElements();
        StringBuffer iteratorKey = new StringBuffer();
        for (int bodyIndex = 0; bodyIndex < bodyList.size(); bodyIndex++) {
            IBodyElement bodyElement = bodyList.get(bodyIndex);

            if (bodyElement.getElementType() == BodyElementType.PARAGRAPH) {
                // 段落处理
                XWPFParagraph paragraph = (XWPFParagraph) bodyElement;
                boolean empty = markClearForParagraph(document, paragraph, iteratorKey);
                if (empty)
                    // 段落没有内容，清除空行
                    document.removeBodyElement(bodyIndex--);
            } else if (bodyElement.getElementType() == BodyElementType.TABLE) {
                // 表格处理
                XWPFTable table = (XWPFTable) bodyElement;

                List<XWPFTableRow> rowList = table.getRows();
                for (XWPFTableRow row : rowList) { // 循环处理行
                    List<XWPFTableCell> cellList = row.getTableCells();
                    for (XWPFTableCell cell : cellList) { // 循环处理列
                        List<XWPFParagraph> cellPghList = cell.getParagraphs();
                        for (int pos = 0; pos < cellPghList.size(); pos++) { // 处理单元格内段落
                            /*boolean empty = */markClearForParagraph(document, cellPghList.get(pos), iteratorKey);
                            //                            if (empty)
                            //                                // 段落没有内容，清除空行
                            //                                document.removeBodyElement(bodyIndex--);
                        }
                    }
                }

            }
        }
        markClear = true;
    }

    /**
     * 清理段落标记
     * @author: slzs
     * 2015-2-12 下午2:35:11
     * @param document
     * @param paragraph
     * @param 迭代key
     * @return 是否删除空行
     */
    private boolean markClearForParagraph(XWPFDocument document, XWPFParagraph paragraph, StringBuffer iteratorKey) {
        boolean empty = false;
        List<XWPFRun> runList = paragraph.getRuns();
        String content = runList.toString();
        if (iteratorKey.length() > 0 && content.matches(".*\\$\\{" + iteratorKey + "\\:end\\}.*")) {
            iteratorKey.delete(0, iteratorKey.length());
        }
        if (iteratorKey.length() > 0) {
            empty = true;
            if (log.isDebugEnabled()) {
                log.debug("清理文档  "+iteratorKey+"空迭代数据");
            }
        } else if (runList.toString().matches(".*\\$\\{.*\\}.*")) {// 包含key标记
            empty = true;
            for (int p = 0; p < runList.size(); p++) {
                XWPFRun markRun = runList.get(p);
                String runTextTemp = markRun.toString();
                if (log.isDebugEnabled()) {
                    log.debug("清理文档${标记} 清理前:" + runTextTemp);
                }

                if (iteratorKey.length() == 0 && runTextTemp.matches(".*\\$\\{.*\\:start\\}.*")) {
                    // 发现start标记
                    String key = runTextTemp.replaceAll(".*\\$\\{(.*)\\:start\\}.*", "$1"); //取key${key:start}
                    iteratorKey.append(key);
                }

                if (iteratorKey.length() > 0 && runTextTemp.matches(".*\\$\\{" + iteratorKey + "\\:end\\}.*")) {
                    // 发现end标记
                    iteratorKey.delete(0, iteratorKey.length());
                }

                // 去除标记
                runTextTemp = iteratorKey.length() > 0 ? "" : runTextTemp.replaceAll("\\$\\{.*?\\}", "");

                markRun.setText(runTextTemp, 0);
                empty = StringUtil.isEmpty(runTextTemp); // 是否空段落

                if (log.isDebugEnabled()) {
                    log.debug("清理文档${标记} 清理后:" + runTextTemp);
                }
            }
        }
        return empty;
    }

    /**
     * ubb标记解析
     * @author slzs 
     * 2017年5月9日 下午2:08:50
     * @param document
     * @throws Exception 
     */
    public void ubbAnalysis(XWPFDocument document) throws Exception {
        if (document != null) {
            if (!markClear){
                throw new Exception("UBBAnalysis need call markClear first!");
            }
            if (ubb == null)
                ubb = new UBBAnalyzer(styleAnalyzer);
            ubb.analysis(document);
        }
    }
}
