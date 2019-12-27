package com.slzs.word.process;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;

/**
 * 样式处理
 * @author slzs
 * 2017年4月14日 下午1:17:43
 */
public class StyleAnalyzer {

    /**
     * XWPFRun 样式克隆
     * @author: slzs
     * 2014-4-16 下午8:27:51
     * @param cpSrc 样式源
     * @param cpTo 目标
     */
    void styleClone(XWPFRun cpSrc, XWPFRun cpTo) {
        if (cpSrc == null) {
            return;
        }
        if (cpTo == null) {
            return;
        }

        CTRPr cpToRPR = cpTo.getCTR().getRPr();
        CTRPr cpSrcRPR = cpSrc.getCTR().getRPr();

        cpTo.setBold(cpSrc.isBold()); // 加粗

        cpTo.setColor(cpSrc.getColor()); // 字颜色
        cpTo.setCapitalized(cpSrc.isCapitalized()); //首字母大写
        cpTo.setCharacterSpacing(cpSrc.getCharacterSpacing()); // 字符空隙

        cpTo.setDoubleStrikethrough(cpSrc.isDoubleStrikeThrough()); // 双删除线
        cpTo.setEmbossed(cpSrc.isEmbossed()); // 重影效果

        cpTo.setFontFamily(cpSrc.getFontFamily()); // 字体
        if (cpSrc.getFontSize() > 0) {
            cpTo.setFontSize(cpSrc.getFontSize()); // 字号
        }
        cpTo.setImprinted(cpSrc.isImprinted()); // 与setEmbossed类似效果
        cpTo.setItalic(cpSrc.isItalic()); // 倾斜
        cpTo.setKerning(cpSrc.getKerning()); // 字距
        cpTo.setShadow(cpSrc.isShadowed()); // 阴影
        cpTo.setSmallCaps(cpSrc.isSmallCaps()); // 一种字母格式
        cpTo.setStrikeThrough(cpSrc.isStrikeThrough()); // 单删除线
        cpTo.setStyle(cpSrc.getStyle()); // 应用的统一样式ID
        cpTo.setTextScale(cpSrc.getTextScale()); // 字体缩放

        cpTo.setUnderline(cpSrc.getUnderline()); // 字底线
        cpTo.setTextPosition(cpSrc.getTextPosition()); // 行距位置，基于行基准线的上下偏移量

        if (cpSrcRPR == null) {
            return;
        }
        if (cpToRPR == null) {
            cpToRPR = cpTo.getCTR().addNewRPr();
        }

        if (cpSrcRPR.isSetHighlight()) { // 高亮（背景色）
            cpTo.setTextHighlightColor(cpSrc.getTextHightlightColor().toString());
        }
        if (cpSrcRPR.isSetShd()) { // 设置底纹样式
            CTShd copySourceshd = cpSrcRPR.getShd();
            CTShd copyToshd = cpToRPR.addNewShd();
            copyToshd.set(copySourceshd.copy());
        }
    }

    /**
     * XWPFParagraph 段落样式克隆
     * @author: slzs
     * 2015-1-29 下午3:57:38
     * @param copySource 样式源
     * @param copyTo 目标
     */
    void styleClone(XWPFParagraph copySource, XWPFParagraph copyTo) {

        copyTo.setAlignment(copySource.getAlignment());

        copyTo.setBorderBetween(copySource.getBorderBetween());
        copyTo.setBorderBottom(copySource.getBorderBottom());
        copyTo.setBorderLeft(copySource.getBorderLeft());
        copyTo.setBorderRight(copySource.getBorderRight());
        copyTo.setBorderTop(copySource.getBorderTop());

        copyTo.setFirstLineIndent(copySource.getFirstLineIndent());
        copyTo.setFontAlignment(copySource.getFontAlignment());

        if (copySource.getIndentationFirstLine() > -1)
            copyTo.setIndentationFirstLine(copySource.getIndentationFirstLine());
        if (copySource.getIndentationHanging() > -1)
            copyTo.setIndentationHanging(copySource.getIndentationHanging());
        if (copySource.getIndentationLeft() > -1)
            copyTo.setIndentationLeft(copySource.getIndentationLeft());
        if (copySource.getIndentationRight() > -1)
            copyTo.setIndentationRight(copySource.getIndentationRight());
        if (copySource.getIndentFromLeft() > -1)
            copyTo.setIndentFromLeft(copySource.getIndentFromLeft());
        if (copySource.getIndentFromRight() > -1)
            copyTo.setIndentFromRight(copySource.getIndentFromRight());

        if (copySource.getSpacingAfter() > -1)
            copyTo.setSpacingAfter(copySource.getSpacingAfter());
        if (copySource.getSpacingAfterLines() > -1)
            copyTo.setSpacingAfterLines(copySource.getSpacingAfterLines());
        if (copySource.getSpacingBefore() > -1)
            copyTo.setSpacingBefore(copySource.getSpacingBefore());
        if (copySource.getSpacingBeforeLines() > -1)
            copyTo.setSpacingBeforeLines(copySource.getSpacingBeforeLines());

        copyTo.setSpacingLineRule(copySource.getSpacingLineRule());

        copyTo.setVerticalAlignment(copySource.getVerticalAlignment());

        copyTo.setPageBreak(copySource.isPageBreak());
        if (copySource.isWordWrapped())
            copyTo.setWordWrapped(copySource.isWordWrapped());

        if (copySource.getNumID() != null) {
            // copy序列样式
            XmlObject xml = copySource.getCTP().getPPr().copy();
            copyTo.setNumID(copySource.getNumID());
            copyTo.getCTP().getPPr().set(xml);
        }
        ////        copyTo.setStyle(copySource.getStyleID());
        //        copyTo.setStyle(copySource.getStyle());
    }

    /**
     * 表格单元格样式克隆
     * @author slzs
     * 2017年11月10日 上午10:56:44
     * @param copyCell
     * @param cell
     */
    public void styleClone(XWPFTableCell copySource, XWPFTableCell copyTo) {
        if (copySource.getVerticalAlignment() != null) {
            copyTo.setVerticalAlignment(copySource.getVerticalAlignment());
        }
    }
}
