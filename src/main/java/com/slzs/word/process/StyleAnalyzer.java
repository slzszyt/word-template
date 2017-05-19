package com.slzs.word.process;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;

/**
 * 样式处理
 * @author 北京拓尔思信息技术股份有限公司
 * @author slzs
 * 2017年4月14日 下午1:17:43
 */
public class StyleAnalyzer {

    /**
     * XWPFRun 样式克隆
     * @author: slzs
     * 2014-4-16 下午8:27:51
     * @param copySource 样式源
     * @param copyTo 目标
     */
    void styleClone(XWPFRun copySource, XWPFRun copyTo) {
        if (copySource != null && copyTo != null) {
            copyTo.setBold(copySource.isBold());
            copyTo.setColor(copySource.getColor());
            copyTo.setFontFamily(copySource.getFontFamily());
            if (copySource.getFontSize() > 0) {
                copyTo.setFontSize(copySource.getFontSize());
            }
            copyTo.setItalic(copySource.isItalic());
            copyTo.setStrikeThrough(copySource.isStrikeThrough());
            copyTo.setSubscript(copySource.getSubscript());
            copyTo.setUnderline(copySource.getUnderline());
            copyTo.setTextPosition(copySource.getTextPosition());

            CTRPr copyToRPR = copyTo.getCTR().getRPr();
            CTRPr copySourceRPR = copySource.getCTR().getRPr();
            if (copySourceRPR != null && copySourceRPR.isSetShd()) { // 设置样式
                CTShd copySourceshd = copySourceRPR.getShd();
                CTShd copyToshd = copyToRPR.addNewShd();
                copyToshd.set(copySourceshd.copy());
            }
        }
    }

    /**
     * XWPFParagraph 段落样式克隆
     * @author: slzs
     * 2015-1-29 下午3:57:38
     * @param copySource 样式源
     * @param copyTo 目标
     * 
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

        copyTo.setIndentationFirstLine(copySource.getIndentationFirstLine());
        copyTo.setIndentationHanging(copySource.getIndentationHanging());
        copyTo.setIndentationLeft(copySource.getIndentationLeft());
        copyTo.setIndentationRight(copySource.getIndentationRight());

        copyTo.setIndentFromLeft(copySource.getIndentFromLeft());
        copyTo.setIndentFromRight(copySource.getIndentFromRight());

        copyTo.setPageBreak(copySource.isPageBreak());

        copyTo.setSpacingAfter(copySource.getSpacingAfter());
//        copyTo.setSpacingAfterLines(copySource.getSpacingAfterLines());
        copyTo.setSpacingBefore(copySource.getSpacingBefore());
//        copyTo.setSpacingBeforeLines(copySource.getSpacingBeforeLines());
        copyTo.setSpacingLineRule(copySource.getSpacingLineRule());

        copyTo.setVerticalAlignment(copySource.getVerticalAlignment());

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
}
