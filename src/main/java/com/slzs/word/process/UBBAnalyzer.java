package com.slzs.word.process;

import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;

import com.slzs.util.StringUtil;

/**
 * UBB代码处理器
 * @author 北京拓尔思信息技术股份有限公司
 * @author slzs
 * 2017年5月9日 下午1:57:14
 */
 class UBBAnalyzer {

    private final String[] TAG_BOLD      = { ".*?(\\[b\\]).*", "[/b]" };
    private final String[] TAG_ITALIC    = { ".*?(\\[i\\]).*", "[/i]" };
    private final String[] TAG_UNDERLINE = { ".*?(\\[u\\]).*", "[/u]" };
    private final String[] TAG_STRIKE    = { ".*?(\\[strike\\]).*", "[/strike]" };
    private final String[] TAG_STRIKES   = { ".*?(\\[strikes\\]).*", "[/strikes]" };
    private final String[] TAG_COLOR     = { ".*?(\\[color=([0-9A-Fa-f]{6})\\]).*", "[/color]" };
    private final String[] TAG_FONT      = { ".*?(\\[font=([a-zA-Z\u4e00-\u9fa5 ]+)\\]).*", "[/font]" };
    private final String[] TAG_SIZE      = { ".*?(\\[size=([0-9]{2})\\]).*", "[/size]" };
    
    private StyleAnalyzer styleAnalyzer;
    
    UBBAnalyzer(StyleAnalyzer styleAnalyzer){
        this.styleAnalyzer = styleAnalyzer;
    }
    
    /**
     * 分析文档
     * @author slzs
     * 2017年5月9日 下午2:12:08
     * @param document
     */
    public void analysis(XWPFDocument document) {
        List<IBodyElement> bodys = document.getBodyElements();
        analysis(bodys);
    }

    /**
     * 分析body类型
     * @author slzs 
     * 2017年5月9日 下午2:33:34
     * @param bodyElements
     */
    private void analysis(List<IBodyElement> bodyElements) {
        for (IBodyElement element : bodyElements) {
            if (element instanceof XWPFParagraph) { // 段落
                XWPFParagraph paragraph = (XWPFParagraph) element;
                if (paragraph.getRuns() != null && !paragraph.getRuns().isEmpty() && paragraph.getText().matches(".*\\[.*\\].*\\[.*\\].*")) {
                    analysis(paragraph);
                }
            } else if (element instanceof XWPFTable) { // 表格
                analysis((XWPFTable) element);
            }
        }
    }
    
    
    /**
     * 分析表格类型
     * @author slzs 
     * 2017年5月9日 下午2:33:52
     * @param table
     */
    private void analysis(XWPFTable table) {
        List<IBodyElement> bodyElements;
        List<XWPFTableRow> rowList = table.getRows();
        List<XWPFTableCell> cellList;
        for (XWPFTableRow row : rowList) { // 行
            cellList = row.getTableCells();
            for (XWPFTableCell cell : cellList) { // 单元格
                bodyElements = cell.getBodyElements();
                analysis(bodyElements);
            }
        }
    }

    /**
     * 分析段落
     * @author slzs 
     * 2017年5月9日 下午2:33:58
     * @param paragraph
     */
    private void analysis(XWPFParagraph paragraph) {
        bold(paragraph);
        italic(paragraph);
        underline(paragraph);
        strike(paragraph);
        strikes(paragraph);
        size(paragraph);
        color(paragraph);
        font(paragraph);
    }
    
    /**
     * [color=字色]字色[/font]
     * @author slzs 
     * 2017年5月12日 上午9:02:35
     * @param paragraph
     */
    private void color(XWPFParagraph paragraph) {
        MatchHandler handler = new MatchHandler() {
            @Override
            public boolean ignore(XWPFRun run, String value) {
                return value != null && value.equals(run.getColor());
            }
            
            @Override
            public void dispose(XWPFRun run,String value) {
                run.setColor(value);
            }
        };
        analysis(paragraph, TAG_COLOR[0], TAG_COLOR[1], handler);
    }

    /**
     * [font=黑体]字体[/font]
     * @author slzs 
     * 2017年5月11日 上午11:19:50
     * @param paragraph
     */
    private void font(XWPFParagraph paragraph) {
        MatchHandler handler = new MatchHandler() {
            @Override
            public boolean ignore(XWPFRun run,String value) {
                return value!=null && value.trim().equalsIgnoreCase(run.getFontFamily());
            }
            
            @Override
            public void dispose(XWPFRun run,String value) {
                CTRPr pr = run.getCTR().addNewRPr();
                CTFonts fonts = pr.addNewRFonts();

                fonts.setAscii(value);
                fonts.setCs(value);
                fonts.setEastAsia(value);
                fonts.setHAnsi(value);
            }
        };
        analysis(paragraph, TAG_FONT[0], TAG_FONT[1], handler);
    }

    /**
     * [size=([0-9]{2})]字号[/size]
     * @author slzs 
     * 2017年5月11日 上午10:11:00
     * @param paragraph
     */
    private void size(XWPFParagraph paragraph) {
        MatchHandler handler = new MatchHandler() {
            @Override
            public boolean ignore(XWPFRun run,String value) {
                return (run.getFontSize()+"").equals(value);
            }
            
            @Override
            public void dispose(XWPFRun run,String value) {
                int size = Integer.parseInt(value.trim());
                run.setFontSize(size);
            }
        };
        analysis(paragraph, TAG_SIZE[0], TAG_SIZE[1], handler);
    }


    /**
     * [u]下划线[/u]
     * @author slzs 
     * 2017年5月11日 上午9:07:29
     * @param paragraph
     */
    private void underline(XWPFParagraph paragraph) {
        MatchHandler handler = new MatchHandler() {
            @Override
            public boolean ignore(XWPFRun run,String value) {
                return run.getUnderline()!=null && run.getUnderline() == UnderlinePatterns.SINGLE;
            }
            
            @Override
            public void dispose(XWPFRun run,String value) {
                run.setUnderline(UnderlinePatterns.SINGLE);
            }
        };
        analysis(paragraph, TAG_UNDERLINE[0], TAG_UNDERLINE[1], handler);
    }
    
    /**
     * [strike]删除线[/strike]
     * @author slzs 
     * 2017年5月12日 上午9:42:51
     * @param paragraph
     */
    private void strike(XWPFParagraph paragraph) {
        MatchHandler handler = new MatchHandler() {
            @Override
            public boolean ignore(XWPFRun run,String value) {
                return run.isStrikeThrough();
            }
            
            @Override
            public void dispose(XWPFRun run,String value) {
                run.setStrikeThrough(true);
            }
        };
        analysis(paragraph, TAG_STRIKE[0], TAG_STRIKE[1], handler);
    }

    
    /**
     * [strikes]双删除线[/strikes]
     * @author slzs 
     * 2017年5月12日 上午9:42:51
     * @param paragraph
     */
    private void strikes(XWPFParagraph paragraph) {
        MatchHandler handler = new MatchHandler() {
            @Override
            public boolean ignore(XWPFRun run,String value) {
                return run.isDoubleStrikeThrough();
            }
            
            @Override
            public void dispose(XWPFRun run,String value) {
                run.setDoubleStrikethrough(true);
            }
        };
        analysis(paragraph, TAG_STRIKES[0], TAG_STRIKES[1], handler);
    }

    /**
     * [i]斜体[/i]
     * @author slzs 
     * 2017年5月11日 上午9:07:27
     * @param paragraph
     */
    private void italic(XWPFParagraph paragraph) {
        MatchHandler handler = new MatchHandler() {
            @Override
            public boolean ignore(XWPFRun run,String value) {
                return run.isItalic();
            }
            
            @Override
            public void dispose(XWPFRun run,String value) {
                run.setItalic(true);
            }
        };
        analysis(paragraph, TAG_ITALIC[0], TAG_ITALIC[1], handler);
    }

    /**
     * [b]粗体[/b]
     * @author slzs 
     * 2017年5月9日 下午2:45:09
     * @param run
     */
    private void bold(XWPFParagraph paragraph) {
        MatchHandler handler = new MatchHandler() {
            @Override
            public boolean ignore(XWPFRun run,String value) {
                return run.isBold();
            }
            
            @Override
            public void dispose(XWPFRun run,String value) {
                run.setBold(true);
            }
        };
        analysis(paragraph, TAG_BOLD[0], TAG_BOLD[1], handler);
    }
    
    interface MatchHandler{  
         boolean ignore(XWPFRun run,String value); // 忽略部分run
         void dispose(XWPFRun run,String value);  // 处理匹配的run
    }
    
    private void analysis(XWPFParagraph paragraph, String beginTagReg, String endTag,MatchHandler matchHandler){
        String pText = paragraph.getText();
        if (!pText.matches(beginTagReg)) {
            return;
        }
        
        Pattern pattern = Pattern.compile(beginTagReg);
        Matcher tagM = null;
        
        boolean tagOpen = false;
        String beginTag = null; // open tag匹配的当前完整开始标记
        String tagValue = null; // opentag属性值
        List<XWPFRun> runList = paragraph.getRuns();
        
        XWPFRun run;
        for (int i = 0 ; i < runList.size() ; i++ ) {
            run = runList.get(i);
            String text = run.text();
            if (StringUtil.isNotEmpty(text)
                    && (tagOpen || (tagM = pattern.matcher(text)).matches() )) {
                
                if (!tagOpen) {
                    beginTag = tagM.group(1);
                    if (tagM.groupCount() > 1)
                        tagValue = tagM.group(2); // 属性值
                }
                
                int addNum = 0;
                if(!matchHandler.ignore(run,tagValue)){
                    boolean tagKeep = tagOpen;
                    int beginIx = tagKeep ? 0 : text.indexOf(beginTag);
                    int endIx = text.indexOf(endTag);
                    
                    if (tagOpen || beginIx > -1) {
                        tagOpen = true;
    
                        String frontStr = beginIx > 0 ? text.substring(0, beginIx) : "";
                        
                        int matchStartIx;
                        if (tagKeep) { // 继续last加粗标记
                            matchStartIx = 0;
                            if (endIx > -1){
                                endIx += matchStartIx;
                                tagOpen = false; // tag结束
                            } else {
                                endIx = text.length();
                            }
                        } else { 
                            matchStartIx = beginIx + beginTag.length();
                            if (endIx > -1) {
                                if (beginIx > endIx) {
                                    endIx = text.substring(matchStartIx).indexOf(endTag);
                                    if (endIx > -1){
                                        endIx += matchStartIx;
                                        tagOpen = false; // tag结束
                                    } else {
                                        endIx = text.length();
                                    }
                                } else {
                                    tagOpen = false; // tag结束
                                }
                            } else {
                                endIx = text.length();
                            }
                        }
                        
                        String matchStr = text.substring(matchStartIx, endIx);
                        String behindStr = tagOpen ? text.substring(endIx) : text.substring(endIx + endTag.length());
                        
                        if (frontStr != null && !"".equals(frontStr)) {
                            // 前部分
                            XWPFRun frontRun = paragraph.insertNewRun(i + ++addNum);
                            styleAnalyzer.styleClone(run, frontRun); // 原样式copy
                            frontRun.setText(frontStr.replace(beginTag, ""));
                        }
        
                        if (matchStr != null && !"".equals(matchStr)) {
                            // 匹配部分
                            XWPFRun matchRun = paragraph.insertNewRun(i + ++addNum);
                            styleAnalyzer.styleClone(run, matchRun); // 原样式copy
                            matchRun.setText(matchStr.replace(beginTag, ""));
                            matchHandler.dispose(matchRun,tagValue);
                        }
                        
                        if (behindStr != null && !"".equals(behindStr)) {
                            // 剩余文本
                            XWPFRun behindRun = paragraph.insertNewRun(i + addNum +1); //剩余部分仍需再次匹配addNum不増
                            styleAnalyzer.styleClone(run, behindRun); // 原样式copy
                            behindRun.setText(behindStr);
                        }
                        
                        if(addNum > 0){
                            paragraph.removeRun(i--);
                            i += addNum;
                        }
                    } 
                }
                if (addNum == 0)
                    run.setText(text.replace(beginTag, "").replace(endTag, ""),0);
            }
        }
    
    }
}