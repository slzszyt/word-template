package com.slzs.word.process;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.w3c.dom.Node;

import com.slzs.util.ObjectUtil;
import com.slzs.word.model.DataFillType;
import com.slzs.word.model.Status;
import com.slzs.word.model.WordData;

import lombok.extern.log4j.Log4j2;

/**
 * 段落内容处理
 * @author 北京拓尔思信息技术股份有限公司
 * @author slzs
 * 2017年4月26日 上午9:36:37
 */
@Log4j2
class ParagraphAnalyzer {

    private DataFillAnalyzer dfAnalyzer;
    private IteratorAnalyzer iteratorAnalyzer;

    public ParagraphAnalyzer(DataFillAnalyzer dfAnalyzer, IteratorAnalyzer iteratorAnalyzer) {
        this.dfAnalyzer = dfAnalyzer;
        this.iteratorAnalyzer = iteratorAnalyzer;
    }

    
    /**
     * 段落数据处理
     * @author: slzs
     * 2015-2-5 下午1:50:44
     * @param document 文档对象
     * @param WordData 填充数据
     */
    Status setParagraphContent(POIXMLDocumentPart document, XWPFParagraph paragraph, WordData data) {
        Status res = Status.SUCCESS;
        if (paragraph.getText().matches(".*\\$\\{.*\\}.*")) {
            Map<String, String> textFieldMap = data.getTextFieldMap(); // 文本
            Map<String, Object> imageFieldMap = data.getImageFieldMap(); // 图片

            if (ObjectUtil.isNotEmpty(textFieldMap)) {
                Map<String, Integer> nextKeyMap = new HashMap<String, Integer>();
                // 设置文本数据
                res = setFieldMapData(textFieldMap, paragraph, nextKeyMap, DataFillType.TEXT_DATA);
                if (res == Status.SUCCESS) {
                    // 设置图片数据
                    res = setFieldMapData(imageFieldMap, paragraph, nextKeyMap,
                            DataFillType.IMAGE_DATA);

                    if (ObjectUtil.isNotEmpty(nextKeyMap)) {
                        for (String key : nextKeyMap.keySet()) {
                            iteratorAnalyzer.iteratorNext(data, key); // 下一行迭代数据
                        }
                    }
                }
            }
        }
        return res;
    }
    

    /**
     * word插入图片
     * @author: slzs
     * 2014-4-10 上午10:50:09
     * @param document 操作对象
     * @param Object 图片数据对象
     * @param markRun 插入位置
     * @param width 图片显示宽度
     * @param height 图片显示高度
     * @param type 图片格式
     * @throws InvalidFormatException
     * @throws FileNotFoundException 
     * @throws XmlException 
     * 
     */
    private void insertPicture(XWPFDocument document, Object fileObj, XWPFRun markRun, Integer width, Integer height, Integer type) {

        try {
            //      paragraph.setSpacingLineRule(LineSpacingRule.AUTO); // 行间距设置为自动，会根据图片高度自动调整
            InputStream inputStream;
            // 将各类型转换为文件流
            if (fileObj instanceof File) {
                inputStream = new FileInputStream((File) fileObj);// 文件
            } else if (fileObj instanceof InputStream) { // 文件流
                inputStream = (InputStream) fileObj;
            } else {
                inputStream = new FileInputStream(fileObj.toString());//文件地址
            }
            if (width == null) {
                width = 300;
            }
            if (height == null) {
                height = 200;
            }
            markRun.setText("", 0);
            markRun.addPicture(inputStream, Document.PICTURE_TYPE_JPEG, "slzs", Units.toEMU(width), Units.toEMU(height));
        } catch (Exception e) {
            log.error("图片写入异常:", e);
        }

    }
    

    /**
     * 设置数据
     * @author: slzs
     * 2015-2-9 下午4:22:42
     * @param fieldMap 字段数据
     * @param paragraph 当前替换段落
     * @param nextKeyMap 记录迭代数据标记位置
     * @param iteratorIndexMap 迭代索引数据 
     * @param dataType 数据类型
     */
    private Status setFieldMapData(Map<String, ?> fieldMap, XWPFParagraph paragraph, Map<String, Integer> nextKeyMap, DataFillType dataType) {
        if (ObjectUtil.isNotEmpty(fieldMap)) {
            for (String key : fieldMap.keySet()) {
                if (!key.contains("#")) { // 忽略属性键值
                    Object obj = fieldMap.get(key);
                    if (obj != null) {
                        List<XWPFRun> runList = paragraph.getRuns();
                        String content = runList.toString();
                        String markKey = "${" + key.trim() + "}";// ${标记}字符串
                        String nextMark = null;
                        String prefix = null;
                        Integer count = 0;
                        int tagIndex = key.lastIndexOf(".");
                        if (tagIndex > 0) {
                            prefix = key.trim().substring(0, tagIndex);
                            nextMark = "${" + prefix + "#next}";// ${prefix#next}字符串
                            count = dfAnalyzer.iteratorIndexMap.get(prefix);

                            if (count == null)
                                count = 0;

                            // 处理嵌套迭代标记处理
                            String iterRegex = "(.*\\$\\{" + prefix + ")(\\.[^\\.]*(\\.|\\:).*\\}.*)";
                            if (content.matches(iterRegex)) {
                                for (int p = 0; p < runList.size(); p++) {
                                    XWPFRun markRun = runList.get(p);
                                    if (markRun.toString().matches(iterRegex)) {
                                        markRun.setText(markRun.toString().replaceAll(iterRegex, "$1" + count + "$2"), 0);
                                    }
                                }
                            }
                        }

                        if (content.contains(markKey) || (nextMark != null && content.contains(nextMark))) {// 包含key标记
                            for (int p = 0; p < runList.size() && (!nextKeyMap.containsKey(prefix) || p < nextKeyMap.get(prefix)); p++) {
                                XWPFRun markRun = runList.get(p);
                                if (markRun.toString().contains(markKey)) {

                                    switch (dataType) {
                                        case TEXT_DATA: // 文本类型数据
                                            
                                            insertText((String) obj, markRun, markKey);

                                            if (fieldMap.containsKey(key + "#link")) { // 包含超链接                                                   
                                                // 设置超链接
                                                String relationId =
                                                        paragraph.getDocument().getPackagePart().addExternalRelationship((String) fieldMap.get(key + "#link"), XWPFRelation.HYPERLINK.getRelation())
                                                                .getId();
                                                if (markRun instanceof XWPFHyperlinkRun) { // 超链接格式
                                                    ((XWPFHyperlinkRun) markRun).setHyperlinkId(relationId); // 修改链接地址
                                                } else { // 非超链接格式转换超链接-由于poi本身不支持，需要处理后重新构建文档，增加读写效率略有影响。可在模板上直接加上超链接，可直接修改链接省略转换的过程。

                                                    CTP ctp = paragraph.getCTP();
                                                    CTHyperlink hyperlink = ctp.addNewHyperlink(); // 增加hyperlink，目前只支持末尾插入，再通过xml调整位置
                                                    hyperlink.setId(relationId); // 设置链接

                                                    XmlCursor cursorH = hyperlink.newCursor();

                                                    CTR ctr = markRun.getCTR();
                                                    XmlCursor cursorR = ctr.newCursor();
                                                    cursorH.moveXml(cursorR); // 移动链接到标记位置

                                                    cursorR.toPrevSibling(); // 选择hyperlink

                                                    Node node = cursorR.getDomNode(); // 获取hyperlink dom节点对象
                                                    node.appendChild(ctr.getDomNode()); // 将标记dom节点移到hyperlink下

                                                    return Status.REBUILD;
                                                }
                                            }

                                            break;
                                        case IMAGE_DATA:
                                            // 插入图片
                                            insertPicture(paragraph.getDocument(), obj, markRun, (Integer) fieldMap.get(key + "#width"), (Integer) fieldMap.get(key + "#height"),
                                                    (Integer) fieldMap.get(key + "#type"));
                                            break;
                                        default:
                                            break;
                                    }

                                } else if (!nextKeyMap.containsKey(prefix) && nextMark != null && markRun.toString().contains(nextMark)) {
                                    // 记录迭代结束位置
                                    nextKeyMap.put(prefix, p);
                                    dfAnalyzer.iteratorIndexMap.put(prefix, ++count);
                                }
                            }
                        }
                    }
                }
            }
        }

        return Status.SUCCESS;

    }


    /**
     * 设置文本
     * @author slzs 
     * 2017年4月26日 上午10:41:22
     */
    private void insertText(String text,XWPFRun markRun,String markKey) {
        String br = "\n";

        boolean next = false;
        do {
            boolean hasBr = text.contains(br);
            String textTmp;
            if (hasBr) // 换行前文本
                textTmp = text.substring(0, text.indexOf(br));
            else
                textTmp = text;

            if (markRun.toString().contains(markKey))
                markRun.setText(markRun.toString().replace(markKey, textTmp), 0);
            else
                markRun.setText(textTmp);

            if (next = hasBr) {
                markRun.addBreak();
                if (next = (text.indexOf(br) + 1 < text.length())) {
                    text = text.substring(text.indexOf(br) + 1);
                }
            }
        } while (next);
    }
}
