package com.slzs.word.process;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.xmlbeans.XmlCursor;

import com.slzs.util.ObjectUtil;
import com.slzs.util.StringUtil;
import com.slzs.word.model.Status;
import com.slzs.word.model.WordData;

/**
 * 迭代数据分析
 * @author slzs
 * 2017年4月14日 下午2:50:17
 */
class IteratorAnalyzer {
    private DataFillAnalyzer dfAnalyzer;
    private StyleAnalyzer styleAnalyzer;

    /**
     * @param dataFillAnalyzer
     */
    public IteratorAnalyzer(DataFillAnalyzer dfAnalyzer,StyleAnalyzer styleAnalyzer) {
        this.dfAnalyzer = dfAnalyzer;
        this.styleAnalyzer = styleAnalyzer;
    }

    /**
     * 迭代块解析
     * @author: slzs 
     * 2015-1-30 下午5:44:41
     * @param document
     * @param iteratorMap
     * @return 
     */
    Status analy(XWPFDocument document, Map<String, List<WordData>> iteratorMap) {
        Status res = Status.SUCCESS;
        for (String key : iteratorMap.keySet()) {
            List<WordData> iteratorDataList = iteratorMap.get(key); // 迭代器数据

            Map<IBodyElement, List<XWPFRun>> cloneRunMapList = getCloneElement(document, iteratorDataList, key); // 解析迭代需要克隆的内容

            res = copyCloneData(document, cloneRunMapList, iteratorDataList, key); // 复制克隆数据
        }
        return res;
    }
    
    /**
     * 获取文档中指定的迭代部分
     * @author: slzs
     * 2015-2-4 下午6:51:49
     * @param document
     * @param iteratorDataList
     * @param key 迭代标记${key:start}${key:end}
     * @return Map<IBodyElement,List<XWPFRun>> 段落value为run集合，表格value为null
     */
    private Map<IBodyElement, List<XWPFRun>> getCloneElement(XWPFDocument document, List<WordData> iteratorDataList, String key) {

        Map<IBodyElement, List<XWPFRun>> cloneRunMapList = new LinkedHashMap<IBodyElement, List<XWPFRun>>(); // 存储迭代克隆内容

        List<IBodyElement> bodyElementList = document.getBodyElements();

        /** 解析标签，保存克隆数据 **/
        boolean hasData = ObjectUtil.isNotEmpty(iteratorDataList); // true:包含数据
        String markStartKey = "${" + key.trim() + ":start}";// ${标记}字符串
        String markEndKey = "${" + key.trim() + ":end}";// ${标记}字符串
        boolean iteratorOpen = false;
        for (int bodyPos = 0; bodyPos < bodyElementList.size(); bodyPos++) {
            IBodyElement bodyElement = bodyElementList.get(bodyPos);
            if (bodyElement.getElementType() == BodyElementType.PARAGRAPH) {
                // 段落处理
                XWPFParagraph paragraph = (XWPFParagraph) bodyElement;
                List<XWPFRun> runList = paragraph.getRuns();
                boolean runListHasStartMark = runList.toString().contains(markStartKey);
                boolean runListHasEndMark = runList.toString().contains(markEndKey);
                if (iteratorOpen || runListHasStartMark) {
                    List<XWPFRun> cloneRunList = null;

                    if (runList.size() > 0) {
                        boolean runHasStartMark = false;
                        boolean runHasEndMark = false;

                        boolean clearBlank = false;
                        int runPos = 0, pass = 0;
                        for (runPos = 0; runPos < runList.size(); runPos++) {
                            XWPFRun run = runList.get(runPos);
                            runHasStartMark = run.toString().contains(markStartKey); // 包含开始标记
                            runHasEndMark = run.toString().contains(markEndKey); // 包含结束标记则停止
                            if (!iteratorOpen) {
                                iteratorOpen = runHasStartMark;
                            }

                            if (!hasData || runHasStartMark || runHasEndMark
                                    || ((runListHasStartMark || runListHasEndMark)
                                            && StringUtil.isEmpty(run.toString()))) {
                                // 没有数据移除start-end段，清理开始结束标签空白标记
                                if (paragraph.removeRun(runPos)) {
                                    runPos--;
                                } else {
                                    ParagraphAnalyzer.clearRun(run);
                                    pass++;
                                }
                                clearBlank = true;
                            } else if (iteratorOpen && hasData) {
                                // 存储复制
                                if (cloneRunMapList.containsKey(bodyElement)) {
                                    cloneRunList = cloneRunMapList.get(bodyElement);
                                }
                                if (cloneRunList == null) {
                                    cloneRunList = new ArrayList<XWPFRun>();
                                }
                                cloneRunList.add(run);
                                cloneRunMapList.put(bodyElement, cloneRunList);
                            }

                            if (runHasEndMark) {
                                iteratorOpen = false;
                            }
                        }

                        if (clearBlank && runPos - pass <= 0) {
                            // 移除空白行
                            document.removeBodyElement(bodyPos--);
                        }
                    } else if (iteratorOpen) {
                        // 存储复制
                        cloneRunMapList.put(bodyElement, null);
                    }
                }
            } else if (iteratorOpen) {
                if (!hasData) {
                    // 删除
                    document.removeBodyElement(bodyPos--);
                } else {
                    // 复制
                    cloneRunMapList.put(bodyElement, null);
                }
            }
        }

        return cloneRunMapList;
    }
    


    /**
     * 在迭代位置拷贝数据
     * @author: slzs
     * 2015-2-4 下午7:01:29
     * @param document
     * @param cloneRunMapList 需要复制的数据
     * @param iteratorDataList 迭代的数据
     * @param key
     * @return 
     */
    private Status copyCloneData(XWPFDocument document, Map<IBodyElement, List<XWPFRun>> cloneRunMapList,
            List<WordData> iteratorDataList, String key) {
        if (ObjectUtil.isEmpty(cloneRunMapList)) {
            return Status.SUCCESS;
        }

        /** 处理迭代数据，复制数据添加标记 **/
        Status res = Status.SUCCESS;
        // 最后添加的元素
        IBodyElement lastBodyElement = null;
        for (int i = 1; i < iteratorDataList.size(); i++) {
            IBodyElement[] keyArray = new IBodyElement[cloneRunMapList.keySet().size()];
            cloneRunMapList.keySet().toArray(keyArray);
            for (int c = 0; c < keyArray.length; c++) {
                IBodyElement bodyElement = keyArray[c];

                XmlCursor cursor = null;

                int bIndex = document.getBodyElements()
                        .lastIndexOf(lastBodyElement == null ? keyArray[keyArray.length - 1] : lastBodyElement);

                int pIndex = document.getParagraphPos(bIndex);
                int tIndex = document.getTablePos(bIndex);
                IBodyElement bTemp = null; // 临时位置，用于第一次定位，写入新位置后删除

                /** 在上次位置后位置追加 **/
                if (pIndex > -1) {
                    // 最后一次为段落
                    pIndex++;
                    if (pIndex >= document.getParagraphs().size())
                        bTemp = document.createParagraph(); // 以后面位置做参照向前追加，如果后面没有需要创建临时位置
                    cursor = document.getDocument().getBody().getPArray(pIndex).newCursor();
                } else if (tIndex > -1) {
                    // 最后一次为表格
                    tIndex++;
                    if (tIndex >= document.getTables().size())
                        bTemp = document.createTable(); // 以后面位置做参照向前追加，如果后面没有需要创建临时位置
                    cursor = document.getDocument().getBody().getTblArray(tIndex).newCursor();
                }

                if (c == 0) {
                    XWPFParagraph nodeParagraph = document.insertNewParagraph(cursor);
                    XWPFRun nodeRun = nodeParagraph.createRun(); // 标记循环节点
                    nodeRun.setText("${" + key + "#next}");
                    if (pIndex > -1) {
                        cursor = document.getParagraphArray(pIndex + 1).getCTP().newCursor();
                    } else {
                        cursor = document.getTableArray(tIndex).getCTTbl().newCursor();
                    }
                }

                if (bodyElement.getElementType() == BodyElementType.PARAGRAPH) {

                    XWPFParagraph newParagraph = document.insertNewParagraph(cursor);
                    styleAnalyzer.styleClone((XWPFParagraph) bodyElement, newParagraph);// 克隆段落样式

                    List<XWPFRun> cloneRunList = cloneRunMapList.get(bodyElement);
                    if (ObjectUtil.isNotEmpty(cloneRunList)) {// 写入段中内容
                        for (XWPFRun sourceRun : cloneRunList) {
                            XWPFRun newRun = newParagraph.createRun();
                            newRun.getCTR().set(sourceRun.getCTR().copy());
                        }
                    }

                    // 最后写入对象
                    lastBodyElement = newParagraph;
                } else if (bodyElement.getElementType() == BodyElementType.TABLE) {

                    XWPFTable newTable = document.insertNewTbl(cursor);
                    XWPFTable srcTable = ((XWPFTable) bodyElement);

                    newTable.getCTTbl().set(srcTable.getCTTbl().copy()); // 基于xml完全拷贝
                    
                    // 最后写入对象
                    lastBodyElement = newTable;
                    res = Status.REBUILD;
                }
                if (bTemp != null) {
                    // 删除临时标记段落
                    document.removeBodyElement(document.getBodyElements().indexOf(bTemp));
                }
            }

        }
    
        return res;
    }


    /**
     * 加载指定迭代器下一行数据
     * @author: slzs
     * 2015-2-5 下午3:28:40
     * @param data
     * @param iteratorKey
     * @return WordData
     */
    WordData iteratorNext(WordData data, String iteratorKey) {
        if (data.hasIterator()) {
            // 迭代器填充数据
            Map<String, List<WordData>> iteratorMap = data.getIteratorMap();
            List<WordData> wordDataList = iteratorMap.get(iteratorKey);
            dfAnalyzer.mergeMove(data, iteratorKey, wordDataList);
        }
        return data;
    }
    
    /**
     * 全部迭代器加载下一行迭代数据
     * @author: slzs
     * 2015-2-5 下午3:28:40
     * @param data
     * @return WordData
     */
    WordData iteratorAllNext(WordData data) {
        if (data.hasIterator()) {
            // 迭代器填充数据
            Map<String, List<WordData>> iteratorMap = data.getIteratorMap();
            Set<String> keySet = iteratorMap.keySet();
            for (String iteratorKey : keySet) {
                List<WordData> wordDataList = iteratorMap.get(iteratorKey);
                dfAnalyzer.mergeMove(data, iteratorKey, wordDataList);
            }
        }
        return data;
    }
}
