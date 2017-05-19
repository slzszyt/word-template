package com.slzs.word.process;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import com.slzs.util.ObjectUtil;
import com.slzs.word.model.Status;
import com.slzs.word.model.WordData;

/**
 * 数据填充分析器
 * @author 北京拓尔思信息技术股份有限公司
 * @author slzs
 * 2017年4月14日 下午1:17:43
 */
public class DataFillAnalyzer {
    Map<String, Integer>      iteratorIndexMap;
    public WordData           nextIterator;     // 待迭代数据
    private List<WordData>    analyzedList;     // 分析过的数据
    private TableAnalyzer     tableAnalyzer;
    private IteratorAnalyzer  iteratorAnalyzer;
    private ParagraphAnalyzer paragraphAnalyzer;

    public DataFillAnalyzer(StyleAnalyzer styleAnalyzer) {
        iteratorIndexMap = new HashMap<String, Integer>();
        iteratorAnalyzer = new IteratorAnalyzer(this,styleAnalyzer);
        paragraphAnalyzer = new ParagraphAnalyzer(this, iteratorAnalyzer);
        tableAnalyzer = new TableAnalyzer(this,paragraphAnalyzer,styleAnalyzer);
    }
    
    /**
     * 合并后移除数据
     * @author: slzs 
     * 2015-2-5 下午3:41:14
     * @param data
     * @param iteratorKey
     * @param wordDataList
     */
    void mergeMove(WordData data, String iteratorKey, List<WordData> wordDataList) {
        if (ObjectUtil.isNotEmpty(wordDataList)) {
            WordData wordData = wordDataList.get(0); // 当遇到迭代器next节点时移除0

            // 清空旧数据
            Map<String, String> dataTextMap = data.getTextFieldMap();
            Map<String, Object> dataImageMap = data.getImageFieldMap();
            removeByKeyPrefix(iteratorKey, dataTextMap); // 先清空旧数据
            removeByKeyPrefix(iteratorKey, dataImageMap); // 先清空旧数据
            if (wordData != null) {
                if (wordData.hasTextField()) {
                    Map<String, String> tempMap = wordData.getTextFieldMap();
                    if (ObjectUtil.isNotEmpty(dataTextMap)) {// 设置下一组文本
                        dataTextMap.putAll(tempMap);
                    } else {
                        data.setTextFieldMap(tempMap);
                    }
                }

                if (wordData.hasImageField()) { // 设置下一组图片
                    Map<String, Object> tempMap = wordData.getImageFieldMap();
                    if (ObjectUtil.isNotEmpty(dataImageMap)) {
                        dataImageMap.putAll(tempMap);
                    } else {
                        data.setImageFieldMap(tempMap);
                    }
                }

                if (wordData.hasIterator()) {
                    // 嵌套迭代数据处理
                    if (nextIterator == null) {
                        nextIterator = new WordData();
                    }
                    Map<String, List<WordData>> nextMap = wordData.getIteratorMap();
                    Integer index = iteratorIndexMap.get(iteratorKey);
                    if (index == null)
                        index = 0;
                    for (String key : nextMap.keySet()) {
                        nextIterator.addIterator(
                                iteratorKey + index + key.substring(key.indexOf(iteratorKey) + iteratorKey.length()),
                                nextMap.get(key));
                    }
                }

                if (wordData.hasTable()) {
                    throw new RuntimeException("当前版本尚未实现迭代表格数据方法……");
                }
            }

            wordDataList.remove(0); // 移除迭代数据
        }
    }

    /**
     * 依据key前缀删除map数据
     * @author: slzs
     * 2015-2-5 下午6:15:16
     * @param prefix
     * @param dataMap
     */
    private Map<String, ?> removeByKeyPrefix(String prefix, Map<String, ?> dataMap) {
        if (ObjectUtil.isNotEmpty(dataMap)) {
            Object[] keyArray = dataMap.keySet().toArray();
            for (Object key : keyArray) {
                if (key.toString().startsWith(prefix + ".")) {
                    dataMap.remove(key);
                }
            }
        }
        return dataMap;
    }
    
    /**
     * 数据标记分析
     * 
     * @author: slzs 
     * 2016-1-12 上午11:46:58
     * @param document 
     * @param data
     */
    public void analyData(XWPFDocument document, WordData data) {
        if (analyzedList == null) {
            analyzedList = new ArrayList<>();
        }
        if (!analyzedList.contains(data)) {
            // 解析表格标签
            tableAnalyzer.analy(document, data.getTableMap());
            if (data.hasTable()) {
                data = tableAnalyzer.tableAllNext(data);
            }

            if (data.hasIterator()) {
                // 迭代标签解析
                iteratorAnalyzer.analy(document, data.getIteratorMap());
                data = iteratorAnalyzer.iteratorAllNext(data);
            }
            analyzedList.add(data);
        }
    }

    /**
     * @see ParagraphAnalyzer
     * @author slzs 
     * 2017年4月26日 上午10:23:36
     * @param document2
     * @param paragraph
     * @param data
     * @return
     */
    public Status setParagraphContent(XWPFDocument document, XWPFParagraph paragraph, WordData data) {
        return paragraphAnalyzer.setParagraphContent(document, paragraph, data);
    }

    /**
     * @see TableAnalyzer
     * @author slzs 
     * 2017年4月26日 上午10:23:47
     * @param document2
     * @param table
     * @param data
     */
    public void setTableContent(XWPFDocument document, XWPFTable table, WordData data) {
        tableAnalyzer.setTableContent(document, table, data);
    }
}
