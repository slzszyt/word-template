package com.slzs.word.model;

import java.io.File;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.slzs.util.ObjectUtil;

/**
 * word报告数据
 * @author slzs
 * 2015-1-22 上午11:42:00
 * each engineer has a duty to keep the code elegant
 */
public class WordData {

    private Map<String, String>         textFieldMap; // 文本字段

    private Map<String, Object>         imageFieldMap; // 图片字段

    private Map<String, List<WordData>> iteratorMap;  // 迭代器

    private Map<String, List<WordData>> tableMap;     // 表格数据

    public WordData() {
    }

    /**
     * 添加文本字段
     * @author: slzs
     * 2015-1-22 上午11:48:16
     * @param key ${key}
     * @param text 显示文本
     */
    public void addTextField(String key, String text) {
        if (textFieldMap == null)
            textFieldMap = new HashMap<String, String>();
        textFieldMap.put(key, text);
    }

    /**
     * 添加带链接字段
     * @author: slzs
     * 2015-1-22 上午11:49:07
     * @param key ${key}
     * @param text 显示文本
     * @param link 链接地址
     */
    public void addLinkField(String key, String text, String link) {
        addTextField(key, text);
        addTextField(key + "#link", link);
    }

    /**
     * 添加图片字段
     * @author: slzs
     * 2015-1-22 下午1:51:12
     * @param key ${key}
     * @param imageFile 图片
     * @param width 固定宽度,null:自动
     * @param height 固定高度,null:自动
     */
    public void addImageField(String key, File imageFile, Integer width, Integer height) {
        addImageObject(key, imageFile, width, height);
    }


    /**
     * 添加图片字段
     * @author: slzs
     * 2015-1-22 下午1:51:12
     * @param key ${key}
     * @param InputStream 图片流
     * @param width 固定宽度,null:自动
     * @param height 固定高度,null:自动
     */
    public void addImageField(String key, InputStream imageStream, Integer width, Integer height) {
        addImageObject(key, imageStream, width, height);
    }
    
    private void addImageObject(String key, Object image, Integer width, Integer height){
        if (imageFieldMap == null)
            imageFieldMap = new HashMap<String, Object>();
        imageFieldMap.put(key, image);
        imageFieldMap.put(key + "#width", width);
        imageFieldMap.put(key + "#height", height);
    }

    /**
     * 添加迭代数据
     * @author: slzs
     * 2015-1-22 下午2:02:54
     * @param key 唯一键,重复会覆盖
     * @param dataList 数据集
     */
    public void addIterator(String key, List<WordData> dataList) {
        if (iteratorMap == null)
            iteratorMap = new HashMap<String, List<WordData>>();
        for (WordData wordData : dataList) {
            addKeyPrefix(key, wordData.getTextFieldMap());
            addKeyPrefix(key, wordData.getImageFieldMap());
            addKeyPrefix(key, wordData.getTableMap());
            addKeyPrefix(key, wordData.getIteratorMap());
        }
        iteratorMap.put(key, dataList);
    }

    /**
     * 给key表格添加列数据，重复columnIndex列新增一行数据
     * @author: slzs
     * 2015-1-22 上午11:48:16
     * @param key 第一个参照行的第一个单元格标记${key:rows}
     * @param dataList 数据 没条记录为一行 
     */
    public void addTable(String key, List<WordData> dataList) {
        if (tableMap == null)
            tableMap = new HashMap<String, List<WordData>>();

        for (WordData wordData : dataList) {
            addKeyPrefix(key, wordData.getTextFieldMap());
            addKeyPrefix(key, wordData.getImageFieldMap());
            addKeyPrefix(key, wordData.getTableMap());
            addKeyPrefix(key, wordData.getIteratorMap());
        }

        tableMap.put(key, dataList);
    }

    @SuppressWarnings("unchecked")
    private void addKeyPrefix(String prefix, Map<String, ?> dataMap) {
        if (ObjectUtil.isNotEmpty(dataMap)) {
            int index = prefix.indexOf(".");
            if (index > 0) {
                prefix = prefix.substring(0, index);
            }
            Object[] keyArray = dataMap.keySet().toArray();
            for (Object key : keyArray) {
                ((Map<String, Object>) dataMap).put(prefix + "." + key, (Object) dataMap.get(key));
                dataMap.remove(key);
            }
        }
    }

    /**
     * 包含图片数据
     * @author: slzs
     * 2015-2-5 下午2:38:14
     * @return boolean
     */
    public boolean hasImageField() {
        return ObjectUtil.isNotEmpty(imageFieldMap);
    }

    /**
     * 包含文本数据
     * @author: slzs
     * 2015-2-5 下午2:38:00
     * @return boolean
     */
    public boolean hasTextField() {
        return ObjectUtil.isNotEmpty(textFieldMap);
    }

    /**
     * 包含迭代数据
     * @author: slzs
     * 2015-1-22 下午5:29:03
     * @return boolean
     */
    public boolean hasIterator() {
        return ObjectUtil.isNotEmpty(iteratorMap);
    }

    /**
     * 包含表格
     * @author: slzs
     * 2015-1-22 下午5:29:47
     * @return boolean
     */
    public boolean hasTable() {
        return ObjectUtil.isNotEmpty(tableMap);
    }

    public Map<String, String> getTextFieldMap() {
        return textFieldMap;
    }

    public Map<String, Object> getImageFieldMap() {
        return imageFieldMap;
    }

    public Map<String, List<WordData>> getIteratorMap() {
        return iteratorMap;
    }

    public Map<String, List<WordData>> getTableMap() {
        return tableMap;
    }

    public void setTextFieldMap(Map<String, String> textFieldMap) {
        this.textFieldMap = textFieldMap;
    }

    public void setImageFieldMap(Map<String, Object> imageFieldMap) {
        this.imageFieldMap = imageFieldMap;
    }

}
