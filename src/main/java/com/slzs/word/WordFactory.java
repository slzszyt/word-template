package com.slzs.word;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import com.slzs.word.model.WordData;
import com.slzs.word.process.MergeAnalyzer;
import com.slzs.word.process.WordAnalyzer;

/**
 * WORD生成工具 <br>
 * 支持并发导出，但考虑io性能问题，应尽量控制并发数
 * @author slzs 
 * 2014-4-9 下午3:03:40 
 * each engineer has a duty to keep the code elegant
 */
public class WordFactory {


    private WordFactory() {
    }

    /**
     * 根据模板生成报告,模板支持${标记} “:”为特殊标记符，应避免使用 ${key:start}模板循环开始标记，独占一行
     * ${key:end}模板循环结束标记，独占一行
     * 
     * @author: slzs 
     * 2014-4-16 下午9:18:40
     * @param templateFilePath 模板文件路径
     * @param outFilePath 输出文件路径
     * @param data 报告数据
     * @return boolean 生成报告是否成功
     */
    public static boolean reportByTemplate(String templateFilePath, String outFilePath, WordData data) {
        return reportLaunch(templateFilePath, outFilePath, data) != null;
    }

    /**
     * 根据模板生成报告,模板支持${标记} “:”为特殊标记符，应避免使用 ${key:start}模板循环开始标记，独占一行
     * ${key:end}模板循环结束标记，独占一行
     * 
     * @author: slzs 
     * 2014-4-16 下午9:18:40
     * @param templateInputStream 模板数据流
     * @param outFilePath 输出文件路径
     * @param data 报告数据
     * @return boolean 生成报告是否成功
     */
    public static boolean reportByTemplate(InputStream templateInputStream, String outFilePath, WordData data) {
        return reportLaunch(templateInputStream, outFilePath, data) != null;
    }

    /**
     * 根据模板生成报告,模板支持${标记} “:”为特殊标记符，应避免使用 ${key:start}模板循环开始标记，独占一行
     * ${key:end}模板循环结束标记，独占一行
     * 
     * @author: slzs 
     * 2014-4-16 下午9:18:40
     * @param templateFilePath 模板文件路径
     * @param data 报告数据
     * @return ByteArrayOutputStream
     */
    public static OutputStream reportByTemplate(String templateFilePath, WordData data) {
        return reportLaunch(templateFilePath, null, data);
    }

    /**
     * 根据模板生成报告,模板支持${标记} “:”为特殊标记符，应避免使用 ${key:start}模板循环开始标记，独占一行
     * ${key:end}模板循环结束标记，独占一行
     * 
     * @author: slzs 
     * 2014-4-16 下午9:18:40
     * @param templateInputStream 文件流
     * @param data 报告数据
     * @return ByteArrayOutputStream
     */
    public static OutputStream reportByTemplate(InputStream templateInputStream, WordData data) {
        return reportLaunch(templateInputStream, null, data);
    }
    
    private static OutputStream reportLaunch(Object template, String outFilePath, WordData data) {
        return new WordAnalyzer().report(template, outFilePath, data);
    }
    

    /**
     * 读取多个模板合并输出指定位置
     * @author slzs 
     * 2017年4月27日 上午9:51:35
     * @param outFilePath 指定输出文件路径
     * @param templateInputStreams 模板文件输入流,多个
     * @return true 成功 false 失败
     * @throws IOException 
     */
    public static boolean mergeDocument(java.lang.String outFilePath, java.io.InputStream ... templateInputStreams) throws IOException{
        return mergeAll(outFilePath, templateInputStreams) != null;
    }

    /**
     * 读取多个模板合并输出指定位置
     * @author slzs 
     * 2017年4月27日 上午9:51:35
     * @param outFilePath 指定输出文件路径
     * @param templateFilePaths 模板文件位置，多个
     * @return true 成功 false 失败
     * @throws IOException 
     */
    public static boolean mergeDocument(java.lang.String outFilePath,java.lang.String [] templateFilePaths) throws IOException{
        return mergeAll(outFilePath, templateFilePaths) != null;
    }

    /**
     * 合并多个文档流返回文件流
     * @author slzs 
     * 2017年4月27日 上午9:51:35
     * @param templateInputStreams 模板文件流，多个
     * @return 合并后的文件流
     * @throws IOException 
     */
    public static OutputStream mergeDocument(InputStream ... templateInputStreams) throws IOException{
        return mergeAll(null, templateInputStreams);
    }

    /**
     * 读取多个模板合并返回文件流
     * @author slzs 
     * 2017年4月27日 上午9:51:35
     * @param templateFilePaths 模板文件位置，多个
     * @return 合并后的文件流
     * @throws IOException 
     */
    public static OutputStream mergeDocument(java.lang.String ... templateFilePaths) throws IOException{
        return mergeAll(null, templateFilePaths);
    }
    

    private static OutputStream mergeAll(String outFilePath , Object[] templates) throws IOException{
        return new MergeAnalyzer().mergeAll(outFilePath, templates);
    }
}
