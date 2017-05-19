package com.slzs.word;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import com.slzs.util.ObjectUtil;
import com.slzs.util.StringUtil;

import lombok.extern.log4j.Log4j2;

/**
 * 模板生成工具
 * @author 北京拓尔思信息技术股份有限公司
 * @author slzs
 * 2017年4月26日 下午5:44:26
 */
@Log4j2
public class WordTemplateFactory {

    private static WordTemplateFactory templateFactory;
    private XWPFDocument               document;
    
    private WordTemplateFactory(){}

    /**
     * 获取实例
     * @author: slzs 
     * 2015-1-22 下午4:08:14
     * @return WordFactory
     */
    public synchronized static WordTemplateFactory getInstance() {
        if (templateFactory == null) {
            templateFactory = new WordTemplateFactory();
        }
        return templateFactory;
    }
    
    /**
     * 读取多个模板合并输出指定位置
     * @author slzs 
     * 2017年4月27日 上午9:51:35
     * @param outFilePath 指定输出文件路径
     * @param templateInputStreams 模板文件输入流,多个
     * @return true 成功 false 失败
     */
    public boolean mergeDocument(java.lang.String outFilePath, java.io.InputStream ... templateInputStreams){
        return mergeAll(outFilePath, templateInputStreams) != null;
    }

    /**
     * 读取多个模板合并输出指定位置
     * @author slzs 
     * 2017年4月27日 上午9:51:35
     * @param outFilePath 指定输出文件路径
     * @param templateFilePaths 模板文件位置，多个
     * @return true 成功 false 失败
     */
    public boolean mergeDocument(java.lang.String outFilePath,java.lang.String [] templateFilePaths){
        return mergeAll(outFilePath, templateFilePaths) != null;
    }

    /**
     * 合并多个文档流返回文件流
     * @author slzs 
     * 2017年4月27日 上午9:51:35
     * @param templateInputStreams 模板文件流，多个
     * @return 合并后的文件流
     */
    public java.io.OutputStream mergeDocument(java.io.InputStream ... templateInputStreams){
        return mergeAll(null, templateInputStreams);
    }

    /**
     * 读取多个模板合并返回文件流
     * @author slzs 
     * 2017年4月27日 上午9:51:35
     * @param templateFilePaths 模板文件位置，多个
     * @return 合并后的文件流
     */
    public java.io.OutputStream mergeDocument(java.lang.String ... templateFilePaths){
        return mergeAll(null, templateFilePaths);
    }
    

    private java.io.OutputStream mergeAll(String outFilePath , Object[] templates){
        
        if (ObjectUtil.isEmpty(templates)) {
            return null;
        }
        
        OutputStream outStream = null;
        try {
            for (Object template : templates) {
                mergeOne(template); // 合并每个文档
            }

            if (StringUtil.isNotEmpty(outFilePath)) {
                outStream = new FileOutputStream(outFilePath);
            } else {
                outStream = new ByteArrayOutputStream();
            }
            
            document.write(outStream);
        } catch (IOException e) {
            log.error("template merge error:", e);
        } finally {
            if (outStream != null)
                try {
                    outStream.flush();
                    outStream.close();
                } catch (IOException e) {
                    log.error("template flush error:", e);
                }
            if (document != null)
                try {
                    document.close();
                } catch (IOException e) {
                    log.error("template close error:", e);
                }
            
            this.document = null;
        }
        
        return outStream;
    }
    
    private void mergeOne(Object template) throws IOException {
        XWPFDocument documentTmp = null;
        if (template instanceof String) {
            // 路径
            documentTmp = new XWPFDocument(POIXMLDocument.openPackage(template.toString()));
        } else if (template instanceof InputStream) {
            // 流
            documentTmp = new XWPFDocument((InputStream) template);
        } else if (template instanceof File) {
            // 文件对象
            documentTmp = new XWPFDocument(new FileInputStream((File) template));
        }
        if (document == null) { // 创建新模板
            
            log.info("创建Word ByteArrayOutputStream..");
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            documentTmp.getPackage().save(bos);
            
            InputStream inputStream = new ByteArrayInputStream(bos.toByteArray());
            document = new XWPFDocument(inputStream);
            
        } else { // 合并元素
            List<IBodyElement> bodys = documentTmp.getBodyElements();
            for (IBodyElement body : bodys) {
                if(body instanceof XWPFParagraph){
                    XWPFParagraph p = (XWPFParagraph) body;
                    document.createParagraph().getCTP().set(p.getCTP().copy());          
                } else if(body instanceof XWPFTable ){
                    XWPFTable t = (XWPFTable) body;
                    document.createTable().getCTTbl().set(t.getCTTbl().copy());
                }
            }
        }
        documentTmp.close();
        documentTmp = null;
    }
    
}
