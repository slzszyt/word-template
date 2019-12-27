package com.slzs.word.process;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyles;

import com.slzs.util.ObjectUtil;
import com.slzs.util.StringUtil;
import com.slzs.word.WordFactory;

import lombok.extern.log4j.Log4j2;

/**
 * 此处为word导出的具体实现,可通过WordFacory中的静态方法调用导出
 * @see WordFactory
 * @author slzs
 * @date 2019/04/17 13:08
 */
@Log4j2
public class MergeAnalyzer {

    /**
     * 多个word文档合并
     * @author slzs
     * @date 2019/12/27 13:18
     * @param outFilePath 存在输出路径时将合并的文件输出到目标路径
     * @param templates 要合并的文档，目前支持路径、文件流、文件对象三种
     * @return 返回合并后文件输出流
     * @throws IOException
     * @throws XmlException 
     */
    public OutputStream mergeAll(String outFilePath , Object[] templates) throws IOException {
        if (ObjectUtil.isEmpty(templates)) {
            return null;
        }
        XWPFDocument document = null;
        for (Object template : templates) {
            try {
                document = mergeOne(document, template); // 合并每个文档
            } catch (XmlException e) {
                log.error("文档合并异常", e);
            }
        }
        
        if(document == null)
            return null;

        try (OutputStream outStream = 
                StringUtil.isEmptyTrim(outFilePath) ? new ByteArrayOutputStream() // 无路径输出到内存
                : new FileOutputStream(outFilePath)) { // 有路径，输出到路径
            document.write(outStream);
            return outStream;
        } catch (IOException e) {
            log.error("template merge error:", e);
            throw e;
        } finally {
            document.close();
        }
    }
    
    /**
     * 合并文档，目标文档应为只读，避免发生错误
     * 根据模板传入方式解析document，目前支持路径、文件流、文件对象三种
     * 读取文档后，document对象不调用close，close会刷新源文件
     * @author slzs
     * @date 2019/12/27 13:13
     * @param document 基础文档，合并后返回该文档，如果未null则新创建文档
     * @param template 目前支持路径、文件流、文件对象三种
     * @return 返回合并后的基础文档
     * @throws IOException
     * @throws XmlException 
     */
    private XWPFDocument mergeOne(XWPFDocument document,Object template) throws IOException, XmlException {
        XWPFDocument documentTmp = initTmpDocument(template); // 读取文档信息，因为只读，不要调用close，否则会覆盖刷新源文件，多线程造成模板损坏
        if(documentTmp == null) {
            return null;
        }
        
        if (document == null) { // 创建新模板，第一次只有一个文件不需要合并
            log.info("创建Word ByteArrayOutputStream..");
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            documentTmp.getPackage().save(bos);
            documentTmp.getPackage().revert();
            InputStream inputStream = new ByteArrayInputStream(bos.toByteArray());
            document = new XWPFDocument(inputStream);
        } else { 
            /**
             * 文档的合并切分为：
             *   段落内容合并
             *   表格内容合并
             *   图片数据合并
             *   样式合并
             *   关联关系数据合并 
             *      - 由于合并后关联关系id会重复，所以需要特殊处理
             *      - 关联关系较为复杂，文字与超链接点关系、图片位置与图片数据的关系、序号关系等等，每种关联关系要单独处理，避免混乱
             * */

            
            /*
             *  首先进行关联关系合并，记录关系ID在原有文档与新文档生成ID的对应关系，在后续内容合并时进行ID替换
             *  如果后续有其它关联关系需要支持时，都应在此处添加，参考其它关系处理
             */
            Map<String,String> ridMap = null;
            log.info("合并图片...");
            ridMap = mergeImg(document, documentTmp,ridMap);
            log.info("合并超链接...");
            ridMap = mergeHyLink(document, documentTmp, ridMap);

            log.info("合并word...");
            List<IBodyElement> bodys = documentTmp.getBodyElements();
            for (IBodyElement body : bodys) {
                if (body instanceof XWPFParagraph) {
                    log.info("合并段落...");
                    XWPFParagraph p = (XWPFParagraph) body;
                    XmlObject xml = p.getCTP().copy();
                    document.createParagraph().getCTP().set(replaceRID(xml, ridMap));
                } else if (body instanceof XWPFTable) {
                    log.info("合并表格...");
                    XWPFTable t = (XWPFTable) body;
                    XmlObject xml = t.getCTTbl().copy();
                    document.createTable().getCTTbl().set(replaceRID(xml, ridMap));
                } else {
                    log.info("尚未支持的合并部分：" + body.getClass());
                }
            }
            
            log.info("合并样式...");
            try {
                CTStyles stsTmp = documentTmp.getStyle();
                if(stsTmp == null) {
                    return null;
                }
                XWPFStyles sts = document.createStyles();
                sts.setStyles(stsTmp);
            } catch (XmlException e) {
                log.error("样式合并异常。。。");
            }
        }
        return document;
    }

    /**
     * 超链接关系数据处理
     * @author slzs
     * @date 2019/12/27 15:24
     * @param document
     * @param documentTmp
     * @param ridMap
     * @return
     */
    private Map<String, String> mergeHyLink(XWPFDocument document, XWPFDocument documentTmp,
            Map<String, String> ridMap) {
        if (ridMap == null)
            ridMap = new HashMap<>();
        try {
            // 拿到超链接的关系数据，每种关联数据单独处理，避免不可预料的影响
            PackageRelationshipCollection relations = documentTmp.getPackagePart().getRelationshipsByType(XWPFRelation.HYPERLINK.getRelation());
            for (PackageRelationship rel : relations) {
                PackageRelationship r = document.getPackagePart().addRelationship(rel.getTargetURI(), rel.getTargetMode(), rel.getRelationshipType());
                if (rel.getId().equals(r.getId())) // 相同不更新
                    continue;
                ridMap.put(rel.getId(), r.getId());
            }
        } catch (InvalidFormatException e) {
            log.error("超链接合并异常", e);
        }
        
        return ridMap;
    }
    
    /**
     * 合并图片数据，记录图片新旧文档ID关系
     * @author slzs
     * @param documentTmp 
     * @param document 
     * @return ridMap
     * @date 2019/12/27 14:37
     */
    private Map<String, String> mergeImg(XWPFDocument document, XWPFDocument documentTmp, Map<String, String> ridMap) {
        if (ridMap == null)
            ridMap = new HashMap<>();
        List<XWPFPictureData> pics = documentTmp.getAllPictures();
        for (XWPFPictureData pic : pics) {
            try {
                String oldRId = documentTmp.getRelationId(pic);
                String newRId = document.addPictureData(pic.getData(), pic.getPictureType());
                if(oldRId.equals(newRId)) {
                    continue;
                }
                ridMap.put(oldRId, newRId);
            } catch (InvalidFormatException e) {
                log.error("图片合并异常", e);
            }
        }
        return ridMap;
    }

    /**
     * 新旧文档的关系ID处理
     * @author slzs
     * @date 2019/12/27 14:36
     * @param xml 合并前文档原始xml
     * @param ridMap key：旧文件ID，value：新文件ID
     * @return 返回替换ID合并后的xml
     * @throws XmlException 
     */
    private XmlObject replaceRID(XmlObject xmlObj, Map<String, String> ridMap) throws XmlException {
        String xmlTxt = xmlObj.xmlText();
        for (Entry<String, String> entry : ridMap.entrySet()) {
            xmlTxt = xmlTxt.replaceAll("(<[^>]+?\")" + entry.getKey() + "?(\"[^>]*?>)", "$1" + entry.getValue() + "$2");  
        }
        return XmlObject.Factory.parse(xmlTxt);
    }
    
    /**
     * 根据不同输入类型，解析document
     * 目前支持路径、文件流、文件对象三种
     * @author slzs
     * @date 2019/12/27 13:16
     * @param template 目前支持路径、文件流、文件对象三种
     * @return
     * @throws FileNotFoundException
     * @throws IOException
     */
    private XWPFDocument initTmpDocument(Object template) throws FileNotFoundException, IOException {
        if (template instanceof String)
            return new XWPFDocument(POIXMLDocument.openPackage(template.toString())); // 打开路径文件
        if (template instanceof InputStream)
            return new XWPFDocument((InputStream) template); // 读取流数据
        if (template instanceof File)
            return new XWPFDocument(new FileInputStream((File) template));// 读取文件对象
        return null; // 空数据
    }
}
