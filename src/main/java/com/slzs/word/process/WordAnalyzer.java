package com.slzs.word.process;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import com.slzs.util.StringUtil;
import com.slzs.word.WordFactory;
import com.slzs.word.model.Status;
import com.slzs.word.model.WordData;

import lombok.extern.log4j.Log4j2;

/**
 * 此处为word导出的具体实现,可通过WordFacory中的静态方法调用导出
 * @see WordFactory
 * @author slzs
 * @date 2019/04/17 13:08
 */
@Log4j2
public class WordAnalyzer {

    private DataFillAnalyzer dataFillAnalyzer;
    private MarkAnalyzer     markAnalyzer;
    private StyleAnalyzer    styleAnalyzer;

    // 重建位置，特殊情况下重新构建文档，记录下已经处理的数据位置
    private int              REBUILD_POS;

    private XWPFDocument     document;

    /**
     * 初始化各项分析器
     * @author slzs
     * 2017年4月25日 下午6:03:30
     */
    private void initAnalyzer() {
        styleAnalyzer = new StyleAnalyzer();
        dataFillAnalyzer = new DataFillAnalyzer(styleAnalyzer);
        markAnalyzer = new MarkAnalyzer(styleAnalyzer);
    }

    /**
     * 相同对象导出时串行处理，避免数据混乱。可通过创建多对象并发导出。
     * @author slzs
     * @date 2019/04/17 13:17
     * @param template
     * @param outFilePath
     * @param data
     * @return
     */
    public OutputStream report(Object template, String outFilePath, WordData data) {
        synchronized (this) { // 相同对象导出时串行处理，避免数据混乱。可通过创建多对象并发导出。
            boolean success = false;
            long startTime = System.currentTimeMillis();
            OutputStream outStream = null;
            try {

                document = loadTemplate(template);

                if (document == null) {
                    return null;
                }

                initAnalyzer();

                markAnalyzer.markComb(document); // 梳理模板标记

                log.info("开始填充段落标记数据....");

                while (data != null) {
                    Status res;
                    do {
                        res = setDataForMark(document, data); // 填充段落标记数据
                    } while (res == Status.REBUILD && rebuild());

                    data = dataFillAnalyzer.nextIterator; // 嵌套迭代数据处理
                    dataFillAnalyzer.nextIterator = null;
                }

                markAnalyzer.markClear(document); // 清理无用标记
                markAnalyzer.ubbAnalysis(document); // UBB代码分析

                if (StringUtil.isNotEmpty(outFilePath)) {
                    log.info("创建Word文件：" + outFilePath);
                    outStream = new FileOutputStream(outFilePath);
                } else {
                    log.info("创建Word ByteArrayOutputStream..");
                    outStream = new ByteArrayOutputStream();
                }

                document.write(outStream);

                success = true;
            } catch (Exception e) {
                log.error("", e);
            } finally {
                if (outStream != null)
                    try {
                        outStream.flush();
                        outStream.close();
                    } catch (Exception e) {
                        log.error("outStream关闭异常", e);
                    }
                if (document != null)
                    try {
                        document.close();
                    } catch (IOException e) {
                        log.error("document关闭异常", e);
                    }

                this.document = null;
                this.dataFillAnalyzer = null;
                this.markAnalyzer = null;
                this.REBUILD_POS = 0;

                log.info("word处理" + (success ? "成功" : "失败") + "...耗时："
                        + (System.currentTimeMillis() - startTime / 1000f) + "s");
            }

            return outStream;

        }
    }

    /**
     * @author slzs
     * 2017年5月9日 上午9:16:29
     * @param template
     * @param outFilePath
     * @throws IOException
     */
    private XWPFDocument loadTemplate(Object template) {

        try (OutputStream outStream = new ByteArrayOutputStream()) {
            log.info("加载Word模板....");
            if (template instanceof String) {
                // 路径
                document = new XWPFDocument(POIXMLDocument.openPackage(template.toString()));
            } else if (template instanceof InputStream) {
                // 流
                document = new XWPFDocument((InputStream) template);
            } else if (template instanceof File) {
                // 文件对象
                document = new XWPFDocument(new FileInputStream((File) template));
            }

            if (document != null) { // clone模板
                log.info("clone WordTmplate to ByteArrayOutputStream..");
                document.getPackage().save(outStream);
                document.getPackage().revert();
                document = null;
                InputStream inputStream = new ByteArrayInputStream(((ByteArrayOutputStream) outStream).toByteArray());
                document = new XWPFDocument(inputStream);
            }
        } catch (IOException e) {
            log.error("模板加载异常", e);
        }
        return document;
    }

    /**
     * 重建文档 直接修改dom会导致部分函数解析错误，可能是函数变量关联的一致性问题，重新加载可以解决这一问题
     * @author: slzs
     * 2016-1-11 上午11:00:01
     */
    private boolean rebuild() {
        boolean success = true;
        File tmpFile = null;
        FileOutputStream outStream = null;
        try {
            tmpFile = File.createTempFile("rpt_reload_", ".docx");

            log.debug("生成Word临时文件:" + tmpFile.getAbsolutePath());
            outStream = new FileOutputStream(tmpFile);
            document.write(outStream);

        } catch (Exception e) {
            log.error("模板重建异常", e);
            success = false;
        } finally {
            if (outStream != null)
                try {
                    outStream.flush();
                    outStream.close();
                } catch (Exception e) {
                    log.error("输出流关闭异常", e);
                }
            if (document != null)
                try {
                    document.close();
                } catch (IOException e) {
                    log.error("document关闭异常", e);
                }
        }

        if (success) {
            InputStream inputStream = null;
            try {
                inputStream = new FileInputStream(tmpFile);
                // 加载模板文档
                document = new XWPFDocument(inputStream);
            } catch (Exception e) {
                log.error("加载重建模板错误", e);
            } finally {
                try {
                    if (inputStream != null) {
                        inputStream.close();
                    }
                } catch (Exception e) {
                    log.error("inputStream.close faild", e);
                }
            }
        }

        if (tmpFile != null) {
            tmpFile.delete();
        }

        if (success) {
            markAnalyzer.markComb(document); // 梳理模板标记
        }
        return true;
    }

    /**
     * 标记变量填充数据，可完整保留标记样式
     * @author: slzs
     * 2014-4-16 下午2:16:05
     * @param document
     * @param data 填充数据
     * @return
     */
    private Status setDataForMark(XWPFDocument document, WordData data) {
        Status res = Status.SUCCESS;

        res = dataFillAnalyzer.analyData(document, data);

        List<IBodyElement> bodyElementList = document.getBodyElements();
        for (int bodyPos = REBUILD_POS; res == Status.SUCCESS && bodyPos < bodyElementList.size(); bodyPos++) {
            IBodyElement bodyElement = bodyElementList.get(bodyPos);
            if (bodyElement.getElementType() == BodyElementType.PARAGRAPH) {
                // 段落处理
                XWPFParagraph paragraph = (XWPFParagraph) bodyElement;
                res = dataFillAnalyzer.setParagraphContent(paragraph, data);
            } else if (bodyElement.getElementType() == BodyElementType.TABLE) {
                // 表格处理
                XWPFTable table = (XWPFTable) bodyElement;
                dataFillAnalyzer.setTableContent(table, data);
            }
            if (res == Status.REBUILD) {
                REBUILD_POS = bodyPos;
            }
        }
        return res;
    }
}
