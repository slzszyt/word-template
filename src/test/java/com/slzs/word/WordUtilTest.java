package com.slzs.word;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import org.junit.Test;

import com.slzs.word.model.WordData;

/**
 * word模版导出测试
 * @author slzs
 * 2014-4-9 下午3:03:40
 * each engineer has a duty to keep the code elegant
 */
public class WordUtilTest {

    @Test
    public void reportWord() {

        String classPath = this.getClass().getResource("/").getPath();

        String srcPath = classPath + "report/week_template.docx";
        String destPath = classPath + "report/reportItem_word_"
                + Long.parseLong((Math.random() * 10000 + "").split("\\.")[0]) + ".docx";

        WordData data = new WordData();
        data.addTextField("organName", "新华社");
        data.addLinkField("organName", "新华社", "http://www.sina.com/");
        data.addTextField("createYear", "2015");
        data.addTextField("createDate", "12月31日");
        data.addTextField("title", "标题title标题标题标题标题标题标题标题");
        data.addTextField("totalCount", "8888888");
        data.addImageField("image", new File(classPath + "report/123.jpg"), null, null);
        data.addImageField("image2", new File(classPath + "report/123.gif"), 40, 50);
        {
            // 表格
            List<WordData> tableDataList = new ArrayList<WordData>();

            WordData rowData = new WordData();
            rowData.addTextField("title", "标题111");
            rowData.addTextField("name", "名称111");
            tableDataList.add(rowData);

            rowData = new WordData();
            rowData.addTextField("title", "标题2222");
            rowData.addTextField("name", "名称2222");
            tableDataList.add(rowData);

            data.addTable("weekTable", tableDataList);
        }
        { // 热点循环
            List<WordData> hotDataList = new ArrayList<WordData>();
            WordData hotData = new WordData();
            hotData.addTextField("title", "第一个热点");
            hotData.addTextField("pubDate", "2011-01-01");
            hotData.addTextField("source", "来自星星");
            hotData.addLinkField("link", "超级链接", "http://www.baidu.com/");
            hotData.addTextField("description", "说的是一只士大夫加快了daslfoiu维尔...");
            hotData.addImageField("image", new File(classPath + "report/123.gif"), 100, 100);
            hotDataList.add(hotData);

            hotData = new WordData();
            hotData.addTextField("title", "第二天2但当发生的");
            hotData.addTextField("pubDate", "2012-02-02");
            hotData.addTextField("source", "来自北京");
            hotData.addLinkField("link", "超级link", "http://www.163.com/");
            hotData.addTextField("description", "简单描述一下，\n啊\n事态严重...");
            hotData.addImageField("image", new File(classPath + "report/123.jpg"), 130, 70);
            hotDataList.add(hotData);

            hotData = new WordData();
            hotData.addTextField("title", "第三天只是标题了3");
            hotDataList.add(hotData);

            data.addIterator("hot", hotDataList);
        }

        {
            // 排行循环
            List<WordData> rankDataList = new ArrayList<WordData>();
            WordData rankData = new WordData();
            rankData.addTextField("title", "舆情热点rank!");
            rankDataList.add(rankData);

            rankData = new WordData();
            rankData.addTextField("title", "rank~~~~~~");

            { // 嵌套迭代
                List<WordData> childrenList = new ArrayList<WordData>();
                WordData doc = new WordData();
                doc.addTextField("title", "文档标题");
                childrenList.add(doc);
                doc = new WordData();
                doc.addTextField("title", "文档标题!!阿苏大夫");
                childrenList.add(doc);
                rankData.addIterator("doc", childrenList);
            }

            rankDataList.add(rankData);

            data.addIterator("rank", rankDataList);

        }

        WordFactory.getInstance().reportByTemplate(srcPath, destPath, data);

    }

    @Test
    public void testHot() {

        String classPath = this.getClass().getResource("/").getPath();

        String srcPath = classPath + "report/hotReport.dot";
        String destPath = classPath + "report/hotReport_" + (int) (Math.random() * 10000) + ".docx";

        WordData data = new WordData();
        data.addTextField("reportDate", "2929-20-01");

        List<WordData> dataList = new ArrayList<WordData>();

        WordData data2;
        { // 第一层数据
            data2 = new WordData();
            data2.addTextField("title", "一个热点事件,共享单车年!");

            List<WordData> dataCList = new ArrayList<WordData>();
            WordData data3;
            { // 第二层数据
                data3 = new WordData();
                data3.addTextField("title", "子标题-小黄车肿么了?");
                data3.addTextField("description", "子描述-小黄车肿了,座椅丢失,车胎被扒...");
                dataCList.add(data3);
            }
            {
                data3 = new WordData();
                data3.addTextField("title", "子标题2:摩拜单车的红包");
                data3.addTextField("description", "子描述2-膜拜摩拜,周末与摩拜一起玩捉迷藏");
                dataCList.add(data3);
            }
            {
                data3 = new WordData();
                data3.addTextField("title", "子标题3:小蓝车后起之秀");
                data3.addTextField("description", "子描述3-小蓝车与支付宝联合,芝麻信用免押金!");
                dataCList.add(data3);
            }

            data2.addIterator("doc", dataCList);

            dataList.add(data2);
        }

        {
            data2 = new WordData();
            data2.addTextField("title", "另一个热点事件,出轨出柜橱柜?");

            List<WordData> dataCList = new ArrayList<WordData>();
            WordData data3;
            {
                data3 = new WordData();
                data3.addTextField("title", "子标题-威廉王子爆出猛料");
                data3.addTextField("description", "子描述-新华社小道消息,威廉王子于西班牙男子");
                dataCList.add(data3);
            }
            {
                data3 = new WordData();
                data3.addTextField("title", "子标题2:白百合的故事");
                data3.addTextField("description", "子描述2-一朵白色的百合花,开了又谢了");
                dataCList.add(data3);
            }
            {
                data3 = new WordData();
                data3.addTextField("title", "子标题3:橱邦厨具");
                data3.addTextField("description", "子描述3-橱邦橱柜,买一送五!");
                dataCList.add(data3);
            }
            data2.addIterator("doc", dataCList);

            dataList.add(data2);
        }

        data.addIterator("hot", dataList);

        WordFactory.getInstance().reportByTemplate(srcPath, destPath, data);
    }

    @Test
    public void templateMerge() {
        String classPath = this.getClass().getResource("/").getPath();

        String template1 = classPath + "report/ttt1.docx";
        String template2 = classPath + "report/ttt2.docx";
        //        String template3 = classPath + "report/t3.docx";

        String outPath = classPath + "report/merge_" + (int) (Math.random() * 10000) + ".docx";

        WordTemplateFactory.getInstance().mergeDocument(outPath, new String[] { template1, template2 });
    }

    @Test
    public void ubbTest() {
        String classPath = this.getClass().getResource("/").getPath();

        String template1 = classPath + "report/t1.docx";

        String outPath = classPath + "report/ubb_" + (int) (Math.random() * 10000) + ".docx";

        WordData data = new WordData();
        data.addTextField("createYear", "单独[b]粗体[/b]一个");
        data.addTextField("test_b", "多个[b]粗体[/b]多个文字[b]加粗[/b]多组标记");
        data.addTextField("test_i", "多个[i]斜体[/i]多个文字[i]斜体[/i]多组标记");
        data.addTextField("test_u", "多个[u]下划线[/u]多个文字[u]下线[/u]多组标记");
        data.addTextField("test_strike", "多个[strike]删除线[/strike]多个文字[strike]删除[/strike]多组标记");
        data.addTextField("test_strikes", "多个[strikes]双删除线[/strikes]多个文字[strikes]双删[/strikes]多组标记");
        data.addTextField("test_size", "大一号的[size=30]大字体[/size]阿道夫");
        data.addTextField("test_color", "看一看[color=00ff77]字色[/color]加[color=0fa0Fa]另一色[/color]");
        data.addTextField("test_font", "字体的[font=黑体]黑体[/font]加[font=隶书]隶书[/font]");
        data.addTextField("test_all",
                "组合[strikes]双删[/strikes][u]下[b]粗[/b]划线[/u]多个[b]粗[i]粗斜[/i]粗[size=30]大[color=ff0000]红[/color]粗[/size]字[/b][font=微软雅黑]微软[strike]删除[/strike]雅黑[/font]组标记");
        WordFactory.getInstance().reportByTemplate(template1, outPath, data);
    }

    @Test
    public void hottest() {

        WordData data = new WordData();
        List<WordData> dataList = new ArrayList<WordData>();
        {
            WordData rowData = new WordData();
            rowData.addTextField("title", "标题");
            dataList.add(rowData);
        }
        {
            WordData rowData = new WordData();
            rowData.addTextField("title", "标题2");
            dataList.add(rowData);
        }
        {
            WordData rowData = new WordData();
            rowData.addTextField("title", "标题2444444");
            dataList.add(rowData);
        }
        data.addTable("hottest", dataList);

        String classPath = this.getClass().getResource("/").getPath();
        String srcPath = classPath + "report/ttt2.docx";
        String destPath = classPath + "report/hottest_" + Long.parseLong((Math.random() * 10000 + "").split("\\.")[0])
                + ".docx";

        WordFactory.getInstance().reportByTemplate(srcPath, destPath, data);
    }

}
