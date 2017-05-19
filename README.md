## 功能
  WORD生成工具 ，实现Data ——> Word<br>
  按照Word模板填充数据后生成Word文件，保留模板的格式。<br>
  功能包括 段落迭代 / 表格行迭代 / 文本填充 / 图片填充 / 超链接填充 / 样式克隆 / ubb标记样式 / word合并
  
## 模板
  模板为07版本word.docx文件，样式格式所见即所得。<br>
  直接对标记设置字体样式，填充的数据即应用相应字体样式。

## 标记
  \${text} ：普通文本<br>
  \${link} ：超链接文本，与文本标记相同，在模板中为此标记设置超链接，url可随意指定<br>
  \${image} ：图片标记，与文本标记相同，传递数据时设置图片数据，可指定宽度高度 \${image?w=100&h=100}，未指定则应用默认值<br>
  \${key:start} ：段落迭代器开始标记，与end必须为完整的一组<br>
  \${key:end} ：段落迭代器结束标记，与start必须为完整的一组<br>
  \${key:rows} ：表格迭代行标记，在表格行第一列加入此标记，该标记声明此行为迭代行

## 标记作用域
  \${key.text} 迭代中通过"key点text"方式指定局部变量，在迭代范围内应用，\${text}则为全局变量，可重复使用相同的变量，输出数据相同。
  
## 动态样式
  采用UBB标记格式，支持以下几种标记。<br>
  加粗：[b]加粗文字[/b]<br>
  斜体：[i]斜体文字[/i]<br>
  下划线：[u]下划线[/u]<br>
  单删除线：[strike]单删除线[/strike]<br>
  双删除线：[strikes]双删除线[/strikes]<br>
  字体：[font=黑体]设置字体[/font]<br>
  字色：[color=ff0000]字色[/color]<br>
  字号：[size=16]设置字号[size]<br>
  换行符：\n
  
## WordData数据
### addTextField
  public void addTextField(java.lang.String key, java.lang.String text)<br>
  添加文本字段<br>
  参数:<br>
  key - \${key}<br>
  text - 显示文本

### addLinkField
  public void addLinkField(java.lang.String key, java.lang.String text, java.lang.String link)<br>
  添加带链接字段<br>
  参数:<br>
  key - \${key}<br>
  text - 显示文本<br>
  link - 链接地址

### addImageField
  public void addImageField(java.lang.String key, java.io.File imageFile, java.lang.Integer width, java.lang.Integer height)<br>
  添加图片字段<br>
  参数:<br>
  key - \${key}<br>
  imageFile - 图片<br>
  width - 固定宽度,null:自动<br>
  height - 固定高度,null:自动
 
### addIterator
  public void addIterator(java.lang.String key, java.util.List<WordData> dataList)<br>
  添加迭代数据<br>
  参数:<br>
  key - 唯一键,重复会覆盖<br>
  dataList - 数据集<br>
 
### addTable
  public void addTable(java.lang.String key, java.util.List<WordData> dataList)<br>
  给key表格添加列数据，重复columnIndex列新增一行数据<br>
  参数:<br>
  key - 第一个参照行的第一个单元格标记\${key:rows}<br>
  dataList - 数据 没条记录为一行
  
## WordFactory
### getInstance
  public static WordFactory getInstance()<br>
  获取实例<br>
  返回:<br>
  WordFactory

### reportByTemplate
  public boolean reportByTemplate(java.lang.String templateFilePath,java.lang.String outFilePath,WordData data)<br>
  根据模板路径生成报告,输出到指定位置.<br>
  模板支持\${标记} “:”为特殊标记符，应避免使用 <br>
  \${key:start}模板循环开始标记，独占一行 <br>
  \${key:end}模板循环结束标记，独占一行<br>
  参数:<br>
  templateFilePath - 模板文件路径<br>
  outFilePath - 输出文件路径<br>
  data - 报告数据<br>
  返回:<br>
  boolean 生成报告是否成功
 
### reportByTemplate
  public boolean reportByTemplate(java.io.InputStream templateFile, java.lang.String outFilePath, WordData data)<br>
  根据模板文件流生成报告,输出到指定位置.<br>
  模板支持\${标记} “:”为特殊标记符，应避免使用 <br>
  \${key:start}模板循环开始标记，独占一行<br>
  \${key:end}模板循环结束标记，独占一行<br>
  参数:<br>
  templateFile - 模板数据流<br>
  outFilePath - 输出文件路径<br>
  data - 报告数据<br>
  返回:<br>
  boolean 生成报告是否成功
 
### reportByTemplate
  public java.io.OutputStream reportByTemplate(java.lang.String templateFilePath, WordData data)<br>
  根据模板文件路径生成报告,返回文件流.<br>
  模板支持\${标记} “:”为特殊标记符，应避免使用 <br>
  \${key:start}模板循环开始标记，独占一行 \<br>
  \${key:end}模板循环结束标记，独占一行<br>
  参数:<br>
  templateFilePath - 模板文件路径<br>
  data - 报告数据<br>
  返回:<br>
  ByteArrayOutputStream
 
### reportByTemplate
  public java.io.OutputStream reportByTemplate(java.io.InputStream templateInputStream, WordData data)<br>
  根据模板文件流生成报告,返回文件流.<br>
  模板支持\${标记} “:”为特殊标记符，应避免使用 <br>
  \${key:start}模板循环开始标记，独占一行<br>
  \${key:end}模板循环结束标记，独占一行<br>
  参数:<br>
  templateInputStream - 文件流<br>
  data - 报告数据返回:ByteArrayOutputStream

## TemplateFactory
### getInstance
public static TemplateFactory getInstance()<br>
获取实例<br>
返回:<br>
TemplateFactory
 
### mergeDocument
public boolean mergeDocument(java.lang.String outFilePath, java.io.InputStream ... templateInputStreams)<br>
合并多个文档流输出到指定位置<br>
参数:<br>
outFilePath - 指定输出文件路径<br>
templateInputStreams - 模板文件输入流,多个<br>
返回:<br>
boolean true 成功 false 失败
 
### mergeDocument
public boolean mergeDocument(java.lang.String outFilePath,java.lang.String ... templateFilePaths)<br>
读取多个模板合并输出指定位置<br>
参数:<br>
outFilePath - 指定输出文件路径<br>
templateFilePaths - 模板文件位置，多个<br>
返回:<br>
boolean true 成功 false 失败
 
### mergeDocument
public java.io.OutputStream mergeDocument(java.io.InputStream ... templateInputStreams)<br>
合并多个文档流返回文件流<br>
参数:<br>
templateInputStreams - 模板文件输入流,多个<br>
返回:<br>
java.io.OutputStream 合并后的文件流
 
### mergeDocument
public java.io.OutputStream mergeDocument(java.lang.String ... templateFilePaths)<br>
读取多个模板合并返回文件流<br>
参数:<br>
templateFilePaths - 模板文件位置，多个<br>
返回:<br>
java.io.OutputStream 合并后的文件流
=======
## 功能
  WORD生成工具 ，实现Data ——> Word<br>
  按照Word模板填充数据后生成Word文件，保留模板的格式。<br>
  功能包括 段落迭代 / 表格行迭代 / 文本填充 / 图片填充 / 超链接填充 / 样式克隆 / ubb标记样式 / word合并
  
## 模板
  模板为07版本word.docx文件，样式格式所见即所得。<br>
  直接对标记设置字体样式，填充的数据即应用相应字体样式。

## 标记
  \${text} ：普通文本<br>
  \${link} ：超链接文本，与文本标记相同，在模板中为此标记设置超链接，url可随意指定<br>
  \${image} ：图片标记，与文本标记相同，传递数据时设置图片数据，可指定宽度高度 \${image?w=100&h=100}，未指定则应用默认值<br>
  \${key:start} ：段落迭代器开始标记，与end必须为完整的一组<br>
  \${key:end} ：段落迭代器结束标记，与start必须为完整的一组<br>
  \${key:rows} ：表格迭代行标记，在表格行第一列加入此标记，该标记声明此行为迭代行

## 标记作用域
  \${key.text} 迭代中通过"key点text"方式指定局部变量，在迭代范围内应用，\${text}则为全局变量，可重复使用相同的变量，输出数据相同。
  
## 动态样式
  采用UBB标记格式，支持以下几种标记。<br>
  加粗：[b]加粗文字[/b]<br>
  斜体：[i]斜体文字[/i]<br>
  下划线：[u]下划线[/u]<br>
  单删除线：[strike]单删除线[/strike]<br>
  双删除线：[strikes]双删除线[/strikes]<br>
  字体：[font=黑体]设置字体[/font]<br>
  字色：[color=ff0000]字色[/color]<br>
  字号：[size=16]设置字号[size]<br>
  换行符：\n
  
## WordData数据
### addTextField
  public void addTextField(java.lang.String key, java.lang.String text)<br>
  添加文本字段<br>
  参数:<br>
  key - \${key}<br>
  text - 显示文本

### addLinkField
  public void addLinkField(java.lang.String key, java.lang.String text, java.lang.String link)<br>
  添加带链接字段<br>
  参数:<br>
  key - \${key}<br>
  text - 显示文本<br>
  link - 链接地址

### addImageField
  public void addImageField(java.lang.String key, java.io.File imageFile, java.lang.Integer width, java.lang.Integer height)<br>
  添加图片字段<br>
  参数:<br>
  key - \${key}<br>
  imageFile - 图片<br>
  width - 固定宽度,null:自动<br>
  height - 固定高度,null:自动
 
### addIterator
  public void addIterator(java.lang.String key, java.util.List<WordData> dataList)<br>
  添加迭代数据<br>
  参数:<br>
  key - 唯一键,重复会覆盖<br>
  dataList - 数据集<br>
 
### addTable
  public void addTable(java.lang.String key, java.util.List<WordData> dataList)<br>
  给key表格添加列数据，重复columnIndex列新增一行数据<br>
  参数:<br>
  key - 第一个参照行的第一个单元格标记\${key:rows}<br>
  dataList - 数据 没条记录为一行
  
## WordFactory
### getInstance
  public static WordFactory getInstance()<br>
  获取实例<br>
  返回:<br>
  WordFactory

### reportByTemplate
  public boolean reportByTemplate(java.lang.String templateFilePath,java.lang.String outFilePath,WordData data)<br>
  根据模板路径生成报告,输出到指定位置.<br>
  模板支持\${标记} “:”为特殊标记符，应避免使用 <br>
  \${key:start}模板循环开始标记，独占一行 <br>
  \${key:end}模板循环结束标记，独占一行<br>
  参数:<br>
  templateFilePath - 模板文件路径<br>
  outFilePath - 输出文件路径<br>
  data - 报告数据<br>
  返回:<br>
  boolean 生成报告是否成功
 
### reportByTemplate
  public boolean reportByTemplate(java.io.InputStream templateFile, java.lang.String outFilePath, WordData data)<br>
  根据模板文件流生成报告,输出到指定位置.<br>
  模板支持\${标记} “:”为特殊标记符，应避免使用 <br>
  \${key:start}模板循环开始标记，独占一行<br>
  \${key:end}模板循环结束标记，独占一行<br>
  参数:<br>
  templateFile - 模板数据流<br>
  outFilePath - 输出文件路径<br>
  data - 报告数据<br>
  返回:<br>
  boolean 生成报告是否成功
 
### reportByTemplate
  public java.io.OutputStream reportByTemplate(java.lang.String templateFilePath, WordData data)<br>
  根据模板文件路径生成报告,返回文件流.<br>
  模板支持\${标记} “:”为特殊标记符，应避免使用 <br>
  \${key:start}模板循环开始标记，独占一行 \<br>
  \${key:end}模板循环结束标记，独占一行<br>
  参数:<br>
  templateFilePath - 模板文件路径<br>
  data - 报告数据<br>
  返回:<br>
  ByteArrayOutputStream
 
### reportByTemplate
  public java.io.OutputStream reportByTemplate(java.io.InputStream templateInputStream, WordData data)<br>
  根据模板文件流生成报告,返回文件流.<br>
  模板支持\${标记} “:”为特殊标记符，应避免使用 <br>
  \${key:start}模板循环开始标记，独占一行<br>
  \${key:end}模板循环结束标记，独占一行<br>
  参数:<br>
  templateInputStream - 文件流<br>
  data - 报告数据返回:ByteArrayOutputStream

## TemplateFactory
### getInstance
public static TemplateFactory getInstance()<br>
获取实例<br>
返回:<br>
TemplateFactory
 
### mergeDocument
public boolean mergeDocument(java.lang.String outFilePath, java.io.InputStream ... templateInputStreams)<br>
合并多个文档流输出到指定位置<br>
参数:<br>
outFilePath - 指定输出文件路径<br>
templateInputStreams - 模板文件输入流,多个<br>
返回:<br>
boolean true 成功 false 失败
 
### mergeDocument
public boolean mergeDocument(java.lang.String outFilePath,java.lang.String ... templateFilePaths)<br>
读取多个模板合并输出指定位置<br>
参数:<br>
outFilePath - 指定输出文件路径<br>
templateFilePaths - 模板文件位置，多个<br>
返回:<br>
boolean true 成功 false 失败
 
### mergeDocument
public java.io.OutputStream mergeDocument(java.io.InputStream ... templateInputStreams)<br>
合并多个文档流返回文件流<br>
参数:<br>
templateInputStreams - 模板文件输入流,多个<br>
返回:<br>
java.io.OutputStream 合并后的文件流
 
### mergeDocument
public java.io.OutputStream mergeDocument(java.lang.String ... templateFilePaths)<br>
读取多个模板合并返回文件流<br>
参数:<br>
templateFilePaths - 模板文件位置，多个<br>
返回:<br>
java.io.OutputStream 合并后的文件流