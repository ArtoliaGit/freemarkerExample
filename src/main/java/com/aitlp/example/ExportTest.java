package com.aitlp.example;

import java.util.HashMap;
import java.util.Map;

public class ExportTest {
    public static void main(String[] args) {
        ExportTest test = new ExportTest();
        test.testDoc();
        test.testDocx();
    }

    public void testDoc() {
        WordUtil wordUtil = new WordUtil();
        Map<String, String> dataMap = new HashMap<>();
        dataMap.put("name", "四个空格-https://www.4spaces.org");
        wordUtil.ftlToDoc("docTemplete.ftl", dataMap, "E:\\freemarker\\testDoc.doc");
    }

    public void testDocx() {
        // docx模板文件的路径和文件名
        String docxTemplate = "E:\\freemarker\\template\\docxTemplate.docx";

        // docx模板文件名称，该文件可以直接使用解压软件打开docx文件，复制word/document.xml文件内容进行修改
        String docxXmlTemplate = "docxTemplate.xml";

        // docx需要的临时xml文件路径，名称和路径都无所谓，只是中间过程会用到，之后可以删除,文件不需存在，但路径必须存在
        String tempDocxXmlPath = "E:\\freemarker\\temp\\temp.xml";

        // 目标文件名
        String outputFilePath = "E:\\freemarker\\testDocx.docx";

        // 需要动态传入的数据
        Map<String, Object> params = new HashMap<String, Object>();
        params.put("name", "四个空格-https://www.4spaces.org");

        WordUtil xtd = new WordUtil();
        xtd.xmlToDocx(docxTemplate, docxXmlTemplate, tempDocxXmlPath, params, outputFilePath);
    }
}
