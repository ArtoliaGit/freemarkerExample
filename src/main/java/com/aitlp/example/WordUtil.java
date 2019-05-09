package com.aitlp.example;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;

import java.io.*;
import java.util.Enumeration;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipException;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

public class WordUtil {
    // 初始化配置
    private Configuration configuration;
    // 模板文件目录(doc模板ftl文件，docx模板xml文件都放在此目录)
    private String templateDir = "E:\\freemarker\\template\\";


    public WordUtil() {
        configuration = new Configuration();
        /** 设置编码 **/
        configuration.setDefaultEncoding("utf-8");
        /** 加载目录 **/
        try {
            configuration.setDirectoryForTemplateLoading(new File(templateDir));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 生成doc文件
     *
     * @param ftlFileName 模板ftl文件的名称
     * @param params      动态传入的数据参数
     * @param outFilePath 生成的最终doc文件的保存完整路径
     */
    public void ftlToDoc(String ftlFileName, Map params, String outFilePath) {
        try {
            /** 加载模板文件 **/
            Template template = configuration.getTemplate(ftlFileName);
            /** 指定输出word文件的路径 **/
            File docFile = new File(outFilePath);
            FileOutputStream fos = new FileOutputStream(docFile);
            Writer bufferedWriter = new BufferedWriter(new OutputStreamWriter(fos, "utf-8"), 10240);
            template.process(params, bufferedWriter);
            if (bufferedWriter != null) {
                bufferedWriter.close();
            }
        } catch (TemplateException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 生成docx文件
     *
     * @param docxTemplate    docx的模板docx文件路径
     * @param docxXmlTemplate docx的模板xml文件名称
     * @param tempDocxXmlPath docx的临时xml文件(docx的模板xml文件填充完数据生成的临时文件)
     * @param params          填充到docx的临时xml文件中的数据
     * @param toFilePath      最终输出的docx文件路径
     */
    public void xmlToDocx(String docxTemplate, String docxXmlTemplate, String tempDocxXmlPath, Map params, String toFilePath) {
        try {
            Template template = configuration.getTemplate(docxXmlTemplate);

            Writer fileWriter = new FileWriter(new File(tempDocxXmlPath));
            template.process(params, fileWriter);
            if (fileWriter != null) {
                fileWriter.close();
            }

            File docxFile = new File(docxTemplate);
            ZipFile zipFile = new ZipFile(docxFile);
            Enumeration<? extends ZipEntry> zipEntrys = zipFile.entries();
            ZipOutputStream zipout = new ZipOutputStream(new FileOutputStream(toFilePath));
            int len = -1;
            byte[] buffer = new byte[1024];
            while (zipEntrys.hasMoreElements()) {
                ZipEntry next = zipEntrys.nextElement();
                InputStream is = zipFile.getInputStream(next);
                //把输入流的文件传到输出流中 如果是word/document.xml由我们输入
                zipout.putNextEntry(new ZipEntry(next.toString()));
                if ("word/document.xml".equals(next.toString())) {
                    //InputStream in = new FileInputStream(new File(XmlToDocx.class.getClassLoader().getResource("").toURI().getPath()+"template/test.xml"));
                    InputStream in = new FileInputStream(tempDocxXmlPath);
                    while ((len = in.read(buffer)) != -1) {
                        zipout.write(buffer, 0, len);
                    }
                    in.close();
                } else {
                    while ((len = is.read(buffer)) != -1) {
                        zipout.write(buffer, 0, len);
                    }
                    is.close();
                }
            }
            zipout.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (ZipException e) {
            e.printStackTrace();
        } catch (TemplateException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
