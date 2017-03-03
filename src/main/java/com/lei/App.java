package com.lei;

import com.lei.util.WordReplaceTextUtil;
import org.apache.poi.hwpf.HWPFDocument;

/**
 * Hello world!
 *
 */
public class App 
{
    public static final String SOURCE_FILE = "mycv.doc";
    public static final String OUTPUT_FILE = "newcv.doc";

    public static void main(String[] args) throws Exception {
        WordReplaceTextUtil instance = new WordReplaceTextUtil();
        HWPFDocument doc = instance.openDocument(SOURCE_FILE);
        if (doc != null) {
            doc = instance.replaceText(doc, "张三", "李四");
            doc = instance.replaceText(doc, "15175232122", "8888888888");
            instance.saveDocument(doc, OUTPUT_FILE);
        }
    }
}
