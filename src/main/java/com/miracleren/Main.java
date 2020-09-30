package com.miracleren;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.OutputStream;
import java.util.*;

public class Main {

    public static void main(String[] args) {
        // write your code here

        //示例
        String path = "D://print//docMachine.docx";
        NiceDoc doc = new NiceDoc(path);
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("title", "测试文书记录");
        map.put("same", 1);
        map.put("nosame", "无说明");
        map.put("parson", 6);
        map.put("prodate", "2019-10-10");
        map.put("proname", "东莞生产总企业");
        map.put("isshow", 1);
        doc.setLabel(map);


        List<Map<String, Object>> table1 = new ArrayList();
        Map<String, Object> tableMap1 = new HashMap();
        tableMap1.put("name", "陈先生");
        tableMap1.put("date", "2020");
        tableMap1.put("code", "代码");
        table1.add(tableMap1);
        Map<String, Object> tableMap2 = new HashMap();
        tableMap2.put("name", "何先生");
        tableMap2.put("date", "2019");
        tableMap2.put("code", "代码2");
        table1.add(tableMap2);
        doc.setTable(table1, "firstTable");


        List<Map<String, Object>> table2 = new ArrayList();
        Map<String, Object> second = new HashMap();
        second.put("type", "皮革");
        second.put("desc", "牛皮、羊皮");
        table2.add(second);
        Map<String, Object> second1 = new HashMap();
        second1.put("type", "皮革");
        second1.put("desc", "牛皮铁、羊皮钢");
        table2.add(second1);
        Map<String, Object> second2 = new HashMap();
        second2.put("type", "工业");
        second2.put("desc", "铁、钢");
        table2.add(second2);
        Map<String, Object> second3 = new HashMap();
        second3.put("type", "工业");
        second3.put("desc", "铁、钢");
        table2.add(second3);
        Map<String, Object> second4 = new HashMap();
        second4.put("type", "工业2");
        second4.put("desc", "铁、钢");
        table2.add(second4);
        doc.setTable(table2, "secondTable");

        doc.save("D://print//docx//" + UUID.randomUUID() + ".docx");
        //doc.saveOnlyComments("D://print//docx//" + UUID.randomUUID() + ".docx");
        //doc.savePdf("D://print//docx//" + UUID.randomUUID() + ".pdf");

        //OutputStream stream = doc.saveStream();



        //File file = new File("D://file.")

        System.out.println("aspword-cli run");
    }
}
