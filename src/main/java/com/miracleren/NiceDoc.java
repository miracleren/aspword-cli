package com.miracleren;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;

import com.aspose.words.*;
import jdk.nashorn.internal.runtime.regexp.joni.Regex;

public class NiceDoc {

    /**
     * 20200715 基于aspose模板生成word，
     * by miracleren
     */

    private static final String ASPOSE_VERSION = "18.6.0";
    Document doc;

    /**
     * 验证
     *
     * @return
     */
    private boolean setLicense() {
        try {
            ClassLoader loader = Thread.currentThread().getContextClassLoader();
            InputStream license = this.getClass().getResourceAsStream("/license.xml");
            License aposeLic = new License();
            aposeLic.setLicense(license);
            return true;
        } catch (Exception e) {
            System.out.println(e.toString());
            return false;
        }
    }

    /**
     * 初始化模板
     *
     * @param tempPath
     */
    public NiceDoc(String tempPath) {
        try {
            if (setLicense()) {
                doc = new Document(tempPath);
                System.out.println("create docx successully");
            }
        } catch (Exception e) {
            System.out.println(e.toString());
        }
    }
    public NiceDoc(InputStream tempStream) {
        try {
            if (setLicense()) {
                doc = new Document(tempStream);
                System.out.println("create docx successully");
            }
        } catch (Exception e) {
            System.out.println(e.toString());
        }
    }

    /**
     * 标签数据替换
     *
     * @param values map值列表
     */
    public void setLabel(Map<String, Object> values) {
        Range range = doc.getRange();
        String wordText = range.getText();
        Matcher pars = matcher(wordText);
        while (pars.find()) {
            //System.out.println(pars.group().toString());
            String con = pars.group();
            String[] cons = con.split(":");
            //纯内容标签替换
            try {
                if (cons.length == 1) {
                    String labVal = StringOf(values.get(con));
                    rangeReplace(con, labVal);
                } else {
                    if (cons.length == 3) {
                        //类型标签
                        String typeName = cons[0];
                        String typePar = cons[1];
                        String typeVal = cons[2];
                        if ("SC".equals(typeName)) {
                            //单选
                            if (StringOf(values.get(typePar)).equals(typeVal))
                                rangeReplace(con, "√");
                            else
                                rangeReplace(con, "□");
                        } else if ("MC".equals(typeName)) {
                            //多选
                            //String value = StringOf(values.get(typePar));
                            int parval = values.get(typePar) == null ? 0 : Integer.parseInt(StringOf(values.get(typePar)));
                            int val = Integer.parseInt(typeVal);
                            if ((parval & val) == val)
                                rangeReplace(con, "√");
                            else
                                rangeReplace(con, "□");
                        }
                    }
                }
            } catch (Exception e) {
                System.out.println(con + ":" + e);
            }

        }

        //标签更新完成，处理表达式
        setSyntax(values);
    }

    /**
     * 表格循环数据填充
     *
     * @param list
     * @param tableName
     */
    public void setTable(List<Map<String, Object>> list, String tableName) {
        NodeCollection bookTables = doc.getChildNodes(NodeType.TABLE, true);
        for (Object table : bookTables) {
            Table tb = (Table) table;
            //判断是否循环列表
            String rowFistText = tb.getRows().get(0).getText();
            String tableConfig = getFirstParName(rowFistText);
            if (!tableConfig.equals("")) {
                //第一行为表格配置信息
                String[] cons = tableConfig.split(":");
                if (!cons[0].equals("TABLE") && !cons[1].equals(tableName))
                    break;
            } else
                break;

            //查找配置循环列
            int i = 0, tempIndex = -1;
            Row tempRow = null;
            for (Row trow : ((Table) table).getRows()) {
                if (tempRow != null)
                    break;
                for (Cell tcell : trow.getCells()) {
                    if (getFirstParName(tcell.getText()).contains("COL")) {
                        tempRow = trow;
                        tempIndex = i;
                        break;
                    }
                }
                i++;
            }
            if (tempRow == null)
                return;

            //克隆行，并赋值
            for (Map<String, Object> rowData : list) {
                Row newRow = (Row) tempRow.deepClone(true);
                for (Cell newRowCell : newRow.getCells()) {
                    String cellPars = getFirstParName(newRowCell.getRange().getText());
                    if (!cellPars.equals("")) {
                        String[] pars = cellPars.split(":");
                        if (pars[0].equals("COL")) {
                            rangeReplace(newRowCell.getRange(), cellPars, StringOf(rowData.get(pars[1])));
                            //newRow.getRange().replace("{{" + cellPars + "}}", StringOf(rowData.get(pars[1])),new FindReplaceOptions()
//                            try {
//                                newRowCell.getRange().replace("{{" + cellPars + "}}", StringOf(rowData.get(pars[1])), new FindReplaceOptions());
//                            } catch (Exception e) {
//                                System.out.println(cellPars + ">>>>>>" + e.toString());
//                            }
                        }
                    }
                }
                ((Table) table).appendChild(newRow);
            }

            //清除配置行
            ((Table) table).removeChild(((Table) table).getFirstRow());
            ((Table) table).removeChild(tempRow);
        }
    }

    /**
     * 表达式判断
     * <p>
     * 目前支持
     * {{V-IF:par}}{{END:par}}  显示隐藏数据,等号目前支持 ==，！=
     */
    public void setSyntax(Map<String, Object> values) {
        Range range = doc.getRange();
        String wordText = range.getText();
        Matcher pars = matcher(wordText);
        while (pars.find()) {
            String con = pars.group();
            //if显示隐藏表达式
            if (con.contains("V-IF:")) {
                String[] cons = con.split(":");
                String syn = cons[1];
                if (syn.contains("==")) {
                    String[] tem = syn.split("==");
                    if (StringOf(values.get(tem[0])).equals(tem[1])) {
                        rangeReplace(con, "");
                        rangeReplace("END:" + tem[0], "");
                    } else {
                        Pattern pattern = Pattern.compile("(?=\\{\\{" + con + "\\}\\})(.+?)(?<=\\{\\{END:" + tem[0] + "\\}\\})", Pattern.CASE_INSENSITIVE);
                        rangeReplace(pattern, "");

                    }
                } else if (syn.contains("!=")) {
                    String[] tem = syn.split("!=");
                    if (!StringOf(values.get(tem[0])).equals(tem[1])) {
                        rangeReplace(con, "");
                        rangeReplace("END:" + tem[0], "");
                    } else {
                        Pattern pattern = Pattern.compile("(?=\\{\\{" + con + "\\}\\})(.+?)(?<=\\{\\{END:" + tem[0] + "\\}\\})", Pattern.CASE_INSENSITIVE);
                        rangeReplace(pattern, "");

                    }
                } else {
                    if (StringOf(values.get(cons)).equals("true")) {
                        rangeReplace(con, "");
                        rangeReplace("END:" + cons, "");
                    } else {
                        Pattern pattern = Pattern.compile("(?=\\{\\{" + con + "\\}\\})(.+?)(?<=\\{\\{END:" + cons + "\\}\\})", Pattern.CASE_INSENSITIVE);
                        rangeReplace(pattern, "");

                    }
                }
                System.out.println("不支持当前表达式：" + syn);
            }
        }
    }

    /**
     * 实体类转map
     *
     * @param object
     * @return
     */
    public static Map<String, Object> entityToMap(Object object) {
        Map<String, Object> map = new HashMap<String, Object>();
        for (java.lang.reflect.Field field : object.getClass().getDeclaredFields()) {
            try {
                boolean flag = field.isAccessible();
                field.setAccessible(true);
                Object o = field.get(object);
                map.put(field.getName(), o);
                field.setAccessible(flag);
            } catch (Exception e) {
                System.out.println("实体类转换：" + e.toString());
            }
        }
        return map;
    }

    public static Map<String, Object> entityToMap(Object object,boolean isLower) {
        Map<String, Object> map = new HashMap<String, Object>();
        for (java.lang.reflect.Field field : object.getClass().getDeclaredFields()) {
            try {
                boolean flag = field.isAccessible();
                field.setAccessible(true);
                Object o = field.get(object);
                String name = isLower == true? field.getName().toLowerCase():field.getName().toUpperCase();
                map.put(name, o);
                field.setAccessible(flag);
            } catch (Exception e) {
                System.out.println("实体类转换：" + e.toString());
            }
        }
        return map;
    }

    /**
     * 文本替换
     *
     * @param oldStr
     * @param newStr
     */
    private void rangeReplace(String oldStr, String newStr) {
        Range range = doc.getRange();
        try {
            range.replace("{{" + oldStr + "}}", newStr, new FindReplaceOptions());
        } catch (Exception e) {
            System.out.println(oldStr + ">>>>>>" + e.toString());
        }
    }

    /**
     * 文本替换
     *
     * @param pattern
     * @param newStr
     */
    private void rangeReplace(Pattern pattern, String newStr) {
        Range range = doc.getRange();
        try {
            range.replace(pattern, newStr, new FindReplaceOptions());
        } catch (Exception e) {
            System.out.println("pattern >>>>>>" + e.toString());
        }
    }

    /**
     * 文本替换
     *
     * @param range
     * @param oldStr
     * @param newStr
     */
    private void rangeReplace(Range range, String oldStr, String newStr) {
        try {
            range.replace("{{" + oldStr + "}}", newStr, new FindReplaceOptions());
            //range.replace("{{" + oldStr + "}}", newStr,true,false);
        } catch (Exception e) {
            System.out.println(oldStr + ">>>>>>" + e.toString());
        }
    }


    /**
     * {{par}} 参数查找正则
     *
     * @param str 查找串
     * @return 返结果
     */
    private Matcher matcher(String str) {
        Pattern pattern = Pattern.compile("(?<=\\{\\{)(.+?)(?=\\}\\})", Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(str);
        return matcher;
    }

    /**
     * 获取数据里第一个标签名称
     *
     * @param str
     * @return
     */
    private String getFirstParName(String str) {
        Pattern pattern = Pattern.compile("(?<=\\{\\{)(.+?)(?=\\}\\})", Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(str);
        if (matcher.find())
            return matcher.group();
        else
            return "";
    }


    /**
     * 空字符转占位空格
     */
    private String StringOf(Object val) {
        return val == null ? "        " : val.toString();
    }

    /**
     * 去除word水印
     *
     * @param ptch
     */
//    private void removeMark(String ptch) {
//        try {
//            //word去除水印
//            InputStream in = new FileInputStream(ptch);
//            XWPFDocument doctemp = new XWPFDocument(in);
//            doctemp.removeBodyElement(0);
//            OutputStream outputStream = new FileOutputStream(ptch);
//            doctemp.write(outputStream);
//            doctemp.close();
//        } catch (Exception e) {
//            System.out.println(e.toString());
//        }
//    }

    /**
     * 去除pdf水印
     *
     * @param ptch
     */
//    private void removeMarkPdf(String ptch) {
//        try {
//            com.itextpdf.text.Document pdf = new com.itextpdf.text.Document();
//            PdfWriter.getInstance(pdf, new FileOutputStream(ptch));
//            pdf.open();
//        } catch (Exception e) {
//            System.out.println(e.toString());
//        }
//    }
    public boolean save(String ptch) {
        try {
            doc.save(ptch);
            return true;
        } catch (Exception e) {
            System.out.println("保存失败：" + e.toString());
            return false;
        }
    }

    public boolean saveOnlyComments(String ptch) {
        try {
            doc.protect(ProtectionType.ALLOW_ONLY_COMMENTS, "teamoneit");
            doc.save(ptch);
            return true;
        } catch (Exception e) {
            System.out.println("保存失败：" + e.toString());
            return false;
        }
    }

    public boolean savePdf(String ptch) {
        try {
            PdfSaveOptions op = new PdfSaveOptions();
            op.setSaveFormat(SaveFormat.PDF);
            doc.save(ptch, op);
            return true;
        } catch (Exception e) {
            System.out.println("保存失败：" + e.toString());
            return false;
        }
    }

    public OutputStream saveStream() {
        OutputStream ms = null;
        try {
            doc.save(ms, new OoxmlSaveOptions(SaveFormat.DOC));
        } catch (Exception e) {
            System.out.println("saveStream保存失败：" + e.toString());
        }
        return ms;
    }

    protected void finalize() {
        doc.remove();
    }

}
