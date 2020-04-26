package com.img;


import com.img.utils.DocUtil;
import com.img.utils.ImgUtils;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import freemarker.template.TemplateException;

import java.io.*;
import java.util.*;

/**
 * @Description: TODO
 * @author: ZJ
 * @time: 2020/4/25 18:49
 * @Version: 1.0.0
 */
public class TestDoc {


    public static void main(String[] args) {
        // 图片文件目录
        String filePath = "E:\\imgs\\imgs";
        // 输出word目录
        String outFilePath = "E:\\imgs\\word";
        File file = new File(filePath);
        File[] files = file.listFiles();
        // 是否处理word为另存为 , 如果改为 true, 则会将word重新转码一次, 特别慢
        Boolean tranFlag = true;
        List<String> successFileName = new ArrayList<String>();
        List<String> failFileName = new ArrayList<String>();
        failFileName.add("#--------------------说 明---------------------------------#");
        failFileName.add("#--------------文件夹名--错误码-----------------------------#");
        failFileName.add("#--------------2  图片超过18张-------------------------------#");
        failFileName.add("#--------------3  图片少于4张-------------------------------#");
        failFileName.add("#--------------4  保存异常, 文件中存在除了图片之外的文件------#");
        failFileName.add("#-----------------------------------------------------------#");
        // 文件总数量
        int fileCount = 0;
        // 成功数量
        int successCount = 0;
        // 失败
        int failCount = 0;
        for (File innerFile : files) {
            if (innerFile.isDirectory()) {
                ++fileCount;
                String path = innerFile.getPath();
                System.out.println("内部文件path--"+path);
                int resultFlag = doCreateWord(path, innerFile.getName(), outFilePath,tranFlag);
                if (1 == resultFlag) {
                    ++successCount;
                    successFileName.add(innerFile.getName());
                    System.out.println(innerFile.getName() + "-> 处理成功");
                } else {
                    ++failCount;
                    failFileName.add(innerFile.getName() + "----" + resultFlag);
                    System.out.println(innerFile.getName() + "-> 处理失败==>" + resultFlag);
                }
            }
        }
        successFileName.add(0,"处理总数:( "+fileCount+" ) 失败总数:( "+successCount+" )");
        failFileName.add(0,"处理总数:( "+fileCount+" ) 失败总数:( "+failCount+" )" );

        printOut(successFileName,outFilePath,"00000_成功结果.txt");
        printOut(failFileName,outFilePath,"00000_失败结果.txt");
    }

    private static void printOut(List<String> stringList, String outFilePath, String txtName) {
        OutputStreamWriter dos1 = null;
        try {
            outFilePath = outFilePath + File.separator + txtName;
            File f = new File(outFilePath);
            FileOutputStream fos1  = new FileOutputStream(f);

            dos1 = new OutputStreamWriter(fos1);

            for (String lineStr : stringList ) {
                dos1.write(lineStr);
                dos1.write("\r\n");
            }

            dos1.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private static int doCreateWord(String path, String filename, String outFilePath, Boolean tranFlag) {

        String saveTempPath = outFilePath + File.separator + filename + "_temp" + ".doc";
        String savePath = outFilePath + File.separator + filename + ".doc";
        int flag = 1;
        try {

            DocUtil docUtil = new DocUtil();
            Map<String, Object> dataMap = new HashMap<String, Object>();

            dataMap.put("filename1", filename.substring(0,4));
            File file = new File(path);
            File[] files = file.listFiles();

            int len = files.length;
            System.out.println(filename + "-> 存在图片" + len);
            if(len > 18){
                flag = 2;
            }else if(len < 4 ){
                flag = 3;
            }else {
                for (int i = 0; i < len; i++) {
                    File innerFile = files[i];
                    String name = innerFile.getName();
                    String lastName = name.substring(name.lastIndexOf(".") + 1).toUpperCase();
                    if (lastName.matches("^[(JPG)|(PNG)|(GIF)|(JPEG)|(JPE)|(JFIF)]+$")) {
                        dataMap.put("img" + (i + 1), ImgUtils.getImageStr(innerFile));
                    }else{
                        System.out.println(filename + "-> 存在不是图片的文件: " + name);
                        len = len - 1;
                        flag = 4;
                    }
                }
                String downloadType = "doc_" + len + ".xml";
                docUtil.createDoc(dataMap, downloadType, saveTempPath);
                if(tranFlag){
                    createNewWord(saveTempPath,savePath);
                }

            }



        } catch (TemplateException e) {
            flag = 9;
            e.printStackTrace();
        } catch (IOException e) {
            flag = 9;
            e.printStackTrace();
        } catch (Exception e) {
            flag = 9;
            e.printStackTrace();
        } finally {
            return flag;
        }

    }

    private static void createNewWord(String saveTempPath, String savePath) {
        ActiveXComponent _app = new ActiveXComponent("Word.Application");
        _app.setProperty("Visible", Variant.VT_FALSE);
        Dispatch documents = _app.getProperty("Documents").toDispatch();
        // 打开FreeMarker生成的Word文档
        Dispatch doc = Dispatch.call(documents, "Open",saveTempPath, Variant.VT_FALSE, Variant.VT_TRUE).toDispatch();
        // 另存为新的Word文档
        Dispatch.call(doc, "SaveAs", savePath, Variant.VT_FALSE, Variant.VT_TRUE);

        Dispatch.call(doc, "Close", Variant.VT_FALSE);
        _app.invoke("Quit", new Variant[] {});
        ComThread.Release();
        File file = new File(saveTempPath);
        file.delete();
    }

    public void aa() {
        DocUtil docUtil = new DocUtil();
        Map<String, Object> dataMap = new HashMap<String, Object>();

        dataMap.put("filename1", "这个文件名称");
        dataMap.put("img1", ImgUtils.getImageStr("E:\\imgs\\imgs\\第一组\\图片1 (1).jpg"));
        dataMap.put("img2", ImgUtils.getImageStr("E:\\imgs\\imgs\\第一组\\图片1 (2).jpg"));
        dataMap.put("img3", ImgUtils.getImageStr("E:\\imgs\\imgs\\第一组\\图片1 (3).jpg"));
        dataMap.put("img4", ImgUtils.getImageStr("E:\\imgs\\imgs\\第一组\\图片1 (4).jpg"));
        dataMap.put("img5", ImgUtils.getImageStr("E:\\imgs\\imgs\\第一组\\图片1 (5).jpg"));
        dataMap.put("img6", ImgUtils.getImageStr("E:\\imgs\\imgs\\第一组\\图片1 (6).jpg"));
        dataMap.put("img7", ImgUtils.getImageStr("E:\\imgs\\imgs\\第一组\\图片1 (7).jpg"));

        dataMap.put("img8", ImgUtils.getImageStr("E:\\imgs\\imgs\\第一组\\图片1 (1).png"));

        dataMap.put("img9", ImgUtils.getImageStr("E:\\imgs\\imgs\\第一组\\图片1 (2).png"));
        dataMap.put("img10", ImgUtils.getImageStr("E:\\imgs\\imgs\\第一组\\图片1 (3).png"));

        dataMap.put("img11", ImgUtils.getImageStr("E:\\imgs\\imgs\\第一组\\图片1 (1).jpg"));
        dataMap.put("img12", ImgUtils.getImageStr("E:\\imgs\\imgs\\第一组\\图片1 (2).jpg"));
        dataMap.put("img13", ImgUtils.getImageStr("E:\\imgs\\imgs\\第一组\\图片1 (3).jpg"));
        dataMap.put("img14", ImgUtils.getImageStr("E:\\imgs\\imgs\\第一组\\图片1 (4).jpg"));
        dataMap.put("img15", ImgUtils.getImageStr("E:\\imgs\\imgs\\第一组\\图片1 (5).jpg"));
        dataMap.put("img16", ImgUtils.getImageStr("E:\\imgs\\imgs\\第一组\\图片1 (6).jpg"));
        dataMap.put("img17", ImgUtils.getImageStr("E:\\imgs\\imgs\\第一组\\图片1 (7).jpg"));

        dataMap.put("img18", ImgUtils.getImageStr("E:\\imgs\\imgs\\第一组\\图片1 (1).png"));
        //docUtil.createDoc(dataMap, "doc_15.xml", "E:\\imgs\\18_15xxx.doc");
        System.out.println("Word文件已生成完毕！目录地址：D:\\eclipseWorkspace\\bridgereport\\4_1_1.doc");
    }
}
