package com.img.utils;

import sun.misc.BASE64Encoder;

import java.io.*;

/**
 * @Description: TODO
 * @author: ZJ
 * @time: 2020/4/25 18:50
 * @Version: 1.0.0
 */
public class ImgUtils {
    public static String getImageStr(String imgFile) {

        InputStream in = null;
        byte[] data = null;
        try {
            in = new FileInputStream(imgFile);
            data = new byte[in.available()];
            in.read(data);
            in.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        BASE64Encoder encoder = new BASE64Encoder();
        return encoder.encode(data);
    }


    public static String getImageStr(File innerFile) throws IOException {
        InputStream in = null;
        byte[] data = null;

        in = new FileInputStream(innerFile);
        data = new byte[in.available()];
        in.read(data);
        in.close();
        BASE64Encoder encoder = new BASE64Encoder();
        return encoder.encode(data);
    }
}
