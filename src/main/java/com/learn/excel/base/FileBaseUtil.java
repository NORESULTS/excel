package com.learn.excel.base;

import org.apache.commons.io.FileUtils;

import java.io.File;

/**
 * @author liangchao
 * @create 2021-04-30 10:27
 */
public class FileBaseUtil {







    /**
     *
     * @param outPathName 文件绝对路径
     * @param bytes 文件字节
     */
    public static void writeByteFile(String outPathName, byte[] bytes) {
        try {
            File fout = new File(outPathName);
            FileUtils.writeByteArrayToFile(fout, bytes);
        } catch (Exception e) {
            System.out.println("writeByteFile"+e);
            throw new RuntimeException("生成excel失败");
        }
    }
}
