package com.learn.excel.base;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

/**
 * @author liangchao
 * @create 2021-04-30 10:33
 */
public class ExcelBaseUtil {




    public static void testCreateXSSFWorkbook() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        workbook.createSheet("sheet01");
        XSSFSheet sheet = workbook.getSheetAt(0);

        XSSFRow row = sheet.createRow(0);

        XSSFCell cell = row.createCell(0);
        cell.setCellValue("0");
        cell.setCellType(XSSFCell.CELL_TYPE_BOOLEAN);

        // 保存文件
        FileOutputStream out = new FileOutputStream("d:\\1111.xlsx"); // 也能用xls，但可能出现格式丢失等问题。
        // 如果存在其他excel格式的同名文件，那么原文件不会被覆盖
        workbook.write(out);
        out.close();
        System.out.println("run ok ");
    }


    public static void testCreateHSSFWrokbook() throws IOException {
        // 创建工作薄
        HSSFWorkbook workbook = new HSSFWorkbook();

        // 创建表
//        HSSFSheet sheet01 = workbook.createSheet("sheet01");
        // 或者
        workbook.createSheet("sheet01");
        HSSFSheet sheet01 = workbook.getSheetAt(0);

        // 创建行
        HSSFRow row = sheet01.createRow(0);
        HSSFCellStyle style = workbook.createCellStyle();
//        style.setAlignment(HSSFCellStyle.CENTER); // 创建一个居中格式   poi 3.14 版本
        style.setAlignment(HorizontalAlignment.CENTER); // 创建一个居中格式 poi 3.17 版本
        // 创建单元格
        HSSFCell cell = row.createCell(0);
        // 使单个单元格具备指定格式
        cell.setCellStyle(style);
        // 设置单元格内容
        cell.setCellValue("0");
        // 多次创建单元格
        row.createCell(1).setCellValue(false);
        row.createCell(2).setCellValue(new Date());
        row.createCell(3).setCellValue(1.111);


        // 保存文件
        FileOutputStream out = new FileOutputStream("d:\\sample.xls");
        workbook.write(out);
        out.close();
        System.out.println("run ok ");

    }

    public static void outputStr() throws IOException {
        String a = "";
        for(int i =0;i<=32767;i++){
            a+="a";
        }
        FileOutputStream out = new FileOutputStream("d:\\sample.txt");
        out.write(a.getBytes());
        out.close();
        System.out.println("run ok ");
    }


    public static void main(String[] args) {

        try {
            testCreateXSSFWorkbook();
        } catch (IOException e) {
            e.printStackTrace();
        }


    }
}
