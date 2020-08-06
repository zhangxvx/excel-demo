package com.example;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author zhangxv
 * @date 2020/8/6
 */
public class ExcelUtil {
    public static void main(String[] args) {
        String path = "src/main/resources/example.xlsx";
        // creatExcel();
        readExcel(path);
    }

    public static void creatExcel(String out) {
        //创建工作簿，对应整个xlsx文件
        XSSFWorkbook workbook = new XSSFWorkbook();
        //创建sheet，对应excel的单个sheet
        XSSFSheet sheet = workbook.createSheet("sheet1");
        //创建行，对应excel中的一行
        XSSFRow row = sheet.createRow(0);
        //创建单元格，对应row中的一格
        XSSFCell cell = row.createCell(0);
        //单元格设置值
        cell.setCellValue("cell");

        XSSFCellStyle style = workbook.createCellStyle();
        //居中
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        //border
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        //单元格设置样式
        cell.setCellStyle(style);

        FileOutputStream file = null;
        try {
            file = new FileOutputStream(out);
            workbook.write(file);
            file.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    public static void readExcel(String in) {
        //创建工作簿
        XSSFWorkbook workbook = null;
        try {
            workbook = new XSSFWorkbook(new FileInputStream(in));
            //读取第一个工作表(这里的下标与list一样的，从0开始取，之后的也是如此)
            XSSFSheet sheet = workbook.getSheetAt(0);
            //获取第一行的数据
            XSSFRow row = sheet.getRow(0);
            //获取该行第一个单元格的数据
            XSSFCell cell = row.getCell(0);
            System.out.println("cello对象：" + cell);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void readAndWriteExcel(String in, String out) {
        //创建工作簿
        XSSFWorkbook workbook = null;
        try {
            workbook = new XSSFWorkbook(new FileInputStream(in));
            //读取第一个工作表(这里的下标与list一样的，从0开始取，之后的也是如此)
            XSSFSheet sheet = workbook.getSheetAt(0);
            //获取第一行的数据
            XSSFRow row = sheet.getRow(0);
            //获取该行第一个单元格的数据
            XSSFCell cell = row.getCell(0);
            System.out.println("cello对象：" + cell);
            cell.setCellValue("new");
            FileOutputStream file = new FileOutputStream(out);
            workbook.write(file);
            file.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
