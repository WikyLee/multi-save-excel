package com.wikylee;

import org.apache.poi.ss.usermodel.*;

import java.io.*;

/**
 * 读取excel文件
 *
 * @author liweijun
 * @create 2018-05-07 14:16
 **/
public class ReadExcel {
    public void read(final String filePath) {
        File file = new File(filePath);
        if (file.isFile() && file.exists()) {
            System.out.println(filePath + "open successfully.");
        } else {
            System.out.println("Error to open" + filePath);
        }
        InputStream inputStream;
        Workbook workbook;
        try {
            inputStream = new FileInputStream(file);
            workbook = WorkbookFactory.create(inputStream);
            inputStream.close();
            //工作表对象
            Sheet sheet = workbook.getSheetAt(0);
            //总行数
            int rowLength = sheet.getLastRowNum() + 1;
            //工作表的列
            Row row = sheet.getRow(0);
            //总列数
            int colLength = row.getLastCellNum();
            System.out.println("行数：" + rowLength + ",列数：" + colLength);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        new ReadExcel().read("D:/测试表格.xlsx");
    }
}
