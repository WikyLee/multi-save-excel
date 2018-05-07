package com.wikylee;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * 写入数据
 *
 * @author liweijun
 * @create 2018-05-07 15:32
 **/
public class WriteExcel {

    public void write(final String filePath, List values, int num) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("0");
        for (int i = 0; i < values.size(); i++) {
            Row row = sheet.createRow(i);
            row.createCell(0).setCellValue(values.get(i).toString());
        }
        workbook.setSheetName(0, "第" + num + "列");
        try {
            File file = new File(filePath);
            FileOutputStream fileoutputStream = new FileOutputStream(file);
            workbook.write(fileoutputStream);
            fileoutputStream.close();
            System.out.println("第" + num + "列写入完成，文件路径：" + filePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        List list = new ArrayList();
        list.add(1);
        list.add(2);
        list.add(3);
        list.add(4);
        list.add(5);
        int num = 1;
        new WriteExcel().write("D:/test.xlsx", list, num);
    }
}
