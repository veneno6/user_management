package com.veo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.FileOutputStream;

@RunWith(SpringRunner.class)
@SpringBootTest
public class POISimpleExample {

    //操作低版本的Excel，使用类HSSFWorkBook，文件扩展名为：xxx.xls
    @Test
    public void handlePOILow() throws Exception {
        //创建工作薄
        Workbook workbook = new HSSFWorkbook();

        //创建工作表
        Sheet sheet = workbook.createSheet("POI操作Excel");

        //创建行
        Row row = sheet.createRow(0);

        //创建单元格
        Cell cell = row.createCell(0);
        //写入单元格内容
        cell.setCellValue("使用POI操作的一个Excel表格案例");

        workbook.write(new FileOutputStream("Test.xls"));
        workbook.close();
    }


    //操作高版本的Excel，使用类XSSFWorkBook，文件扩展名为：xxx.xlsx
    @Test
    public void handlePOIHigh() throws Exception {
        //创建工作薄
        Workbook workbook = new XSSFWorkbook();

        //创建工作表
        Sheet sheet = workbook.createSheet("POI操作Excel");

        //创建行
        Row row = sheet.createRow(0);

        //创建单元格
        Cell cell = row.createCell(0);
        //写入单元格内容
        cell.setCellValue("使用POI操作的一个Excel表格案例");

        workbook.write(new FileOutputStream("Test.xlsx"));
        workbook.close();
    }
}
