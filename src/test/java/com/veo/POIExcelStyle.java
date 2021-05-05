package com.veo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import java.io.FileOutputStream;

public class POIExcelStyle {

    //1.边框线 2.合并单元格 3.行高列宽 4. 对齐方式 5.字体
    @Test
    public void handleExcelStyle() throws Exception {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("带样式的数据");
        //设置表格的列宽
        sheet.setColumnWidth(0,5*256);
        sheet.setColumnWidth(1,8*256);
        sheet.setColumnWidth(2,15*256);
        sheet.setColumnWidth(3,15*256);
        sheet.setColumnWidth(4,30*256);
        //需求：1.边框线：全边框 2.合并单元格：第一行的第1个单元格到第5个 3.行高：42 4. 对齐方式：水平垂直居中 5.字体:黑体，18
        //使用workbook创建样式，可以用于当前的workbook
        CellStyle bitTitleRowStyle = workbook.createCellStyle();
        //设置样式的上、下、左、右，全边框线，参数：边框线的类型
        bitTitleRowStyle.setBorderTop(BorderStyle.THIN);
        bitTitleRowStyle.setBorderBottom(BorderStyle.THIN);
        bitTitleRowStyle.setBorderLeft(BorderStyle.THIN);
        bitTitleRowStyle.setBorderRight(BorderStyle.THIN);
        //设置对齐的方式：水平、垂直居中对齐,参数：对齐方式
        bitTitleRowStyle.setAlignment(HorizontalAlignment.CENTER);
        bitTitleRowStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        //使用workbook创建字体，可应用于当前的workbook
        Font font = workbook.createFont();
        //字体样式为：黑体 18号
        font.setFontName("黑体");
        font.setFontHeightInPoints((short) 18);
        //应用样式，到标题行
        bitTitleRowStyle.setFont(font);

        Row bitTitleRow = sheet.createRow(0);
        for (int i = 0; i <=4 ; i++) {
            //创建单元格
            Cell cell = bitTitleRow.createCell(i);
            //为单元格应用样式
            cell.setCellStyle(bitTitleRowStyle);
        }
        //合并单元格：int firstRow 起始行, int lastRow 结束行, int firstCol 起始列, int lastCol 结束列
        sheet.addMergedRegion(new CellRangeAddress(0,0,0,4));
        //向创建好的单元格中放入数据
        sheet.getRow(0).getCell(0).setCellValue("用户数据");

        //小标题样式
        CellStyle littleTitleStyle = workbook.createCellStyle();
        //克隆大标题的基本样式，减少代码的冗余
        littleTitleStyle.cloneStyleFrom(bitTitleRowStyle);
        //设置字体：宋体 12号 加粗
        Font littleFont = workbook.createFont();
        littleFont.setFontName("宋体");
        littleFont.setFontHeightInPoints((short) 12);
        littleFont.setBold(true);
        littleTitleStyle.setFont(littleFont);

        //内容样式
        CellStyle contentStyle = workbook.createCellStyle();
        //克隆大标题的基本样式，减少代码的冗余
        contentStyle.cloneStyleFrom(bitTitleRowStyle);
        //水平居左
        contentStyle.setAlignment(HorizontalAlignment.LEFT);
        //设置字体：宋体 11号
        Font contentFont = workbook.createFont();
        contentFont.setFontName("宋体");
        contentFont.setFontHeightInPoints((short) 11);
        contentFont.setBold(false);
        contentStyle.setFont(contentFont);

        //处理固定的标题
        Row titleRow = sheet.createRow(1);
        Cell cell = null;
        String[] title = {"编号","姓名","手机号","入职日期","现住址"};
        for (int i = 0; i < title.length; i++) {
            cell = titleRow.createCell(i);
            cell.setCellValue(title[i]);
            cell.setCellStyle(littleTitleStyle);
        }

        //模拟表格内容
        Row contentRow = sheet.createRow(2);
        String[] content = {"1","大一","13800000001","2001-03-29","北京市西城区宣武大街1号院"};
        for (int i = 0; i < content.length; i++) {
            cell = contentRow.createCell(i);
            cell.setCellValue(content[i]);
            cell.setCellStyle(contentStyle);
        }

        //导出文件
        workbook.write(new FileOutputStream("testStyle.xlsx"));
    }
}
