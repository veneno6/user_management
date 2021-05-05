package com.veo;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.util.List;

public class WordDemo {
    public static void main(String[] args) throws Exception {
        //创建一个XWPFDocument对象
        XWPFDocument document = new XWPFDocument(new FileInputStream("test.docx"));

        //获取文档中的所有段落，遍历段落获取段落中的文本，在遍历获取最小单元
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (XWPFParagraph paragraph : paragraphs) {
            //打印段落的内容
            System.out.println(paragraph.getText());
            //获取段落中的文本
            List<XWPFRun> runs = paragraph.getRuns();
            for (XWPFRun run : runs) {
                //读取文本
                System.out.println(run.getText(0));
            }
        }

        //读取表格,获取文档中的第一张表格，遍历表格获取表格中的段落在遍历获取文本和最小单元
        XWPFTable table = document.getTables().get(0);
        //获取表格中的行
        List<XWPFTableRow> rows = table.getRows();
        for (XWPFTableRow row : rows) {
            //获取行中的所有单元格，即所有列
            List<XWPFTableCell> tableCells = row.getTableCells();
            for (XWPFTableCell tableCell : tableCells) {
                //获取每个单元格中的段落
                List<XWPFParagraph> paragraphs1 = tableCell.getParagraphs();
                for (XWPFParagraph xwpfParagraph : paragraphs1) {
                    //打印段落的内容
                    System.out.println(xwpfParagraph.getText());
                }
            }
        }
    }
}
